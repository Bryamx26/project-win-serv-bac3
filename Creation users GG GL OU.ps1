# ==============================================================================
# SCRIPT MAÎTRE DE PROVISIONING ACTIVE DIRECTORY V15
# Amélioration : Séparation de la création User et de l'ajout Groupe.
# Si l'utilisateur existe déjà, le script tente quand même l'ajout au groupe.
# ==============================================================================
Import-Module ActiveDirectory

# ==============================================================================
# 1. CONFIGURATION GLOBALE
# ==============================================================================
$CSVPath = ".\EmployesL2.csv"

# *** DÉLIMITEUR ***
# Mettez "," si votre fichier est séparé par des virgules (comme votre exemple texte)
# Mettez ";" si c'est un CSV Excel standard français.
$CSVDelimiter = "," 

$ServerDC = "ADAG4.agence4.local" 
$BaseDN = "DC=agence4,DC=local"
$ADDomain = "agence4.local" 
$MaxLenSamAccountName = 13

# ------------------------------------------------------------
# 2. TABLES DE CORRESPONDANCE
# ------------------------------------------------------------
$DepartmentToOU = @{
    "Site1/Marketing"                   = "Direction/Marketing/Site 1"
    "Site2/Marketing"                   = "Direction/Marketing/Site 2"
    "Site3/Marketing"                   = "Direction/Marketing/Site 3"
    "Site4/Marketing"                   = "Direction/Marketing/Site 4"
    "Comptabilite/Finances"             = "Direction/Finances/Comptabilite"
    "Investissements/Finances"          = "Direction/Finances/Investissements"
    "Recherche/R&D"                     = "Direction/R&D/Recherche"
    "Testing/R&D"                       = "Direction/R&D/Testing"
    "Techniciens/Technique"             = "Direction/Technique/Techniciens"
    "Achat/Technique"                   = "Direction/Technique/Achat"
    "Systemes/Informatique"             = "Direction/Informatique/Systemes"
    "Developpement/Informatique"        = "Direction/Informatique/Developpement"
    "HotLine/Informatique"              = "Direction/Informatique/HotLine"
    "Sedentaires/Commerciaux"           = "Direction/Commerciaux/Sedentaires"
    "Technico/Commerciaux"              = "Direction/Commerciaux/Technico"
    "Gestion du personnel/Ressources humaines" = "Direction/Ressources humaines/Gestion du personnel"
    "Recrutement/Ressources humaines"          = "Direction/Ressources humaines/Recrutement"
}

# Liste des OUs pour la création de structure
$OUList = @(
    "Direction",
    "Direction/Ressources humaines",
    "Direction/Ressources humaines/Gestion du personnel",
    "Direction/Ressources humaines/Recrutement",
    "Direction/R&D",
    "Direction/R&D/Recherche",
    "Direction/R&D/Testing",
    "Direction/Marketing",
    "Direction/Marketing/Site 1",
    "Direction/Marketing/Site 2",
    "Direction/Marketing/Site 3",
    "Direction/Marketing/Site 4",
    "Direction/Finances",
    "Direction/Finances/Comptabilite",
    "Direction/Finances/Investissements",
    "Direction/Technique",
    "Direction/Technique/Techniciens",
    "Direction/Technique/Achat",
    "Direction/Commerciaux",
    "Direction/Commerciaux/Sedentaires",
    "Direction/Commerciaux/Technico",
    "Direction/Informatique",
    "Direction/Informatique/Systemes",
    "Direction/Informatique/Developpement",
    "Direction/Informatique/HotLine"
)

# ------------------------------------------------------------
# 3. FONCTIONS UTILITAIRES
# ------------------------------------------------------------

function Get-SamAccountName {
    param([string]$Prenom, [string]$Nom)
    if ($Nom.Length + $Prenom.Length -gt $MaxLenSamAccountName) {
        $Sam = "$($Nom[0]).$Prenom"
    } else {
        $Sam = "$Prenom.$Nom"
    }
    return $Sam.ToLower()
}

function New-UserPassword {
    param(
        [Parameter(Mandatory)]
        $Utilisateur
    )

    # Longueur selon le département
    if ($Utilisateur.Departement -eq "Direction") {
        $length = 15
    } else {
        $length = 12
    }

    $upper   = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    $lower   = "abcdefghijklmnopqrstuvwxyz"
    $digits  = "0123456789"
    $special = "!@#$%^&*()-_=+[]{}:,.?/"

    # Complexité minimale
    $mandatory = @(
        $upper[(Get-Random -Maximum $upper.Length)]
        $lower[(Get-Random -Maximum $lower.Length)]
        $digits[(Get-Random -Maximum $digits.Length)]
        $special[(Get-Random -Maximum $special.Length)]
    )

    $allChars = $upper + $lower + $digits + $special
    $remaining = $length - $mandatory.Count

    $randomChars = -join (1..$remaining | ForEach-Object {
        $allChars[(Get-Random -Maximum $allChars.Length)]
    })

    $chars = ((($mandatory -join "") + $randomChars).ToCharArray() | Sort-Object { Get-Random })
    $password = -join $chars

    return (ConvertTo-SecureString $password -AsPlainText -Force)
}

function Fix-CSVHeaders {
    param([string]$CSVPath)
    if (-not (Test-Path $CSVPath)) { return }

    # Lecture en Default (ANSI) pour les accents
    $lines = Get-Content $CSVPath -Encoding Default 
    if ($lines.Count -lt 1) { return }

    $newHeader = "Nom,Prenom,Description,Departement,Telephone,Bureau"
    
    # Détection virgule ou point-virgule
    global:CSVDelimiter = "," 
    if ($lines[0] -match ";") { 
        $newHeader = $newHeader -replace ",", ";" 
        $global:CSVDelimiter = ";"
        Write-Host "Détection : Le fichier utilise des points-virgules (;)." -ForegroundColor Cyan
    }

    $lines[0] = $newHeader
    $lines = $lines | ForEach-Object { $_ -replace "Marketting", "Marketing" }
    
    Set-Content -Path $CSVPath -Value $lines -Encoding Default
    Write-Host "✔ En-têtes normalisés." -ForegroundColor Green
}

# ------------------------------------------------------------
# 4. FONCTIONS ACTIVE DIRECTORY
# ------------------------------------------------------------

function Convert-PathToDN($Path, $Domain) {
    $Segments = $Path -split "/" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    [array]::Reverse($Segments)
    $OUString = ($Segments | ForEach-Object { "OU=$_" }) -join ","
    $DCString = ($Domain -split "\." | ForEach-Object { "DC=$_" }) -join ","
    return "$OUString,$DCString"
}

function Create-OUHierarchy {
    param($BaseDN, $OUPath, $ServerDC)
    $levels = $OUPath -split "/"
    $currentDN = $BaseDN

    foreach ($ou in $levels) {
        if ([string]::IsNullOrWhiteSpace($ou)) { continue }
        $ouName = $ou.Trim()
        $ouDN = "OU=$ouName,$currentDN"

        # Création OU
        if (-not (Get-ADOrganizationalUnit -Filter "Name -eq '$ouName'" -SearchBase $currentDN -ErrorAction SilentlyContinue)) {
            New-ADOrganizationalUnit -Name $ouName -Path $currentDN -Server $ServerDC
            Write-Host "OU créée : $ouDN" -ForegroundColor Cyan
        }

        # Création GG
        $ggName = "$ouName-GG"
        if (-not (Get-ADGroup -Filter "Name -eq '$ggName'" -SearchBase $ouDN -ErrorAction SilentlyContinue)) {
            New-ADGroup -Name $ggName -SamAccountName $ggName -GroupScope Global -GroupCategory Security -Path $ouDN -Server $ServerDC
            Write-Host "GG créé : $ggName" -ForegroundColor Cyan
        }

        # Création GL et Imbrication
        foreach ($suf in @("_R", "_RW")) {
            $glName = "GL_$ouName$suf"
            if (-not (Get-ADGroup -Filter "Name -eq '$glName'" -SearchBase $ouDN -ErrorAction SilentlyContinue)) {
                New-ADGroup -Name $glName -SamAccountName $glName -GroupScope DomainLocal -GroupCategory Security -Path $ouDN -Server $ServerDC
                Write-Host "GL créé : $glName" -ForegroundColor Cyan
            }
            try {
                Add-ADGroupMember -Identity $glName -Members $ggName -Server $ServerDC -ErrorAction Stop
            } catch {} 
        }
        $currentDN = $ouDN
    }
}

# ------------------------------------------------------------
# 4. FONCTIONS Sauvgade des credentials
# ------------------------------------------------------------


function Save-Credentials {
    param (
        [string]$SamAccountName,
        [System.Security.SecureString]$SecurePassword,
        [string]$NomComplet
    )

    # 1. Définir le dossier à la racine (C:\Credentials)
    $FolderPath = "C:\Credentials"
    $FilePath = "$FolderPath\Liste_MotsDePasse.csv"

    # 2. Créer le dossier s'il n'existe pas
    if (-not (Test-Path $FolderPath)) {
        New-Item -Path $FolderPath -ItemType Directory -Force | Out-Null
        Write-Host " -> Dossier '$FolderPath' créé." -ForegroundColor Cyan
    }

    # 3. Convertir le SecureString en Texte Clair (Nécessaire pour l'export CSV)
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
    $PlainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    # Nettoyage mémoire immédiat
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)

    # 4. Créer l'objet à exporter
    $InfoUtilisateur = [PSCustomObject]@{
        DateCreation   = (Get-Date).ToString("dd/MM/yyyy HH:mm")
        NomComplet     = $NomComplet
        Login          = $SamAccountName
        MotDePasse     = $PlainTextPassword
    }

    # 5. Exporter dans le CSV (Ajout en fin de fichier - Append)
    # On utilise le point-virgule comme vous le préférez
    $InfoUtilisateur | Export-Csv -Path $FilePath -Append -NoTypeInformation -Delimiter ";" -Encoding UTF8
    
    Write-Host " -> Identifiants sauvegardés dans $FilePath" -ForegroundColor Green
}

# ------------------------------------------------------------
# 5. EXÉCUTION
# ------------------------------------------------------------

# 1. Correction En-tête
Fix-CSVHeaders -CSVPath $CSVPath

# 2. Structure AD (Assure que les GGs existent avant d'ajouter les users)
Write-Host "`n--- Vérification Structure AD ---"
foreach ($ou in $OUList) { Create-OUHierarchy -BaseDN $BaseDN -OUPath $ou -ServerDC $ServerDC }

# 3. Import et Nettoyage
Write-Host "`n--- Importation Utilisateurs ---"
if (-not (Test-Path $CSVPath)) { Write-Error "Fichier introuvable"; exit }

$tab = Import-Csv $CSVPath -Delimiter $CSVDelimiter -Encoding Default
if (-not $tab) { Write-Error "CSV vide ou illisible."; exit }

# Nettoyage en mémoire (Accents -> sans accents)
Write-Host "Nettoyage des accents en mémoire..." -ForegroundColor Yellow
foreach ($row in $tab) {
    foreach ($prop in $row.PSObject.Properties) {
        if ($prop.Value -is [string]) {
            $v = $prop.Value
            $v = $v -replace "é", "e" -replace "è", "e" -replace "ê", "e" -replace "ë", "e"
            $v = $v -replace "à", "a" -replace "â", "a" -replace "ä", "a"
            $v = $v -replace "î", "i" -replace "ï", "i"
            $v = $v -replace "ô", "o" -replace "ö", "o"
            $v = $v -replace "ù", "u" -replace "û", "u" -replace "ü", "u"
            $v = $v -replace "ç", "c" -replace "'", ""
            $prop.Value = $v.Trim()
        }
    }
}
Write-Host "✔ Nettoyage terminé." -ForegroundColor Green

# 4. Traitement des utilisateurs (Création + Ajout Groupe)
foreach ($user in $tab) {
    if ([string]::IsNullOrWhiteSpace($user.Nom)) { continue }

    $Nom = $user.Nom
    $Prenom = $user.Prenom
    $Departement = $user.Departement
    $Tel = $user.Telephone

    if (-not $DepartmentToOU.ContainsKey($Departement)) {
        Write-Warning "Département inconnu: $Departement (Utilisateur: $Nom)"
        continue
    }

    $OUPathSimple = $DepartmentToOU[$Departement]
    $OUFullDN = Convert-PathToDN $OUPathSimple $ADDomain
    $Sam = Get-SamAccountName $Prenom $Nom
    $UPN = "$Sam@$ADDomain"
    $Pass = New-UserPassword $user
    


    # Calcul du nom du GG (ex: Site 1 -> Site 1-GG)
    $GGName = "$($OUPathSimple.Split('/')[-1].Trim())-GG"

    Write-Host "`nTraitement: $Prenom $Nom ($Sam)" -ForegroundColor Yellow

    # --- BLOC 1 : CRÉATION UTILISATEUR ---
    try {
        New-ADUser -Name "$Prenom $Nom" -GivenName $Prenom -Surname $Nom -SamAccountName $Sam `
            -UserPrincipalName $UPN -Description $user.Description -OfficePhone $Tel `
            -Office $user.Bureau -Enabled $true -Path $OUFullDN -AccountPassword $Pass `
            -ChangePasswordAtLogon $false -Server $ServerDC -ErrorAction Stop  # <--- RETIRER -WhatIf
        
        Write-Host " -> Compte Créé " -ForegroundColor Green
        Save-Credentials -SamAccountName $Sam -SecurePassword $Pass -NomComplet "$Prenom $Nom"
    }
    catch {
        if ($_.Exception.Message -like "*already exists*") {
            Write-Warning " -> Le compte existe déjà. Passage à l'ajout de groupe."
        }
        else {
            Write-Error " -> Erreur Création: $($_.Exception.Message)"
        }
    }

    # --- BLOC 2 : AJOUT AU GROUPE GLOBAL (GG) ---
    try {
        # On tente l'ajout MEME SI l'utilisateur existait déjà
        Add-ADGroupMember -Identity $GGName -Members $Sam -Server $ServerDC -ErrorAction Stop  # <--- RETIRER -WhatIf
        
        Write-Host " -> Ajouté au groupe $GGName (Simulé)" -ForegroundColor Cyan
    }
    catch {
        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
            Write-Warning " -> Déjà membre du groupe $GGName."
        }
        elseif ($_.Exception.Message -like "*Cannot find an object with identity*") {
             Write-Warning " -> Impossible d'ajouter au groupe (Utilisateur introuvable en mode simulation)."
        }
        else {
            Write-Error " -> Erreur Groupe: $($_.Exception.Message)"
        }
    }
}

Write-Host "`nFIN. Pour exécuter réellement, retirez les '-WhatIf' aux lignes 230 et 243."