###############################################################
# SHARED FOLDER + SUBFOLDERS + NTFS PERMISSIONS + QUOTAS
# Version A - Simple
# No accents, no apostrophes
###############################################################

Import-Module FileServerResourceManager -ErrorAction SilentlyContinue

###############################################################
# ROOT CONFIG
###############################################################

$Root      = "C:\Shares"
$ShareName = "Shares"
$LocalAdmins = "$env:COMPUTERNAME\Administrateurs"

Write-Host "Creating root folder..." -ForegroundColor Cyan
if (-not (Test-Path $Root)) { New-Item -ItemType Directory -Path $Root | Out-Null }

###############################################################
# NTFS ROOT PERMISSIONS
###############################################################

Write-Host "Applying root NTFS permissions..."

icacls $Root /inheritance:d | Out-Null
icacls $Root /grant:r "*S-1-5-18:(OI)(CI)(F)" | Out-Null                # SYSTEM
icacls $Root /grant:r "$LocalAdmins:(OI)(CI)(F)" | Out-Null            # Local admins

###############################################################
# SMB SHARE
###############################################################

Write-Host "Creating SMB share..."

if (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue) {
    Remove-SmbShare -Name $ShareName -Force
}

New-SmbShare -Name $ShareName -Path $Root -FullAccess "$LocalAdmins" | Out-Null

Write-Host "Share available at: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor Green


###############################################################
# FOLDER STRUCTURE
###############################################################

$Folders = @(
    "$Root\Direction",
    "$Root\RessourcesHumaines",
    "$Root\RessourcesHumaines\GestionDuPersonnel",
    "$Root\RessourcesHumaines\Recrutement",
    "$Root\RD",
    "$Root\RD\Recherche",
    "$Root\RD\Testing",
    "$Root\Marketing",
    "$Root\Marketing\Site1",
    "$Root\Marketing\Site2",
    "$Root\Marketing\Site3",
    "$Root\Marketing\Site4",
    "$Root\Finances",
    "$Root\Finances\Comptabilite",
    "$Root\Finances\Investissements",
    "$Root\Technique",
    "$Root\Technique\Techniciens",
    "$Root\Technique\Achat",
    "$Root\Informatique",
    "$Root\Informatique\Systemes",
    "$Root\Informatique\Developpement",
    "$Root\Informatique\HotLine",
    "$Root\Commerciaux",
    "$Root\Commerciaux\Sedentaires",
    "$Root\Commerciaux\Technico",
    "$Root\Commun"
)

Write-Host "Creating folder tree..." -ForegroundColor Cyan

foreach ($Folder in $Folders) {
    if (-not (Test-Path $Folder)) {
        New-Item -ItemType Directory -Path $Folder | Out-Null
        Write-Host "Created: $Folder"
    }
}


###############################################################
# NTFS PERMISSIONS BASED ON GL GROUPS
###############################################################

Write-Host "Applying NTFS permissions for GL groups..." -ForegroundColor Cyan

$NTFSMap = @(
    @{ Path="$Root\Direction";              RW="GL_Direction_RW";         R="GL_Direction_R" }
    @{ Path="$Root\RessourcesHumaines";     RW="GL_RH_RW";                R="GL_RH_R" }
    @{ Path="$Root\RessourcesHumaines\GestionDuPersonnel"; RW="GL_RH_Gestion_RW"; R="GL_RH_Gestion_R" }
    @{ Path="$Root\RessourcesHumaines\Recrutement";         RW="GL_RH_Recrutement_RW"; R="GL_RH_Recrutement_R" }
    @{ Path="$Root\RD";                     RW="GL_RD_RW";                R="GL_RD_R" }
    @{ Path="$Root\RD\Recherche";           RW="GL_RD_Recherche_RW";      R="GL_RD_Recherche_R" }
    @{ Path="$Root\RD\Testing";             RW="GL_RD_Testing_RW";        R="GL_RD_Testing_R" }
    @{ Path="$Root\Commun";                 RW="GL_Commun_RW";            R="GL_Commun_R" }
)

foreach ($item in $NTFSMap) {

    $folder = $item.Path
    $rw     = $item.RW
    $r      = $item.R

    if (-not (Test-Path $folder)) { continue }

    icacls $folder /inheritance:d | Out-Null
    icacls $folder /grant:r "$rw:(OI)(CI)(M)" | Out-Null
    icacls $folder /grant:r "$r:(OI)(CI)(RX)" | Out-Null
    icacls $folder /grant:r "*S-1-5-18:(OI)(CI)(F)" | Out-Null
    icacls $folder /grant:r "$LocalAdmins:(OI)(CI)(F)" | Out-Null

    Write-Host "Applied GL permissions on: $folder"
}


###############################################################
# QUOTAS
###############################################################

Write-Host "Applying quotas..." -ForegroundColor Cyan

$QuotaMap = @(
    @{ Path="$Root\Direction";              Template="Dept_500MB" }
    @{ Path="$Root\RessourcesHumaines";     Template="Dept_500MB" }
    @{ Path="$Root\RD";                     Template="Dept_500MB" }
    @{ Path="$Root\Marketing";              Template="Dept_500MB" }
    @{ Path="$Root\Finances";               Template="Dept_500MB" }
    @{ Path="$Root\Technique";              Template="Dept_500MB" }
    @{ Path="$Root\Informatique";           Template="Dept_500MB" }
    @{ Path="$Root\Commerciaux";            Template="Dept_500MB" }
    @{ Path="$Root\Commun";                 Template="Commun_500MB" }

    @{ Path="$Root\RessourcesHumaines\GestionDuPersonnel"; Template="SubDept_100MB" }
    @{ Path="$Root\RessourcesHumaines\Recrutement";         Template="SubDept_100MB" }
    @{ Path="$Root\RD\Recherche";                           Template="SubDept_100MB" }
    @{ Path="$Root\RD\Testing";                             Template="SubDept_100MB" }
)

foreach ($q in $QuotaMap) {
    if (Test-Path $q.Path) {
        New-FsrmQuota -Path $q.Path -Template $q.Template -ErrorAction SilentlyContinue | Out-Null
        Write-Host "Quota applied on: $q.Path"
    }
}

Write-Host "DONE." -ForegroundColor Yellow
###############################################################
