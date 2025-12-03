###############################################################
# ROOT + BASE PERMISSIONS + SMB SHARE + FULL FOLDER TREE
# No accents, no apostrophes
###############################################################

# Root folder path
$Root = "C:\Shares"

# Share name
$ShareName = "Shares"

Write-Host "Creating root folder..." -ForegroundColor Cyan

# Create root if needed
if (-not (Test-Path $Root)) {
    New-Item -ItemType Directory -Path $Root | Out-Null
    Write-Host "Created root folder: $Root"
} else {
    Write-Host "Root already exists: $Root"
}

###############################################################
# SET DEFAULT NTFS PERMISSIONS ON ROOT (DOMAIN CONTROLLER)
###############################################################

Write-Host "Setting default NTFS permissions..." -ForegroundColor Cyan

icacls $Root /inheritance:d | Out-Null

Write-Host "Base NTFS permissions applied." -ForegroundColor Green


###############################################################
# ENABLE SMB SHARING (DOMAIN CONTROLLER COMPATIBLE)
###############################################################

Write-Host "Configuring SMB share..." -ForegroundColor Cyan

if (Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue) {
    Remove-SmbShare -Name $ShareName -Force
}

New-SmbShare -Name $ShareName -Path $Root

Write-Host "Share available at: \\$env:COMPUTERNAME\$ShareName" -ForegroundColor Green



###############################################################
# DIRECTORY TREE BASED ON ORGANIGRAM
###############################################################

Write-Host "Creating folder tree..." -ForegroundColor Cyan

$Folders = @(
    # Direction
    "$Root\Direction",

    # Ressources Humaines
    "$Root\RessourcesHumaines",
    "$Root\RessourcesHumaines\GestionDuPersonnel",
    "$Root\RessourcesHumaines\Recrutement",

    # R&D
    "$Root\RD",
    "$Root\RD\Recherche",
    "$Root\RD\Testing",

    # Marketing
    "$Root\Marketing",
    "$Root\Marketing\Site1",
    "$Root\Marketing\Site2",
    "$Root\Marketing\Site3",
    "$Root\Marketing\Site4",

    # Finances
    "$Root\Finances",
    "$Root\Finances\Comptabilite",
    "$Root\Finances\Investissements",

    # Technique
    "$Root\Technique",
    "$Root\Technique\Techniciens",
    "$Root\Technique\Achat",

    # Informatique
    "$Root\Informatique",
    "$Root\Informatique\Systemes",
    "$Root\Informatique\Developpement",
    "$Root\Informatique\HotLine",

    # Commerciaux
    "$Root\Commerciaux",
    "$Root\Commerciaux\Sedentaires",
    "$Root\Commerciaux\Technico",

    # Commun
    "$Root\Commun"
)

foreach ($folder in $Folders) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder | Out-Null
        Write-Host "Created: $folder"
    } else {
        Write-Host "Already exists: $folder"
    }
}

Write-Host "Folder tree created successfully." -ForegroundColor Yellow
###############################################################
