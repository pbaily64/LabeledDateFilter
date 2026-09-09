# del_move.ps1 — renomme le .pbiviz fraîchement généré en labeledDateFilter.pbiviz

# Le script s'exécute depuis le dossier où il se trouve (dist/)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location $scriptDir

$target = "labeledDateFilter.pbiviz"
$prefix = "labeledDateFilterD39304CD1980483DB4A67FF5C3395A2D"

# Cherche tous les .pbiviz commençant par le prefix (peu importe la version)
$candidates = Get-ChildItem -Path $scriptDir -Filter "$prefix*.pbiviz" -File

if ($candidates.Count -eq 0) {
    Write-Error "No file '$prefix*.pbiviz' found in $scriptDir"
    exit 1
}

# Prend le plus récent (au cas où il y aurait plusieurs versions)
$source = ($candidates | Sort-Object LastWriteTime -Descending | Select-Object -First 1).Name

# Supprime l'ancien fichier cible s'il existe
if (Test-Path $target) {
    Remove-Item $target -Force
    Write-Host "Deleted : $target"
}

# Renomme le source en target
Rename-Item -Path $source -NewName $target
Write-Host "Renamed : $source -> $target"
