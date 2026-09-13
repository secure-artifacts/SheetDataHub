$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

python -m unittest discover -s tests -v
python -m PyInstaller --noconfirm --clean --distpath dist-onedir SheetDataHub.spec

$isccCandidates = @(
    "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
    "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if ($iscc) {
    & $iscc installer.iss
    Write-Host "Installer created in: $projectRoot\dist-installer"
} else {
    Write-Host 'Inno Setup 6 not found; building the bundled installer.'
    python -m PyInstaller --noconfirm --clean --distpath dist-installer --workpath build-installer SheetDataHubInstaller.spec
    Write-Host "Installer created in: $projectRoot\dist-installer"
}
