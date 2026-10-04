# One step after any change: updates dist\WordPublisherKit.dotm from ribbon\ and src\,
# checks that the code compiles, and installs it into Word's STARTUP folder.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File build\Update-AddIn.ps1
#
# One-time Word setting needed on this PC: File > Options > Trust Center >
# Trust Center Settings > Macro Settings > tick "Trust access to the VBA project
# object model".
param([switch]$NoInstall)
$ErrorActionPreference = 'Stop'
$kit      = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$template = Join-Path $kit 'dist\WordPublisherKit.dotm'
$ribbon   = Join-Path $kit 'ribbon\customUI14.xml'
$srcDir   = Join-Path $kit 'src'
$startup  = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Microsoft Word is open. Save your work, close Word, and run this again.'
}
if (-not (Test-Path $template)) { throw "Missing $template" }

# Unblock (downloaded files can't run macros) and keep a backup of the last good build.
Unblock-File -LiteralPath $template
Copy-Item $template "$template.bak" -Force

# 1. Ribbon: replace customUI/customUI14.xml inside the .dotm (it is a zip file).
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($template, 'Update')
try {
    $old = $zip.GetEntry('customUI/customUI14.xml')
    if ($old) { $old.Delete() }
    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $ribbon, 'customUI/customUI14.xml')
} finally { $zip.Dispose() }
Write-Host 'Ribbon updated.'

# 2. Code: re-import every module, then make sure the project compiles.
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($template)
    try { $vbp = $doc.VBProject; $null = $vbp.VBComponents.Count }
    catch { throw 'Word blocked access to the VBA project. Turn on "Trust access to the VBA project object model" (see the top of this script).' }

    foreach ($file in Get-ChildItem $srcDir -Filter *.bas) {
        $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        foreach ($c in @($vbp.VBComponents)) { if ($c.Name -eq $name) { $vbp.VBComponents.Remove($c) } }
        $null = $vbp.VBComponents.Import($file.FullName)
        Write-Host "Imported $name"
    }
    $doc.Save()

    try { $null = $word.Run('PubVersion') }
    catch {
        $doc.Close(0)
        Copy-Item "$template.bak" $template -Force
        throw "The code does not compile, so the previous version was kept. Word said: $($_.Exception.Message)"
    }
    $doc.Close()
} finally {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
Write-Host 'Code compiled.'

# 3. Install for this Windows user.
if (-not $NoInstall) {
    New-Item -ItemType Directory -Force -Path $startup | Out-Null
    Copy-Item $template (Join-Path $startup 'WordPublisherKit.dotm') -Force
    Write-Host "Installed. Open Word to use the updated Bulletin tab."
}
