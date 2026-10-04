# Adds the VBA modules in ..\src to ..\dist\WordPublisherKit.dotm (run on Windows, once,
# after build_dotm.py and every time the .bas files change).
#
# Word must allow it first: File > Options > Trust Center > Trust Center Settings >
# Macro Settings > tick "Trust access to the VBA project object model".
# (You can untick it again afterwards.)
$ErrorActionPreference = 'Stop'
$kit      = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$template = Join-Path $kit 'dist\WordPublisherKit.dotm'
$srcDir   = Join-Path $kit 'src'

if (-not (Test-Path $template)) { throw "Missing $template. Run build\build_dotm.py first." }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close Microsoft Word first.' }

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($template)
    try {
        $vbp = $doc.VBProject
        $null = $vbp.VBComponents.Count
    } catch {
        throw 'Word blocked access to the VBA project. Turn on "Trust access to the VBA project object model" (see the top of this script).'
    }
    foreach ($file in Get-ChildItem $srcDir -Filter *.bas) {
        $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
        foreach ($c in @($vbp.VBComponents)) {
            if ($c.Name -eq $name) { $vbp.VBComponents.Remove($c) }
        }
        $null = $vbp.VBComponents.Import($file.FullName)
        Write-Host "Imported $name"
    }
    $doc.Save()
    $doc.Close()
    Write-Host "Done: $template"
} finally {
    $word.Quit()
}
