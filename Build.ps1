param([switch]$Test,[switch]$UI,[switch]$Conversion)
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path -LiteralPath $compiler)) { $compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
$framework = Split-Path -Path $compiler -Parent
$wpf = Join-Path $framework 'WPF'
$output = Join-Path $root '发布\Word_Compare.exe'
New-Item -ItemType Directory -Path (Split-Path -Path $output -Parent) -Force | Out-Null
$refs = @('System.dll','System.Core.dll','System.Xml.dll','System.Xml.Linq.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Xaml.dll','System.Web.Extensions.dll') | ForEach-Object { '/reference:' + (Join-Path $framework $_) }
$refs += @('PresentationFramework.dll','PresentationCore.dll','WindowsBase.dll') | ForEach-Object { '/reference:' + (Join-Path $wpf $_) }
$sources = Get-ChildItem -LiteralPath (Join-Path $root 'src') -Filter '*.cs' | Select-Object -ExpandProperty FullName
& $compiler /nologo /target:winexe /platform:anycpu /optimize+ /codepage:65001 "/out:$output" "/win32icon:$root\assets\Word_Compare.ico" "/win32manifest:$root\assets\app.manifest" "/resource:$root\assets\Word_Compare.ico,Word_Compare.ico" @refs @sources
if ($LASTEXITCODE -ne 0) { throw 'Word_Compare 编译失败。' }
if ($Test -or $UI -or $Conversion) {
    $cases = @()
    if ($Test) { $cases += 'EngineTests','FeatureTests' }
    if ($UI) { $cases += 'AppearanceTests' }
    if ($Conversion) { $cases += 'ConversionTests' }
    Copy-Item -LiteralPath $output -Destination (Join-Path $root 'tests\Word_Compare.exe') -Force
    try {
        foreach ($case in $cases) {
            $testExe = Join-Path $root ('tests\' + $case + '.exe')
            & $compiler /nologo /target:exe /codepage:65001 "/out:$testExe" "/reference:$output" @refs (Join-Path $root ('tests\' + $case + '.cs'))
            if ($LASTEXITCODE -ne 0) { throw ($case + ' 编译失败。') }
            if ($case -eq 'EngineTests') { & $testExe (Join-Path $root '示例') (Join-Path $root '验证结果') }
            else { & $testExe $root (Join-Path $root '验证结果') }
            if ($LASTEXITCODE -ne 0) { throw ($case + ' 验证失败。') }
            Remove-Item -LiteralPath $testExe -Force
        }
    } finally {
        foreach ($case in $cases) {
            $testExe = Join-Path $root ('tests\' + $case + '.exe')
            Remove-Item -LiteralPath $testExe -Force -ErrorAction SilentlyContinue
        }
        Remove-Item -LiteralPath (Join-Path $root 'tests\Word_Compare.exe') -Force -ErrorAction SilentlyContinue
    }
}
Get-Item -LiteralPath $output | Select-Object FullName,Length
