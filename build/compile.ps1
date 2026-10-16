param([string]$Src, [string]$Out, [string]$Ico, [string]$RefDir)
Add-Type -AssemblyName Microsoft.CodeAnalysis, Microsoft.CodeAnalysis.CSharp
$code = [IO.File]::ReadAllText($Src, [Text.Encoding]::UTF8)
$opts = [Microsoft.CodeAnalysis.CSharp.CSharpParseOptions]::new([Microsoft.CodeAnalysis.CSharp.LanguageVersion]::CSharp7_3)
$tree = [Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree]::ParseText($code, $opts, $Src, [Text.Encoding]::UTF8)
$refs = 'mscorlib','System','System.Core','System.Drawing','System.Windows.Forms','System.Xml','System.IO.Compression' |
  % { [Microsoft.CodeAnalysis.MetadataReference]::CreateFromFile((Join-Path $RefDir "$_.dll")) }
$copts = [Microsoft.CodeAnalysis.CSharp.CSharpCompilationOptions]::new([Microsoft.CodeAnalysis.OutputKind]::WindowsApplication).
  WithOptimizationLevel([Microsoft.CodeAnalysis.OptimizationLevel]::Release).
  WithPlatform([Microsoft.CodeAnalysis.Platform]::AnyCpu)
$comp = [Microsoft.CodeAnalysis.CSharp.CSharpCompilation]::Create('업무메모', [Microsoft.CodeAnalysis.SyntaxTree[]]@($tree), [Microsoft.CodeAnalysis.MetadataReference[]]$refs, $copts)
$icoStream = [IO.File]::OpenRead($Ico)
$res = $comp.CreateDefaultWin32Resources($true, $false, $null, $icoStream)
$pe = [IO.File]::Create($Out)
$r = $comp.Emit($pe, $null, $null, $res)
$pe.Dispose(); $icoStream.Dispose()
$r.Diagnostics | ? { $_.Severity -ne 'Hidden' } | % { $_.ToString() }
"success: " + $r.Success
if (-not $r.Success) { Remove-Item $Out }
