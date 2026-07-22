param(
  [string]$Source = 'D:\workprj\OfficeRead_new',
  [string]$Cpp = (Join-Path $PSScriptRoot '..\build\officeread.exe'),
  [int]$PerFormat = 10
)
$ErrorActionPreference = 'Stop'
$formats = '.docx','.pptx','.xlsx','.doc','.ppt','.xls'
$rows = foreach ($ext in $formats) {
  $files = Get-ChildItem (Join-Path $Source 'testdata\samples') -File -Filter "*$ext" |
    Where-Object Length -gt 100 | Sort-Object Length | Select-Object -First $PerFormat
  foreach ($file in $files) {
    Push-Location $Source
    try { $goText = (& go run ./cmd/officeread -text-only $file.FullName 2>$null) -join "`n" }
    finally { Pop-Location }
    $cppText = (& $Cpp -text-only $file.FullName 2>$null) -join "`n"
    [pscustomobject]@{
      Format = $ext; File = $file.Name; GoChars = $goText.Length; CppChars = $cppText.Length
      NonEmpty = [bool]$cppText; ContainsGo = [bool]($goText -and $cppText.Contains($goText))
    }
  }
}
$rows | Format-Table -AutoSize
$bad = @($rows | Where-Object { $_.GoChars -gt 0 -and -not $_.NonEmpty })
Write-Output "Compared $($rows.Count) files; C++ empty while Go non-empty: $($bad.Count)"
if ($bad.Count) { exit 1 }
