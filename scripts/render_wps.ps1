param(
 [Parameter(Mandatory=$true)][string]$Source,
 [Parameter(Mandatory=$true)][string]$ImageDirectory,
 [int]$Width=2560
)
$ErrorActionPreference='Stop'
$sourcePath=(Resolve-Path -LiteralPath $Source).Path
$imagePath=[IO.Path]::GetFullPath($ImageDirectory)
if(Test-Path -LiteralPath $imagePath){
 if((Get-ChildItem -LiteralPath $imagePath -Force | Measure-Object).Count -gt 0){throw 'Use a new or empty image directory.'}
}else{New-Item -ItemType Directory -Path $imagePath | Out-Null}
$app=New-Object -ComObject KWPP.Application
$presentation=$null
try {
 $presentation=$app.Presentations.Open($sourcePath,$true,$false,$false)
 $height=[int][Math]::Round($Width*$presentation.PageSetup.SlideHeight/$presentation.PageSetup.SlideWidth)
 for($i=1;$i -le $presentation.Slides.Count;$i++){
  $page=$presentation.Slides.Item($i)
  $page.Export((Join-Path $imagePath ('slide'+$i+'.png')),'PNG',$Width,$height)
 }
 Write-Output ('WPS rendered '+$presentation.Slides.Count+' slides to '+$imagePath)
} finally {
 if($null -ne $presentation){$presentation.Close()}
 # Do not quit WPS: the user may have other files open.
}
