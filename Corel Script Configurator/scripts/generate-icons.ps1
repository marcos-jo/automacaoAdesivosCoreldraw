param([string]$AppSource,[string]$PdfSource,[string]$ProjectRoot)
$ErrorActionPreference='Stop'
if(-not $ProjectRoot){if($PSScriptRoot){$ProjectRoot=Split-Path -Parent $PSScriptRoot}else{$ProjectRoot=(Get-Location).Path}}
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition ([IO.File]::ReadAllText((Join-Path $ProjectRoot 'scripts/icons/IconAssets.cs'))) -ReferencedAssemblies System.Drawing
$iconFolder=[IO.Path]::GetFullPath((Join-Path $ProjectRoot 'build/icons'))
$sourceFolder=Join-Path $iconFolder 'sources'
[IO.Directory]::CreateDirectory($sourceFolder) | Out-Null
foreach($entry in @(@($AppSource,'icone4ps.png'),@($PdfSource,'iconepdf.png'))){
    $destination=Join-Path $sourceFolder $entry[1]
    if($entry[0]){
        $original=(Resolve-Path -LiteralPath $entry[0]).Path
        if(Test-Path -LiteralPath $destination){
            if((Get-FileHash -LiteralPath $original).Hash -ne (Get-FileHash -LiteralPath $destination).Hash){throw "Original diferente da copia preservada: $destination"}
        }else{[IO.File]::Copy($original,$destination,$false)}
    }
    if(-not (Test-Path -LiteralPath $destination)){throw "Arte original ausente: $destination"}
}
[IconAssets]::Generate((Join-Path $sourceFolder 'icone4ps.png'),(Join-Path $iconFolder 'app'))
[IconAssets]::Generate((Join-Path $sourceFolder 'iconepdf.png'),(Join-Path $iconFolder 'pdf'))
[IconAssets]::Preview((Join-Path $sourceFolder 'icone4ps.png'),(Join-Path $sourceFolder 'iconepdf.png'),(Join-Path $iconFolder 'preview.png'))
Write-Output "Gerados app.png, app.ico, pdf.png, pdf.ico e preview.png em $iconFolder. Originais preservados."
