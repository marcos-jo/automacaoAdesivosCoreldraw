param([string]$ProjectRoot=(Split-Path -Parent $PSScriptRoot))
$ErrorActionPreference='Stop'
$lock=Get-Content -LiteralPath (Join-Path $projectRoot 'build/pdfium-lock.json') -Raw | ConvertFrom-Json
$destination=Join-Path $projectRoot 'build/pdfium'
New-Item -ItemType Directory -Path $destination -Force | Out-Null
$archive=Join-Path $destination $lock.asset
if(-not (Test-Path -LiteralPath $archive)){
    Invoke-WebRequest -Uri $lock.url -OutFile $archive
}
$hash=(Get-FileHash -LiteralPath $archive -Algorithm SHA256).Hash.ToLowerInvariant()
if($hash -ne $lock.sha256){throw 'O pacote PDFium difere do hash fixado. O arquivo foi preservado para investigação.'}
$entries=@(& tar -tf $archive)
if($LASTEXITCODE -ne 0){throw 'Não foi possível inspecionar o arquivo PDFium.'}
foreach($entry in $entries){
    if([IO.Path]::IsPathRooted($entry) -or $entry -match '(^|[/\\])\.\.([/\\]|$)'){
        throw 'Caminho de extração inválido.'
    }
}
& tar -xf $archive -C $destination
if($LASTEXITCODE -ne 0){throw 'Extração PDFium falhou.'}
Write-Output ('PDFium '+$lock.tag+' conferido e preparado em '+$destination)
