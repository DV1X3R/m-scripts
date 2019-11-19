$templateFile = ".\ARMTemplateForFactory.json"

Write-Output "ARMTemplateForFactory.json                Resource Linked Info Updater"
Write-Output "Resource Name: $resourceName"
Write-Output "Linked Resource Id: $linkedResourceId"
Write-Host "=============================================================================="

Write-Host "Reading ARMTemplateForFactory.json..."
$template = Get-Content $templateFile | ConvertFrom-Json
$resource = $template.resources | Where( {$_.name -eq "[concat(parameters('factoryName'), '/$resourceName')]"} ) | Select -First 1

if($null -eq $resource){
    Write-Warning "Resource '$resourceName' not found"
    exit 0
}

$currentValue = Invoke-Expression "`$resource.properties.typeProperties.linkedInfo"
if($null -eq $currentValue){
    Write-Host "Adding new linkedInfo member"
    $linkedInfo = @{ resourceId = $linkedResourceId; authorizationType = "Rbac" }
    Add-Member -InputObject $resource.properties.typeProperties -MemberType NoteProperty -Name "linkedInfo" -Value $linkedInfo
}else{
    Write-Host "Updating old linkedInfo member"
    $resource.properties.typeProperties.linkedInfo.resourceId = $linkedResourceId
}

$templateOutput = $template | ConvertTo-Json -Depth 50 | ForEach-Object {
    [Regex]::Replace($_, 
        "\\u(?<Value>[a-fA-FZ0-9]{4})", {
            param($m) ([char]([int]::Parse($m.Groups['Value'].Value,
                [System.Globalization.NumberStyles]::HexNumber))).ToString() } )}

Write-Host "Updating ARMTemplateForFactory.json..."
Set-Content -Path $templateFile -Value $templateOutput
