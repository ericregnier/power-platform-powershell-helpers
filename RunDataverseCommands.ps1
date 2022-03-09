Param(
    [string] [Parameter(Mandatory = $true)] $UserName,
    [string] [Parameter(Mandatory = $true)] $Password,
    [string] [Parameter(Mandatory = $true)] $Region,
    [string] [Parameter(Mandatory = $true)] $Organisation,
    [string] [Parameter(Mandatory = $true)] $TenantId
)

Import-Module "$PSScriptRoot\DataverseOperations.psm1"

$Connection = Get-DataverseConnection -UserName $UserName -Password $Password -Region $Region -Organisation $Organisation -MaxAttempts 3

Add-PowerAppsAccount -TenantID $TenantId -Username $Username -Password $Password

#In the below example, the $Env:SERVICEBUSURL and $Env:AZUREQUEUENAME are from Azure DevOps variable
Update-QueueServiceEndpoint -Connection $Connection -ServiceEndpointId "00000000-0000-0000-0000-000000000000" -SasKey "ServiceBusSasKey" -Address $Env:SERVICEBUSURL -Path $Env:AZUREQUEUENAME

Update-WebhookServiceEndpoint -Connection $Connection -ServiceEndpointId "00000000-0000-0000-0000-000000000000" -SasKey "WebhookKey" -Address "https://address.com"

Remove-Field -FieldName "cr06a_fieldname" -EntityName "cr06a_entityname" -Connection $Connection

#Business Rule category = 2, classic workflow = 0
Remove-Process -ComponentName "workflow" -ComponentType 0 -ComponentId "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-Form -FormId "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-View -ViewId "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-Entity -EntityName "cr06a_entityname" -Connection $Connection

Remove-Field -FieldName "cr06a_fieldname" -EntityName "cr06a_entityname" -Connection $Connection

Remove-OptionSet -Name "cr06a_optionsetname" -Connection $Connection

Remove-OptionSetItem -OptionSetName "cr06a_optionsetname" -Value 100000000 -Connection $Connection

Remove-Plugin -RecordId "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-PluginStep -Id "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-WebResource -WebResourceId "00000000-0000-0000-0000-000000000000" -Connection $Connection

$condition =
    @"
    <filter type="or" >
        <condition attribute="uniquename" operator="in" >
            <value>SolutionX</value>
            <value>SolutionY</value>
        </condition>
    </filter>
"@
Remove-Solutions -FetchXmlFilterCondition $condition -Connection $Connection

Update-EnvVariable -EnvVariableId "00000000-0000-0000-0000-000000000000" -Connection $Connection

Remove-EnvVariable -Name "cr06a_envvar" -Connection $Connection

Remove-CanvasApp -EnvironmentName "prod" -CanvasAppName "cr06a_app"
