Import-Module "$PSScriptRoot\CdsCommonModule.psm1"

Register-Module -Name Microsoft.Xrm.Data.PowerShell

function Remove-Field(
    [CmdletBinding()]
    [string] $FieldName,
    [string] $EntityName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting field" $FieldName -ForegroundColor Green

    $retrieveRequest = new-object Microsoft.Xrm.Sdk.Messages.RetrieveAttributeRequest
    $retrieveRequest.LogicalName = $FieldName
    $retrieveRequest.EntityLogicalName = $EntityName
    $retrieveRequest.RetrieveAsIfPublished = $false
    $response = $Connection.ExecuteCrmOrganizationRequest($retrieveRequest)
     
    if ($null -ne $response.AttributeMetadata) {     
        $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteAttributeRequest
        $deleteRequest.LogicalName = $FieldName
        $deleteRequest.EntityLogicalName = $EntityName
        try {
            $Connection.ExecuteCrmOrganizationRequest($deleteRequest)
        }
        catch { 
            Write-Host "Field $FieldName failed to delete" -ForegroundColor Red
        }
        Write-Host "Field $FieldName deleted" -ForegroundColor Green
    }
    else {
        Write-Host "Field $FieldName does not exist" -ForegroundColor Cyan
    }
}

function Remove-Process(
    [CmdletBinding()]
    [string] $ComponentName,
    [int] $ComponentType,
    [Guid] $ComponentId,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting process" $ComponentName -ForegroundColor Green

    $fetchXml =
    @"
<fetch>
  <entity name="$ComponentName">
  
    <filter type="and" >
        <condition attribute="category" operator="eq" value="$ComponentType" />
        <condition attribute="{0}id" operator="eq" value="$ComponentId" />      
    </filter>
  </entity>
</fetch>
"@ -f $ComponentName

    Write-Output $fetchXml

    $response = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml

    if ($response.CrmRecords.Count -gt 0) {

        Set-CrmRecordState -conn $Connection -EntityLogicalName $ComponentName -Id $ComponentId -StateCode 0 -StatusCode 1    

        $ref = new-object Microsoft.Xrm.Sdk.EntityReference
        $ref.LogicalName = $ComponentName
        $ref.Id = $ComponentId

        $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteRequest
        $deleteRequest.Target = $ref
                 
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)

        Write-Host "Component $ComponentName $ComponentId deleted" -ForegroundColor Green
    }
    else {
        Write-Host "Component $ComponentName $ComponentId does not exist" -ForegroundColor Cyan
    }
}

function Remove-Form(
    [CmdletBinding()]
    [Guid] $FormId,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting form" $FormId -ForegroundColor Green

    $fetchXml =
    @"
<fetch>
  <entity name="systemform">
    <filter type="and" >
        <condition attribute="formid" operator="eq" value="$FormId" />      
    </filter>
  </entity>
</fetch>
"@

    Write-Output $fetchXml

    $response = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml

    if ($response.CrmRecords.Count -gt 0) {

        $ref = new-object Microsoft.Xrm.Sdk.EntityReference
        $ref.LogicalName = "systemform"
        $ref.Id = $FormId

        $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteRequest
        $deleteRequest.Target = $ref
                 
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)

        Write-Host "Form $FormId deleted" -ForegroundColor Green
    }
    else {
        Write-Host "Form $FormId does not exist" -ForegroundColor Cyan
    }
}

function Remove-View(
    [CmdletBinding()]
    [Guid] $ViewId,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting View" $ViewId -ForegroundColor Green

    $fetchXml =
    @"
<fetch>
  <entity name="savedquery">
    <filter type="and" >
        <condition attribute="savedqueryid" operator="eq" value="$ViewId" />      
    </filter>
  </entity>
</fetch>
"@

    Write-Output $fetchXml

    $response = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml

    if ($response.CrmRecords.Count -gt 0) {

        $ref = new-object Microsoft.Xrm.Sdk.EntityReference
        $ref.LogicalName = "savedquery"
        $ref.Id = $ViewId

        $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteRequest
        $deleteRequest.Target = $ref
                 
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)

        Write-Host "View $ViewId deleted" -ForegroundColor Green
    }
    else {
        Write-Host "View $ViewId does not exist" -ForegroundColor Cyan
    }
}

function Remove-WebResource(
    [CmdletBinding()]
    [Guid] $WebResourceId,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting webresource" $WebResourceId -ForegroundColor Green

    $fetchXml =
    @"
<fetch>
  <entity name="webresource">  
    <filter type="and" >
        <condition attribute="webresourceid" operator="eq" value="$WebResourceId" />      
    </filter>
  </entity>
</fetch>
"@

    Write-Output $fetchXml

    $response = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml

    if ($response.CrmRecords.Count -gt 0) {

        $ref = new-object Microsoft.Xrm.Sdk.EntityReference
        $ref.LogicalName = "webresource"
        $ref.Id = $WebResourceId

        $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteRequest
        $deleteRequest.Target = $ref
                 
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)

        Write-Host "WebResource $WebResourceId deleted" -ForegroundColor Green
    }
    else {
        Write-Host "WebResource $WebResourceId does not exist" -ForegroundColor Cyan
    }
}

function Remove-Plugin(
    [CmdletBinding()]
    [Guid] $RecordId,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting plugin $RecordId" -ForegroundColor Green

    try {
        $record = Get-CrmRecord -EntityLogicalName "plugintype" -Id $RecordId -Fields "plugintypeid" -conn $Connection
    }
    catch {
        Write-Host "Plugin $RecordId does not exist" -ForegroundColor Cyan
        return
    }

    $deleteRequest = new-object Microsoft.Crm.Sdk.Messages.RetrieveDependenciesForDeleteRequest
    $deleteRequest.ComponentType = 90 #plugin
    $deleteRequest.ObjectId = $RecordId
      
    $response = $Connection.ExecuteCrmOrganizationRequest($deleteRequest)  
       
    if ($null -ne $response) {
        if ($response.EntityCollection.Entities.Count -gt 0) {
            $response.EntityCollection.Entities | ForEach-Object -Process {
                           
                $id = $_.Attributes["dependentcomponentobjectid"].ToString()              
                Write-Output $id
         
                Remove-CrmRecord -conn $Connection -CrmRecord @{ "sdkmessageprocessingstepid" = $id; "LogicalName" = "sdkmessageprocessingstep" }
            }
        }
    }

    Remove-CrmRecord -conn $Connection -CrmRecord $record

    Write-Host "Plugin $RecordId deleted" -ForegroundColor Green
}

function Remove-Entity(
    [CmdletBinding()]
    [string] $EntityName,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting entity" $EntityName -ForegroundColor Green

    $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteEntityRequest
    $deleteRequest.LogicalName = $EntityName

    try {
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)
    }
    catch { 
        Write-Host "Entity $EntityName failed to delete" -ForegroundColor Red
    }
    Write-Host "Entity $EntityName deleted" -ForegroundColor Green
}

function Remove-Solutions(
    [CmdletBinding()]
    [string] $FetchXmlFilterCondition,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Fetching solutions to delete" -ForegroundColor Green 

    if ($null -eq $FetchXmlFilterCondition) {
        Write-Host "FetchXmlFilterCondition param missing" -ForegroundColor Red
        return;
    }

    $fetchXml =
    @"
<fetch>
  <entity name="solution" >
    <attribute name="uniquename" />
    <attribute name="solutionid" />
    <attribute name="friendlyname" />
    <attribute name="installedon" />
    $FetchXmlFilterCondition
	<order attribute="installedon" descending="true" />
  </entity>
</fetch>
"@
		
    $solutions = Get-CrmRecordsByFetch -conn $Connection -Fetch $fetchXml

    if ($solutions.CrmRecords.Count -gt 0) {
        $solutions.CrmRecords | ForEach-Object -Process {
         
            Write-Host "Deleting solution:" $_.uniquename -ForegroundColor Green
           
            $ref = new-object Microsoft.Xrm.Sdk.EntityReference
            $ref.LogicalName = "solution"
            $ref.Id = $_.solutionid

            $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteRequest
            $deleteRequest.Target = $ref  
				
            $Connection.ExecuteCrmOrganizationRequest($deleteRequest)

            Write-Host "Solution: " $_.uniquename " deleted" -ForegroundColor Green
        }
    }
    else {
        Write-Host "No solutions to delete" -ForegroundColor Cyan
    }
}

function Update-QueueServiceEndpoint(
    [CmdletBinding()]
    [Guid] $ServiceEndpointId,
    [string] $SasKey,
    [string] $Address,
    [string] $Path,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Updating service endpoint" $ServiceEndpointId -ForegroundColor Green 

    $updateFields = @{ }
    $updateFields.Add("namespaceaddress", $Address)   
    $updateFields.Add("saskey", $Saskey)  
    $updateFields.Add("path", $Path)

    Set-CrmRecord -conn $Connection -Fields $updateFields -Id $ServiceEndpointId -EntityLogicalName "serviceendpoint"

    Write-Host "Service endpoint updated" $ServiceEndpointId -ForegroundColor Green 
}

function Update-WebhookServiceEndpoint(
    [CmdletBinding()]
    [Guid] $ServiceEndpointId,
    [string] $SasKey,
    [string] $Address,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Updating webhook endpoint" $ServiceEndpointId -ForegroundColor Green 

    $updateFields = @{ }
    $updateFields.Add("url", $Address)
    $updateFields.Add("authvalue", $Saskey)

    Set-CrmRecord -conn $Connection -Fields $updateFields -Id $ServiceEndpointId -EntityLogicalName "serviceendpoint"

    Write-Host "Service webhook updated" $ServiceEndpointId -ForegroundColor Green 
}

function Remove-OptionSet(
    [CmdletBinding()]
    [string] $Name,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting option set" $Name -ForegroundColor Green
   
    $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteOptionSetRequest
    $deleteRequest.Name = $Name
    try {
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)
    }
    catch { 
        Write-Host "Option set $Name does not exist" -ForegroundColor Cyan
    }
    Write-Host "Option set $Name deleted" -ForegroundColor Green
}

function Remove-OptionSetItem(
    [CmdletBinding()]
    [string] $OptionSetName,
    [string] $Value,
    [Microsoft.Xrm.Tooling.Connector.CrmServiceClient] $Connection
) {
    Write-Host "Deleting option set item" $Value -ForegroundColor Green
   
    $deleteRequest = new-object Microsoft.Xrm.Sdk.Messages.DeleteOptionValueRequest
    $deleteRequest.OptionSetName = $OptionSetName
    $deleteRequest.Value = $Value
    try {
        $Connection.ExecuteCrmOrganizationRequest($deleteRequest)
    }
    catch { 
        Write-Host "Option set item $Value does not exist" -ForegroundColor Cyan
    }
    Write-Host "Option set item $Value deleted" -ForegroundColor Green
}