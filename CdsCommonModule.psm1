# CDS requires [System.Net.SecurityProtocolType]::Tls12
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls -bor [System.Net.SecurityProtocolType]::Tls12

#Set-ExecutionPolicy –ExecutionPolicy RemoteSigned –Scope CurrentUser

<#
.Synopsis
	Checks if the module exists and imports it, otherwise installs it

.Parameter Name
	Module name

.Example 
    Register-Module -Name Microsoft.Xrm.OnlineManagementAPI
#>
function Register-Module {
    [CmdletBinding()]
    Param(
        [string] [Parameter(Mandatory = $true)] $Name
    )
    Write-Host "ERIC" $Name
    if (-not (Get-Module -Name $Name -ListAvailable)) {
        Write-Host "Module $Name not installed so installing" -ForegroundColor Green
        Install-Module -Name $Name -Scope AllUsers
    }
    else {
        Write-Host "Import module $Name" -ForegroundColor Green
        Import-Module $Name
    }
}

<#
.Synopsis
	Initiates a new CDS connection.

.Description
    Uses the Powershell module https://github.com/seanmcne/Microsoft.Xrm.Data.PowerShell

.Parameter UserName
	Account to connect to CDS/D365 with

.Parameter Password
	Password for the account

.Parameter Region
	Region to connect to

.Parameter Organisation
    Name of the CDS instance to connect to
    
.Example 
    $Connection = Get-CdsConnection -UserName $UserName -Password $Password -Region $Region -Organisation $Organisation -MaxAttempts 3
#>
function Get-CdsConnection {    
    [CmdletBinding()]
    Param(
        [parameter(Mandatory = $true)][string]$UserName,
        [parameter(Mandatory = $true)][string]$Password,
        [parameter(Mandatory = $true)][string]$Region,
        [parameter(Mandatory = $true)][string]$Organisation,
        [parameter(Mandatory = $false)][int]$MaxAttempts = 1
    )

    $pw = ConvertTo-SecureString $Password -AsPlainText -Force
    $creds = New-Object System.Management.Automation.PSCredential ($UserName, $pw)

    Write-Host "Connecting to $Organisation with $UserName" -ForegroundColor Green
    $attempt = 0

    do {
        $attempt ++
        $failed = $false
		
        try {
            $conn = Get-CrmConnection -OnLineType Office365 -OrganizationName $Organisation -DeploymentRegion $Region -Credential $creds
        }
        catch {
            $failed = $true
            if ($attempt -ge $MaxAttempts) {
                throw $_
            }
            else {
                Write-Host "Unable to get connection with $attempt attempts. Error: " $_.ToString() -ForegroundColor Cyan
                Start-Sleep -Seconds 300
            }
        }
    } while ($failed)

    return $conn
}