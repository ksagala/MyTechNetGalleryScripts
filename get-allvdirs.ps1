<#
    .SYNOPSIS
	Script lists all virtual directories and names used by all Exchange servers in organization
    By Konrad Sagała
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.4
	History:
	Version 1.0
		Initial version
	Version 1.1
		Few small error fixed
	Version 1.2
		focusing on MAPI HTTP, Updated to Exchange 2013 SP1
    Version 1.3
        updated to Exchange 2016 RTM (15.1)
    Version 1.4
        some small improvements especially for Autodiscover, added writing all information to text file

    .DESCRIPTION
	
	Script lists all virtual directories and names used by all Exchange servers in organization, results are written to the output file

    .PARAMETER Name
    Server Name

    .EXAMPLE
    Get configuration fr whole Exchange organization
    .\get-allvdirs.ps1

#>
$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"

$Logpath = 'c:\Scripts'
if (-not(Test-Path $LogPath))
	{
	Write-Log -Message "Folder doesn't exist. Creating."
	New-Item -Path $Logpath -ItemType Directory
	}
Start-Transcript -path C:\Scripts\vdirs.txt -append

# start collecting information about virtual directories, according to Exchange version  

$ExOrgCfg = Get-OrganizationConfig
if ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)
	{
	if ((($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Build -ge 847) -and ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Minor -eq 0)) -or ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Minor -eq 1))
	{
		$vdirs = @("Activesync","ECP","mapi","OWA","OAB","PowerShell","WebServices")
		If ($ExOrgCfg.MapiHTTPEnabled)
		{Write-Host -ForegroundColor Green "MAPI HTTP Enabled, please verify MAPI VirtualDirectory settings"}
	}
	else
	{
		$vdirs = @("Activesync","ECP","OWA","OAB","PowerShell","WebServices")
	}
	}
else
	{
		$vdirs = @("Activesync","ECP","OWA","OAB","PowerShell","WebServices")
	}

foreach ($i in $vdirs) {
	$cmd = "Get-"+$i+"VirtualDirectory"+" | fl name,server,internalurl,externalurl,*AuthenticationMethods"
	invoke-expression $cmd
}

# check basic information about Autodiscover SCP 

if (($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)-and ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Minor -eq 1))
	{
		Get-ClientAccessService | fl name,OutlookAnywhereEnabled,AutoDiscoverSiteScope,AutoDiscoverServiceInternalUri
	}
else
	{
		Get-ClientAccessServer | fl name,OutlookAnywhereEnabled,AutoDiscoverSiteScope,AutoDiscoverServiceInternalUri
	}


# RPC over HTTPS configuration for all servers

Get-OutlookAnywhere | fl ServerName,SSLOffloading,*hostname,*AuthenticationMethod

# stopping writing results to file
Stop-Transcript
