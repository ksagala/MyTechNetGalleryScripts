#
# Script to export almost all config data from Exchange Servers
# Based on script ExDoc by MikeT52
#
# Konrad Sagala
#
# Tested with Exchange 2013, 2010 and 2007
#
# ver. 1.1
#
# Few small error fixed
# added additional conditions for public folders and certificates verification
#

# date formated in polish standard - could be modified
$DateString = Get-Date -Format yyyyMMdd
$OutputDir = 'C:\Scripts\Docs\'

# Set Basic Variables
$allVersions = @()
$allBuilds = @()
$CASArrayName = ""
$minver=15
$maxver=6
$i=0
	
function RunGetCommand ($getCommand, $parameters){

	$filename = $OutputDir+$DateString+$getCommand

	$function = $getCommand + $parameters + " | fl >> " + $filename + ".txt"
	Invoke-Expression -Command $function
	
}

function RunGetCommandPipeline ($getCommand, $pipelineGetCommand, $parameters){

	$filename = $OutputDir+$DateString+$getCommand

	$function = $pipelineGetCommand + " | " + $getCommand + $parameters + " | fl >> " + $filename + ".txt"
	Invoke-Expression -Command $function
	
}

#
# main script part
#

# Exchange Server oldest Version in organization
$ExchangeServers = [array](Get-ExchangeServer)
foreach ($server in $ExchangeServers)
{
	$actual = $ExchangeServers[$i].AdminDisplayVersion.Major
	$allbuilds = $ExchangeServers[$i].AdminDisplayVersion.Build
	$allVersions += $actual
	if ($minver -gt $actual)
		{$minver = $actual}
	if ($maxver -lt $actual)
		{$maxver = $actual}
	$i++
}
Write-Host "Exchange docs processing for servers from version "$minver" to version "$maxver
#for Exchange Certificates on all servers must be separate loop created
$servers = Get-ExchangeServer
$filename = $OutputDir+$DateString+"Get-ExchangeCertificate.txt"
foreach ($server in $servers)
{
    if ($server.AdminDisplayVersion.Major -ge 8)
        { Get-ExchangeCertificate -Server $server.Name | fl >> $filename }
}
#
# Now we are executing get-* cmdlets for all components with functions defined in the beginning of the script
#
RunGetCommand "Get-AcceptedDomain" ""
RunGetCommand "Get-ActiveSyncOrganizationSettings" ""
if ($minver -eq 14)
{
	RunGetCommand "Get-ActiveSyncMailboxPolicy" ""
	RunGetCommand "Get-ActiveSyncDevice" ""
}
elseif ($minver -eq 15)
{
	# For Exchange 2013 first verifying if public folder mailbox exist
	$pubfoldermbx = $null
	$pubfoldermbx = Get-Mailbox -PublicFolder
	if ($pubfoldermbx -ne $null)
		{RunGetCommand "Get-MailPublicFolder" ""}
	RunGetCommand "Get-MobileDeviceMailboxPolicy" ""
	RunGetCommand "Get-MobileDevice" ""
	RunGetCommand "Get-MalwareFilterPolicy" ""
	RunGetCommand "Get-MalwareFilteringServer" ""
}
RunGetCommand "Get-AddressList" ""
# RunGetCommand "Get-AddressRewriteEntry" "" # Only when using rewriting
RunGetCommand "Get-AdSite" ""
RunGetCommand "Get-AdSiteLink" ""
RunGetCommandPipeline "Get-AutodiscoverVirtualDirectory" "Get-ClientAccessServer" ""
RunGetCommand "Get-AvailabilityAddressSpace" ""
RunGetCommand "Get-AvailabilityConfig" ""
RunGetCommand "Get-CASMailbox" "  -ResultSize Unlimited"
RunGetCommand "Get-ClientAccessArray" ""
RunGetCommand "Get-ClientAccessServer" ""
RunGetCommand "Get-ContentFilterConfig" ""
RunGetCommand "Get-ContentFilterPhrase" ""
RunGetCommand "Get-DetailsTemplate" ""
RunGetCommand "Get-DistributionGroup" " -ResultSize Unlimited"
RunGetCommand "Get-DynamicDistributionGroup" ""
RunGetCommand "Get-EdgeSubscription" ""
RunGetCommand "Get-EmailAddressPolicy" ""
if ($minver -eq 8)
	{ RunGetCommand "Get-ExchangeAdministrator" "" }
RunGetCommand "Get-ExchangeServer" ""
RunGetCommand "Get-ForeignConnector" ""
RunGetCommand "Get-GlobalAddressList" " -DefaultOnly"
RunGetCommand "Get-IPAllowListConfig" ""
RunGetCommand "Get-IPBlockListConfig" ""
RunGetCommand "Get-JournalRule" ""
RunGetCommandPipeline "Get-IMAPSettings" "Get-ClientAccessServer" ""
RunGetCommand "Get-Mailbox" " -ResultSize Unlimited"
RunGetCommandPipeline "Get-MailboxDatabase" "Get-MailboxServer" ""
if ($maxver -lt 15)
{
	RunGetCommand "Get-ManagedContentSettings" ""
	RunGetCommand "Get-ManagedFolder" ""
	RunGetCommand "Get-ManagedFolderMailboxPolicy" ""
	RunGetCommand "Get-RoutingGroupConnector" ""
	RunGetCommand "Get-TransportServer" ""
	RunGetCommand "Get-UMServer" ""
	RunGetCommand "Get-PublicFolder" " -Recurse"
	RunGetCommandPipeline "Get-PublicFolderDatabase" "Get-MailboxServer" ""
}
RunGetCommand "Get-MailboxServer" ""
RunGetCommand "Get-MessageClassification" ""
RunGetCommand "Get-OfflineAddressBook" ""
RunGetCommandPipeline "Get-POPSettings" "Get-ClientAccessServer" ""
RunGetCommand "Get-OrganizationConfig" ""
RunGetCommandPipeline "Get-OutlookAnywhere" "Get-ClientAccessServer" ""
RunGetCommand "Get-OutlookProvider" ""
RunGetCommand "Get-OWAMailboxPolicy" ""
RunGetCommand "Get-ActiveSyncVirtualDirectory" ""
if ($minver -ge 14)
	{ RunGetCommand "Get-ECPVirtualDirectory" "" }
RunGetCommand "Get-OWAVirtualDirectory" ""
RunGetCommand "Get-OABVirtualDirectory" ""
RunGetCommand "Get-WebServicesVirtualDirectory" ""
$ExOrgCfg = Get-OrganizationConfig
if ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)
	{
	if ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Build -ge 847)
	{
		RunGetCommand "Get-MAPIVirtualDirectory" ""
	}
	}
RunGetCommand "Get-ReceiveConnector" ""
RunGetCommand "Get-RemoteDomain" ""
RunGetCommand "Get-SendConnector" ""
RunGetCommand "Get-SenderFilterConfig" ""
RunGetCommand "Get-SenderIdConfig" ""
RunGetCommand "Get-SenderReputationConfig" ""
if ($minver -eq 8)
	{ RunGetCommandPipeline "Get-StorageGroup" "Get-MailboxServer" "" }
RunGetCommand "Get-TransportConfig" ""
RunGetCommand "Get-TransportRule" ""
RunGetCommand "Get-TransportRuleAction" ""
RunGetCommand "Get-TransportRulePredicate" ""
RunGetCommand "Get-UMAutoAttendant" ""
RunGetCommand "Get-UMDialPlan" ""
RunGetCommand "Get-UMHuntGroup" ""
RunGetCommand "Get-UMIPGateway" ""
RunGetCommand "Get-UMMailbox" ""
RunGetCommand "Get-UMMailboxPolicy" ""
if ($minver -eq 8)
	{ RunGetCommand "Get-UMVirtualDirectory" "" }
RunGetCommand "Get-X400AuthoritativeDomain" ""
if ($minver -ge 14)
	{ 
	RunGetCommand "Get-DatabaseAvailabilityGroup" ""
	RunGetCommand "Get-DatabaseAvailabilityGroupNetwork" ""
	RunGetCommand "Get-DatabaseAvailabilityGroupConfiguration" ""
	RunGetCommand "Get-ActiveSyncDeviceAccessRule" ""
	}

Write-Host "All data are written to c:\Scripts\Docs\"