<#
    .SYNOPSIS
	Script to allow you to set all virtual directories to a common name like mail.company.com
    Based on script published by Nathan Winters on his blog
    http://nathanwinters.co.uk/2010/05/30/script-to-set-internalurl-and-externalurl-for-all-exchange-2010-virtual-directories/
    Sources mentioned by Nathan (e.g. exchangeninjas.com and Barry Martin blog) are unavailable now but I'd like to thanks to all of previous authors of this script
    
    By Konrad Sagała
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.2
	History:
	Version 1.0
		Initial version (June 2010)
	Version 1.1
        added powershell virtual directory differentiation between http and https
        verified with Exchange 2016
        fixed bugs with Outlook Anywhere 
	Version 1.2
        Script was updated with corrected build numbers to work in native Exchange 2016 environment and change MAPI vdirs

    .DESCRIPTION
	
	Script to allow you to set all virtual directories to a common name like mail.company.com

    .EXAMPLE
    Get configuration fr whole Exchange organization
    .\set-allvdirs.ps1

#>

# Variables 

$localserver = $env:computerName
$ExOrgCfg = Get-OrganizationConfig

[string]$EASExtend = "/Microsoft-Server-ActiveSync" 
[string]$PShExtend = "/powershell" 
[string]$OWAExtend = "/OWA" 
[string]$OABExtend = "/OAB" 
[string]$SCPExtend = "/Autodiscover/Autodiscover.xml" 
[string]$EWSExtend = "/EWS/Exchange.asmx" 
[string]$ECPExtend = "/ECP"
[string]$MapiExtend = "/mapi"
[string]$ConfirmPrompt = "Set this Value? (Y/N)" 
[string]$NoChangeForeground = "white" 
[string]$NoChangeBackground = "red" 

Write-host "This will allow you to set the virtual directories associated with setting up a single SSL certificate to work with Exchange 2016, 2013 or 2010." 
Write-host "" 
[string]$base = Read-host "FQDN of Exchange Servers assigned for virtual directories (e.g. mail.company.com)" 
write-host "" 

# =================================================================== 
# Validate if OAB downloads are delivered with HTTP or HTTPS protocol
# ===================================================================  
[string]$set = Read-host "Do you want to use HTTP for OAB? (Y/N)" 
Write-host "" 

if ($set -eq "Y")    { 
    [string]$OABprefix = "http://" 
    [boolean]$OABRequireSSL = $false 
}    else    { 
    [string]$OABprefix = "https://" 
    [boolean]$OABRequireSSL = $true 
} 


# ================================================================================ 
# Validate if an operation is delivered to all CAS servers or only to local server 
# ================================================================================
[string]$isglobal = Read-host "Do you want to change vdirs urls for all CAS servers (Y) or for local server only (N)? (Y/N)" 
Write-host "" 


# =============================================
# Build the OAB URL and set the internal Value 
# =============================================
Write-host "Setting OAB Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$OABURL = $OABprefix + $base + $OABExtend 

if ($isglobal -eq "Y")
{
  [array]$OABCurrent = Get-OABVirtualDirectory 

  Foreach ($value in $OABcurrent) { 
    Write-host "Looking at Server: " $value.server 
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $OABUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")    { 
        Set-OABVirtualDirectory -id $value.identity -InternalURL $OABURL -RequireSSL:$OABRequireSSL 
    }
    else
    { 
        write-host "OAB Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $OABUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y") { 
        Set-OABVirtualDirectory -id $value.identity -ExternalURL $OABURL -RequireSSL:$OABRequireSSL 
    }
    else
    { 
        write-host "OAB Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
  }
}
else
{
    Write-host "OAB change for Server: " $localserver
    $value = Get-OABVirtualDirectory -server $localserver
	Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $OABUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")    { 
        Set-OABVirtualDirectory -id $value.identity -InternalURL $OABURL -RequireSSL:$OABRequireSSL 
    }
    else
    { 
        write-host "OAB Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $OABUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y") { 
        Set-OABVirtualDirectory -id $value.identity -ExternalURL $OABURL -RequireSSL:$OABRequireSSL 
    }
    else
    { 
        write-host "OAB Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
}

# ============================================= 
# Build the EWS URL and set the internal Value 
# =============================================
Write-host "Setting Exchange Web Services Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$EWSURL = "https://" + $base + $EWSExtend 

if ($isglobal -eq "Y")
{
  [array]$EWSCurrent = Get-WebServicesVirtualDirectory 

  Foreach ($value in $EWSCurrent) { 
    Write-host "Looking at Server: " $value.server 
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $EWSUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-WebServicesVirtualDirectory -id $value.identity -InternalURL $EWSURL 
    }
    else
    { 
        write-host "Exchange Web Services Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $EWSUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-WebServicesVirtualDirectory -id $value.identity -ExternalURL $EWSURL 
    }
    else
    { 
        write-host "Exchange Web Services Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
  }
}
else
{
    Write-host "EWS change for Server: " $localserver
    $value = Get-WebServicesVirtualDirectory -server $localserver
	
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $EWSUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-WebServicesVirtualDirectory -id $value.identity -InternalURL $EWSURL 
    }
    else
    { 
        write-host "Exchange Web Services Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $EWSUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-WebServicesVirtualDirectory -id $value.identity -ExternalURL $EWSURL 
    }
    else
    { 
        write-host "Exchange Web Services Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
}

# =================================================================================
# Validate if powershell virtual directory is published with HTTP or HTTPS protocol
# ================================================================================= 
[string]$set = Read-host "Do you want to use HTTP for Powershell? (Y/N)" 
Write-host "" 

if ($set -eq "Y")
{ 
    [string]$PSprefix = "http://" 
}
else
{ 
    [string]$PSprefix = "https://" 
} 

# ===================================================
# Build the PowerShell URL and set the internal Value
# ===================================================
Write-host "Setting Powershell Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$PShURL = $PSprefix + $base + $PShExtend 

if ($isglobal -eq "Y")
{
  [array]$PShCurrent = Get-PowerShellVirtualDirectory

  foreach ($value in $PShCurrent) { 
    Write-host "Looking at Server: " $value.server 
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $PShUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-PowerShellVirtualDirectory -id $value.identity -InternalURL $PShURL 
    }
    else
    { 
        write-host "PowerShell Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $PShUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-PowerShellVirtualDirectory -id $value.identity -ExternalURL $PShURL 
    }
    else
    { 
        write-host "PowerShell Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
  }
}
else
{
	$value = Get-PowerShellVirtualDirectory -server $localserver

    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $PShUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-PowerShellVirtualDirectory -id $value.identity -InternalURL $PShURL 
    }
    else
    { 
        write-host "PowerShell Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $PShUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y") { 
        Set-PowerShellVirtualDirectory -id $value.identity -ExternalURL $PShURL 
    }
    else
    { 
        write-host "PowerShell Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
}

# ============================================
# Build the ECP URL and set the internal Value 
# ============================================
Write-host "Setting ECP Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$ECPURL = "https://" + $base + $ECPExtend 

if ($isglobal -eq "Y")
{
  [array]$ECPCurrent = Get-ECPVirtualDirectory 

  foreach ($value in $ECPCurrent) { 
    Write-host "Looking at Server: " $value.server
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $ECPUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ECPVirtualDirectory -id $value.identity -InternalURL $ECPURL 
    }
    else
    { 
        write-host "ECP Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $ECPUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ECPVirtualDirectory -id $value.identity -ExternalURL $ECPURL 
    }
    else
    { 
        write-host "ECP Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    }
  }
}
else
{

	$value = Get-ECPVirtualDirectory -server $localserver
	
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $ECPUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ECPVirtualDirectory -id $value.identity -InternalURL $ECPURL 
    }
    else
    { 
        write-host "ECP Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $ECPUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ECPVirtualDirectory -id $value.identity -ExternalURL $ECPURL 
    }
    else
    { 
        write-host "ECP Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    }
}

# ============================================
# Build the OWA URL and set the internal Value
# ============================================
Write-host "Setting OWA Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$OWAURL = "https://" + $base + $OWAExtend 

if ($isglobal -eq "Y")
{
  [array]$OWACurrent = Get-OWAVirtualDirectory 

  foreach ($value in $OWACurrent) { 
    Write-host "Looking at Server: " $value.server 
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $OWAUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-OWAVirtualDirectory -id $value.identity -InternalURL $OWAURL 
    }
    else
    { 
        write-host "OWA Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $OWAUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-OWAVirtualDirectory -id $value.identity -ExternalURL $OWAURL 
    }
    else
    { 
        write-host "OWA Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
  }
} 
else
{
	$value = Get-OWAVirtualDirectory -server $localserver
	
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $OWAUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-OWAVirtualDirectory -id $value.identity -InternalURL $OWAURL 
    }
    else
    { 
        write-host "OWA Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $OWAUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-OWAVirtualDirectory -id $value.identity -ExternalURL $OWAURL 
    }
    else
    { 
        write-host "OWA Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 
}

# ============================================
# Build the EAS URL and set the internal Value
# ============================================
Write-host "Setting EAS Virtual Directories" -foregroundcolor Yellow 
write-host "" 

$EASURL = "https://" + $base + $EASExtend 

if ($isglobal -eq "Y")
{
  [array]$EASCurrent = Get-ActiveSyncVirtualDirectory 

  foreach ($value in $EASCurrent) { 
    Write-host "Looking at Server: " $value.server 
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $EASUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
      Set-ActiveSyncVirtualDirectory -id $value.identity -InternalURL $EASURL 
    }
    else
    { 
        write-host "EAS Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Looking at Server: " $value.server 
    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $EASUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ActiveSyncVirtualDirectory -id $value.identity -ExternalURL $EASURL 
    }
    else
    { 
        write-host "EAS Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    }
  }
} 
else
{
	$value = Get-ActiveSyncVirtualDirectory -server $localserver
    Write-host "Current Internal Value: " $value.internalURL 
    Write-host "New Internal Value:     " $EASUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
      Set-ActiveSyncVirtualDirectory -id $value.identity -InternalURL $EASURL 
    }
    else
    { 
        write-host "EAS Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    } 

    Write-host "Current External Value: " $value.externalURL 
    Write-host "New External Value:     " $EASUrl 
    [string]$set = Read-host $ConfirmPrompt 
    write-host "" 

    if ($set -eq "Y")
    { 
        Set-ActiveSyncVirtualDirectory -id $value.identity -ExternalURL $EASURL 
    }
    else
    { 
        write-host "EAS Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    }
}
#
# MAPI virtual directories - only for Exchange 2013 CU4 and later
#
# I also decided to set RPC virtual names only for Exchange 2013 from this script
#
if ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)
	{
	if ((($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Build -ge 847) -and ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Minor -eq 0)) -or ($ExOrgCfg.RBACConfigurationVersion.ExchangeBuild.Minor -eq 1))
	{
		Write-host "Setting MAPI Virtual Directories" -foregroundcolor Yellow 
		write-host "" 

		$MAPIURL = "https://" + $base + $MapiExtend 

		if ($isglobal -eq "Y")
		{
            [array]$MAPICurrent = Get-MAPIVirtualDirectory 
			foreach ($value in $MAPICurrent) { 
			    Write-host "Looking at Server: " $value.server 
    			Write-host "Current Internal Value: " $value.internalURL 
    			Write-host "New Internal Value:     " $MAPIUrl 
			    [string]$set = Read-host $ConfirmPrompt 
    			write-host "" 

    			if ($set -eq "Y")
                { 
			        Set-MAPIVirtualDirectory -id $value.identity -InternalURL $MAPIURL -IISAuthenticationMethods @('Ntlm', 'Oauth', 'Negotiate')
				}
                else
                { 
			        write-host "MAPI Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    			} 

			    Write-host "Looking at Server: " $value.server 
    			Write-host "Current External Value: " $value.externalURL 
    			Write-host "New External Value:     " $MAPIUrl 
    			[string]$set = Read-host $ConfirmPrompt 
    			write-host "" 

    			if ($set -eq "Y")
                { 
        			Set-MAPIVirtualDirectory -id $value.identity -ExternalURL $MAPIURL -IISAuthenticationMethods @('Ntlm', 'Oauth', 'Negotiate')
    			}
                else
                { 
        			write-host "MAPI Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
				} 
			}
		}
		else
		{
			$value = Get-MAPIVirtualDirectory -server $localserver
	
		    Write-host "Current Internal Value: " $value.internalURL 
    		Write-host "New Internal Value:     " $MAPIUrl 
    		[string]$set = Read-host $ConfirmPrompt 
    		write-host "" 

    		if ($set -eq "Y")
			{ 
		        Set-MAPIVirtualDirectory -id $value.identity -InternalURL $MAPIURL -IISAuthenticationMethods @('Ntlm', 'Oauth', 'Negotiate')
		    }
			else
			{ 
	        	write-host "MAPI Virtual Directory internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    		} 

    		Write-host "Current External Value: " $value.externalURL 
    		Write-host "New External Value:     " $MAPIUrl 
    		[string]$set = Read-host $ConfirmPrompt 
    		write-host "" 

    		if ($set -eq "Y")
			{ 
        		Set-MAPIVirtualDirectory -id $value.identity -ExternalURL $MAPIURL -IISAuthenticationMethods @('Ntlm', 'Oauth', 'Negotiate')
    		}
			else
			{ 
        		write-host "MAPI Virtual Directory external value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    		} 
		}
	}



#
#	Outlook Anywhere hostnames. Tested on Exchange 2013 and Exchange 2016
#	
	Write-host "Setting Outlook Anywhere hostnames" -foregroundcolor Yellow 
	write-host "" 

	if ($isglobal -eq "Y")
	{
		[array]$OACurrent = Get-OutlookAnywhere 
		foreach ($value in $OACurrent)
		{ 
		    Write-host "Looking at Server: " $value.server 
    		Write-host "Current Internal Value: " $value.InternalHostname 
    		Write-host "New Internal Value:     " $base 
		    [string]$set = Read-host $ConfirmPrompt 
    		write-host "" 

    		if ($set -eq "Y")
			{ 
		        Set-OutlookAnywhere -id $value.identity -InternalHostname $base -InternalClientsRequireSsl $true
			}
			else
			{ 
		        write-host "Outlook Anywhere internal hostname value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground
    		} 

		    Write-host "Looking at Server: " $value.server 
    		Write-host "Current External Value: " $value.externalhostname 
    		Write-host "New External Value:     " $base 
    		[string]$set = Read-host $ConfirmPrompt 
	   		write-host "" 

    		if ($set -eq "Y")
			{ 
       			Set-OutlookAnywhere -id $value.identity -ExternalHostname $base -ExternalClientsRequireSsl $true -ExternalClientAuthenticationMethod Negotiate  -IISAuthenticationMethods @('Ntlm', 'Basic', 'Negotiate')
			}
			else
			{ 
       			write-host "Outlook Anywhere external hostname value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
		    } 
		}
	}
	else
	{
		$value = Get-OutlookAnywhere -server $localserver
		Write-host "Current Internal Value: " $value.internalhostname 
    	Write-host "New Internal Value:     " $base
	   	[string]$set = Read-host $ConfirmPrompt 
    	write-host "" 

    	if ($set -eq "Y")
		{ 
	        Set-OutlookAnywhere -identity $value.identity -InternalHostname $base -InternalClientsRequireSsl $true
		}
		else
		{ 
	    	write-host "Outlook Anywhere internal hostname value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    	} 

    	Write-host "Current External Value: " $value.externalhostname 
    	Write-host "New External Value:     " $base
    	[string]$set = Read-host $ConfirmPrompt 
    	write-host "" 

    	if ($set -eq "Y")
		{ 
     		Set-OutlookAnywhere -identity $value.identity -ExternalHostname $base -ExternalClientsRequireSsl $true -ExternalClientAuthenticationMethod Negotiate -IISAuthenticationMethods @('Ntlm', 'Basic', 'Negotiate')
    	}
		else
		{ 
       		write-host "Outlook Anywhere external hostname value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
    	} 
	}
}

# =========================================================
# Build the Autodiscover internal URL and set the SCP Value
#
# Autodiscover is set for all CAS server at once
#

[string]$SCPset = Read-host "Do you want to use the same FQDN for Autodiscover Internal URI attrubute (SCP) for all of your CAS servers? (Y/N)" 
Write-host "" 

if ($SCPset -eq "Y")    { 
    Write-host "Setting Autodiscover Service Connection Point" -foregroundcolor Yellow 
    write-host "" 

    $SCPURL = "https://" + $base + $SCPExtend 

	if (($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)-and ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Minor -ge 1))
	{
        [array]$SCPCurrent = Get-ClientAccessService 
	}
	else
	{
        [array]$SCPCurrent = Get-ClientAccessServer 
	}
	
    Foreach ($value in $SCPCurrent) { 
        Write-host "Looking at Server: " $value.name 
        Write-host "Current SCP value: " $value.AutoDiscoverServiceInternalUri.absoluteuri 
        Write-host "New SCP Value:     " $SCPURL 
        [string]$set = Read-host $ConfirmPrompt 
        write-host "" 
        if ($set -eq "Y")    { 
			if (($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Major -eq 15)-and ($ExOrgCfg.AdminDisplayVersion.ExchangeBuild.Minor -eq 1))
			{
            Set-ClientAccessService -id $value.identity -AutoDiscoverServiceInternalUri $SCPURL 
			}
		else
			{
            Set-ClientAccessServer -id $value.identity -AutoDiscoverServiceInternalUri $SCPURL 
			}
        }    else { 
            write-host "Autodiscover Service Connection Point internal value NOT changed" -foregroundcolor $NoChangeForeground -backgroundcolor $NoChangeBackground 
        } 
    } 
} 
