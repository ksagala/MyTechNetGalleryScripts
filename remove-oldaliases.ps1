# 
# Script removes unused mail aliases 
# (example removes all X.400 aliases) 
# you can also remove CCMAIL, MSMAIL, notes, or simple SMTP alias (e.g. *@domain.old)
# for all distribution lists and mailboxes in Exchange organization 
# OF course you should earlier remove specific domain 
# from e-mail policies 
# Author: Konrad Sagala 
# 
# version 1.1 updated for Exchange 2010 and Exchange 2013
#

$i = 0 
$dllist = get-distributiongroup -resultsize unlimited
$dllist | ForEach-Object {  
    $maile = $_ 
    for ($i=0; $i -lt $_.emailaddresses.count; $i++) 
        { 
            if ($_.emailaddresses[$i] -like "X.400:*") 
                   { 
                $badaddress = $_.emailaddresses[$i] 
                $maile.emailaddresses.Remove($badaddress)
                $maile | Set-DistributionGroup -EmailAddresses $maile.emailaddresses
                } 
        } 
}
$i = 0 
$mailboxlist = get-mailbox -resultsize unlimited
$mailboxlist | ForEach-Object {  
    $maile = $_ 
    for ($i=0; $i -lt $_.emailaddresses.count; $i++) 
        { 
            if ($_.emailaddresses[$i] -like "X.400:*") 
                   { 
                $badaddress = $_.emailaddresses[$i] 
                $maile.emailaddresses.Remove($badaddress)
                $maile | Set-Mailbox -EmailAddresses $maile.emailaddresses
                } 
        } 
}
