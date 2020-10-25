#
# Simple script to list all mobile devices assosiated to all mailboxes in Exchange organization
# by Konrad Sagala
#
$ActiveSyncUsers = Get-CASMailbox -filter {HasActiveSyncDevicePartnership -eq $True}
write-host "Number of users with mobile devices: "($ActiveSyncUsers).Count
$ActiveSyncUsers| Select-Object displayname,@{Expression={(Get-MobileDevice -Mailbox $_.Identity).DeviceType};Name="Devices"}