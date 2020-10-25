#
# Script to change disabled Echange related Services Startup Type to proper value
#
# By: Konrad Sagala
#
#
# Version 1.0
#

 
# path to csv file with Exchange services list
$file = "C:\Scripts\exchangeservices.csv"

$services = Import-Csv -Path $file

# Change services state according to file

foreach ($service in $services) {
    $Name = $service.SrvName
    $filter = "Name = ""$Name"""
    $srv = Get-WmiObject -Class Win32_Service -Filter $filter
    if ($srv.StartMode -eq 'Disabled')
	    {
	    Set-Service $service.SrvName -Startup $service.Mode
	    }
}


#Restart-Computer -Confirm

# end of the script