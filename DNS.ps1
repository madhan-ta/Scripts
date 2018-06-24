Get-WmiObject -Class win32_networkadapterconfiguration -Filter ipenabled=true | Format-Table -Property ipaddress
Get-WmiObject Win32_NetworkAdapterConfiguration -computername . | select name, DNSServerSearchOrder
# get ping status
(get-wmiobject win32_pingstatus -filter "address='192.168.0.152'").statuscode 