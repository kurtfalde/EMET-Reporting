<#  Setup script for Subscription Server for EMET Events and reporting

Microsoft Services - Kurt Falde

#>

cd C:\EMET

#Setting WinRM Service to automatic start and running quickconfig
Set-Service -Name winrm -StartupType Automatic
winrm quickconfig -quiet

#Set the size of the forwarded events log to 500MB
wevtutil sl forwardedevents /ms:500000000


#Running quickconfig for subscription service
wecutil qc -quiet

#Creating Applocker Subscription from XML files FYI we do delete any existing ones and recreate
If ((wecutil gs "EMET Events") -ne $NULL) {
    wecutil ds "EMET Events"
    wecutil cs .\EMETSubscription.xml
    }
Else {wecutil cs .\EMETSubscription.xml} 


#FYI if you need to export Subscriptions to fix SIDS or anything in an environment 
#use wecutil gs "%subscriptionname%" /f:xml >>"C:\Temp\%subscriptionname%.xml"

#Creating Task Scheduler Item to restart parsing script on reboot of system.
If ((Get-ScheduledTask -TaskName "EMET Parsing Task") -ne $NULL) {
    Unregister-ScheduledTask -TaskName "EMET Parsing Task" -Confir:$false
    schtasks.exe /create /tn "EMET Parsing Task" /xml EMETParsingTask.xml
    }
Else {schtasks.exe /create /tn "EMET Parsing Task" /xml EMETParsingTask.xml}