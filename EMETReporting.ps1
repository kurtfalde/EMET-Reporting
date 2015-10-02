<#
    This Sample Code is provided for the purpose of illustration only and is not 
    intended to be used in a production environment.  THIS SAMPLE CODE AND ANY 
    RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
    EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
    MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a 
    nonexclusive, royalty-free right to use and modify the Sample Code and to 
    reproduce and distribute the object code form of the Sample Code, provided 
    that You agree: 
    (i)            to not use Our name, logo, or trademarks to market Your 
                    software product in which the Sample Code is embedded; 
    (ii)           to include a valid copyright notice on Your software product 
                    in which the Sample Code is embedded; and 
    (iii)          to indemnify, hold harmless, and defend Us and Our suppliers 
                    from and against any claims or lawsuits, including attorneys’ 
                    fees, that arise or result from the use or distribution of 
                    the Sample Code.


    Originally written by Kurt Falde
    Update Kurt Falde 10/17/2014 - Added System.Core Assembly load for older systems
    Update Kurt Falde 10/17/2014 - Added ASR Mitigation (Regex/Matching/Output to CSV)
    Update Kurt Falde 2/18/2015 - Adding logic to check for existing event subscriber and jobs and stop them
    Creating Version and Guidance to publish on GitHub

    This script makes use of the pseventlogwatcher module which can be obtained from http://pseventlogwatcher.codeplex.com/

#>


$workingdir = "c:\emet"

cd $workingdir

#The following line was needed on a 2008 R2 setup in one case
Add-Type -AssemblyName System.Core

#Following added to unregister any existing EMETWatcher Event in case script has already been ran
Unregister-Event EMETWatcher -ErrorAction SilentlyContinue
Remove-Job EMETWatcher -ErrorAction SilentlyContinue


Import-Module .\EventLogWatcher.psm1

$BookmarkToStartFrom = Get-BookmarkToStartFrom

$EventLogQuery = New-EventLogQuery "ForwardedEvents" -Query "*[System[Provider[@Name='EMET'] and (EventID=1 or EventID=2)]]"

$EventLogWatcher = New-EventLogWatcher $EventLogQuery $BookmarkToStartFrom 

$action = {     write-host "Performing Regex and extracting to csv"
                $outfile = "c:\emet\emet.csv"
                write-host $outfile
                $RegexMitigation = "(?<=detected ).*?(?= mitigation)"
                $RegexApplication = "(?<=Application 	: ).*?(.exe=?)"
                $RegexUsername = "(?<=User Name 	: ).*?(?=`n)"
                $RegexModule = "(?<=Module 	: ).*?(?=`n)"
                $RegexDLL = "(?<=DllName 	: ).*?(?=`n)"
                $RegexURL = "(?<=Web address 	: ).*?(?=`n)"
                $RegexZone = "(?<=Url zone 	: ).*?(?=`n)"
                $EventData = $EventRecordXML.Event.EventData.Data

                $EventObj = New-Object psobject
                $EventObj | Add-Member noteproperty EventDate $EventRecord.TimeCreated
                write-host $EventObj
                $EventObj | Add-Member noteproperty EventHost $EventRecord.MachineName
                write-host $EventObj

                $EventMitigationMatch = $EventData -match $RegexMitigation
                $EventMitigation = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventMitigation $EventMitigation
                write-host $EventObj

                $EventApplicationMatch = $EventData -match $RegexApplication
                $EventApplication = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventApplication $EventApplication
                write-host $EventObj

                $EventUsernameMatch = $EventData -match $RegexUsername
                $EventUsername = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventUsername $EventUsername
                write-host $EventObj

                $EventModuleMatch = $EventData -match $RegexModule
                $EventModule = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventModule $EventModule
                write-host $EventObj

                $EventDLLMatch = $EventData -match $RegexDLL
                $EventDLL = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventDLL $EventDLL
                write-host $EventObj

                $EventURLMatch = $EventData -match $RegexURL
                $EventURL = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventURL $EventURL
                write-host $EventObj

                $EventZoneMatch = $EventData -match $RegexZone
                $EventZone = $Matches[0]
                if ($Matches) { $Matches.Clear() }
                $EventObj | Add-Member noteproperty EventZone $EventZone
                write-host $EventObj

                
               If ($Outfile -ne $Null)
            {
                write-host $Outfile
                $EventObj | Convertto-CSV -Outvariable OutData -NoTypeInformation 
                
                $OutPath = $Outfile
                write-host $OutPath
                If (Test-Path $OutPath)
                {
                    $Outdata[1..($Outdata.count - 1)] | ForEach-Object {Out-File -InputObject $_ $OutPath -append default}
                } else {
                    Out-File -InputObject $Outdata $OutPath -Encoding default
                }
            }


            }

Register-EventRecordWrittenEvent $EventLogWatcher -action $action -SourceIdentifier EMETWatcher

$EventLogWatcher.Enabled = $True