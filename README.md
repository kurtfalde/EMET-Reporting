# EMET-Reporting
Various Scripts and Guidance on creating Reporting for the EMET Client

The base reporting solution listed here utilizes Windows Event Forwarding to first centralize EMET Detection Events.
The Events are then parsed using a script which uses the pseventlogwatcher module from http://pseventlogwatcher.codeplex.com/
There is some regex that parses the events and extracts fields of interest and appends them to a .csv file
Also included is a .pbix file which is for the PowerBI Desktop client which can be obtained from https://powerbi.microsoft.com/en-us/desktop
The EMET Dashboard can either be used locally via the PowerBI Desktop client or with the use of a PowerBI account the onpremise data can 
be synced to the PowerBI service with the PowerBI Personal Gateway giving an online dashboard usable either via browser or the various
apps for the PowerBI Service IOS/Android/Windows.

Also intend to add some queries that are useful in a Splunk environment for querying EMET events if they are immported directly
to Splunk.
