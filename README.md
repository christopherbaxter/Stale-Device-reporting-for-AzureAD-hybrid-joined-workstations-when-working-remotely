# Stale Computer Accounts
Managing stale workstations when most users are working from home using AzureAD in a Hybrid configuration is a significant challenge. Here is some code that will help.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes, I would not advise copying my permissions. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what work for now. 

![](https://private-user-images.githubusercontent.com/47409333/293186391-375ba462-33ac-4516-bad6-d1d8cb02f018.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI0MzksIm5iZiI6MTcwMzc2MjEzOSwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzOTEtMzc1YmE0NjItMzNhYy00NTE2LWJhZDYtZDFkOGNiMDJmMDE4LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMTUzOVomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTNjNDk0NTY4MjJkYzAxZDE1MWZjNjE0OGY1NTUzM2IyNGU4MGI2MDg1MmQzYzEyZTQ0MDg1NmE0MWI0MDZmYmMmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.L4FEjRuMorxQh-7w_8qEaFIvx2Zp7MHPZYoBdo7LF4k)

I also elevate my AzureAD account to 'Intune Administrator', 'Cloud Device Administrator' and 'Security Reader'. These permissions also feel more than needed. Understand that I work in a very large environment, that is very fast paced, so I elevate these as I need them for other tasks as well.

You will need to make sure that you have the following PowerShell modules installed. There is a lot to consider with these modules as some cannot run with others. This was a bit of a learning curve. 

ActiveDirectory\
AzureAD\
ImportExcel\
JoinModule\
MSAL.PS\
PSReadline (May not be needed, not tested without this)

Ultimately, I built a VM on-prem in one of our data centres to run this script, including others. My machine has 4 procs and 16Gb RAM, the reason for an on-prem VM is because most of our workforce is working from home (me included), and running this script is a little slow through the VPN. Our ExpressRoute also makes this data collection significantly more efficient. In a small environment, you will not need this.

# Disclaimer

Ok, so my code may not be very pretty, or efficient in terms of coding. I have only been scripting with PowerShell since September 2020, have had very little (if any), formal PowerShell training and have no previous scripting experience to speak of, apart from the '1 liners' that AD engineers normally create, so please, go easy. I have found that I LOVE PowerShell and finding strange solutions like this have become a passion for me.

## Christopher, enough ramble, How does this thing work?

Before we start, I have expected 'runtimes' for each section. This is for my environment, and will not be accurate for your environment. Use the Measure-Command cmdlet to measure for your specific environment. I added this in because the script could run for hours and appear to be doing nothing.

### Parameters

The first section is where we supply the TenantID (Of the AzureAD tenant) and the ClientID of the Service Principal you have created. If you populate these (hard code), then the script will not ask for these and will immediately go to the Authentication process.

![](https://private-user-images.githubusercontent.com/47409333/293186384-7d7ad056-472e-44e6-9551-3abfacb0806c.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzODQtN2Q3YWQwNTYtNDcyZS00NGU2LTk1NTEtM2FiZmFjYjA4MDZjLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTQ2NzI1MGIxNWViMzcwYjVkN2RlNTgxNTU1ZDRmY2ZlZjZjNjZkMmNkNjUxZWEwZGY5N2ZkNWI0Mzg0MzcwOWUmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.UYh1fhchIl1HqDWkSpp9gveunB_6WHzHgv49cErfsIE)

### Functions

The functions needed by the script are included in the script. I have modified the 'Invoke-MSGraphOperation' function significantly. I was running into issues with the token and renewing it. I also noted some of the errors went away with a retry or 2, so I built this into the function. Sorry @JankeSkanke @NickolajA for hacking at your work. :-)

![](https://private-user-images.githubusercontent.com/47409333/293195989-fd9033f5-fead-4d8f-927e-06dd556f6ed0.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxOTU5ODktZmQ5MDMzZjUtZmVhZC00ZDhmLTkyN2UtMDZkZDU1NmY2ZWQwLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTllYmFiMTBjYjIxMDM0NWE2ODRkODRmYWNlZGVjZmMxNGFjOGE1NmRlODkzZThhYTRmYjhlMGFkZjhiM2ZhY2EmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.AqmhhMxUSBdbg0xPIzWMsVixFiAJNr2wccsxg2JMFSI)

### The Variables

The variables get set here. I have a need to upload the report for another team to use for another report. Enable these and you will be able to do the same.

![](https://private-user-images.githubusercontent.com/47409333/293186395-30163089-3f48-441d-b15f-fbf70e03ca97.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzOTUtMzAxNjMwODktM2Y0OC00NDFkLWIxNWYtZmJmNzBlMDNjYTk3LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTUxMGVmMDQ4ZGFmMTdlODQ2YWU3ODViMGI4MjAzOTUxMDY5ODcwNDQwMzYyZDA0ZTFkZTA4YTE2ODg0Y2ViYTImWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.TzLbqD4z5dpiF6aG34GZYFkjYURbM9egBRF_S3OzWAI)

The variable section also has a section to use the system proxy. I was having trouble with the proxy, intermittently. Adding these lines solved the problem

### The initial Authentication and Token Acquisition

Ok, so now the 'fun' starts.

The authentication and token acquisition will allow for auth with MFA. You will notice in the script that I have these commands running a few times in the script. This allows for token renewal without requiring MFA again. I also ran into some strange issues with different MS Graph API resources, where a token used for one resource, could not be used on the next resource, this corrects this issue, no idea why, never dug too deep into it because I needed it to work, not be pretty. :-)

![](https://private-user-images.githubusercontent.com/47409333/293186356-55cee05c-d5ef-4338-b869-937b7fe8ffcf.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNTYtNTVjZWUwNWMtZDVlZi00MzM4LWI4NjktOTM3YjdmZThmZmNmLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTgzYTI0ZjA2MTBhMWM5ODFhYjliZWFhYmExM2U5ZjMyZTgyMDkwMzhhOGE0MTNhNmEyMjk5YTAzY2I2MWVhMjUmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.am_OpMG_jmDYNrtFGI-wiuf3ntx6fGZwTvjvDfvZvlw)

### AzureAD Device Data Extraction

This section also requires an authentication process and will allow for MFA. The reason why I added this in here is that the script takes a long time to run in my environment, and so, if I perform this extraction first, without the initial auth\token process, the script will complete this process, then sit waiting for auth and MFA, and in essence, not run. Same if this was moved to after the MS Graph extractions. 

Having the 'authy' bits in this order, the script will ask for auth and MFA for MS Graph, then auth and MFA for AzureAD, one after the other with no delay, allowing the script to run without manual intervention. 

    $AzureADDevices = [System.Collections.ArrayList]@(Get-AzureADDevice -All:$TRUE | Where-Object { $_.DeviceOSType -like "*Windows*" } | 
    Select-Object @{Name = "AzureADDeviceID"; Expression = { $_.DeviceId.toString() } }, @{Name = "ObjectID"; Expression = { $_.ObjectID.toString() } }, 
    @{Name = "AADEnabled"; Expression = { $_.AccountEnabled } }, 
    @{Name = "AADApproximateLastLogonTimeStamp"; Expression = { (Get-Date -Date $_.ApproximateLastLogonTimeStamp -Format 'yyyy/MM/dd HH:mm') } }, 
    @{Name = "AADDisplayName"; Expression = { $_.DisplayName } }, 
    @{Name = "AADSTALE"; Expression = { if ($_.ApproximateLastLogonTimeStamp -le $StaleDate) { "TRUE" } elseif ($_.ApproximateLastLogonTimeStamp -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } } | Sort-Object azureADDeviceId )    

I extract the data into an ArrayList. This was needed for a previous 'join' function, I left it like this because I noted strange errors in other scripts. I never had the time to validate that this is the case here, so I simply left it in place. At some point, I would like to test other array types and test processing time between them, not now, this works exactly as needed.

I have tried to split this code up a little to allow for easier 'reading', and hopefully understanding.
You will note that I transform the data in the extract and convert 'DeviceId' and 'ObjectId' to string values and call them 'AzureADDeviceID' and 'ObjectId' respectively. The string values allow for easier processing later on in the script. 
I convert 'ApproximateLastLogonTimeStamp' to a date in a format that is uniform for what I need, and call it 'AADApproximateLastLogonTimeStamp'.
I then extract 'DisplayName' and call it 'AADDisplayName'.

#### Now, for the key to the reporting. 

I create a field in the array called 'AADStale', this is a calculation from the 'ApproximateLastLogonTimeStamp'. If the date is 'less than' (older than) the '$StaleDate = (Get-Date).AddDays(-90)' variable (current date minus 90 days in this case), the value is "TRUE" (Stale), else, if the date is 'greater than' (younger than) the '$StaleDate' variable, the value is then "FALSE" (not stale). If there is no date in the data, then the value is "NoLoginDateFound"

You will notice that I sort the data. This is needed to speed up the 'join' processes later on.

![](https://private-user-images.githubusercontent.com/47409333/293186349-375343c1-b934-4a84-9595-356a8c0c9e5e.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNDktMzc1MzQzYzEtYjkzNC00YTg0LTk1OTUtMzU2YThjMGM5ZTVlLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPWFmMDE3ODc4MWI4OGM2Y2RmNzA0MjQxNWY2OTU0ZmEyZjE3YmJiYjY0YTU5NDNjNTViM2FmODZkODhhMmZiNDUmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.vEoLOyOx4-nfbqZoG8N2liKeuSGczgsuKMPTNde4pAc)

You will notice a hashed out export line, as well as a resource cleanup (Remove-Variable and Clear-ResourceEnvironment). This is to serve 2 purposes, and is included in most of the sections. 

1. Allow for faster troubleshooting of the code (In my environment, the data extraction can take hours, and with a failure, this will mean that I will be waiting, a lot). Enable this to dump the file to the directory of your choice. It is not wise to leave them here.
2. Depending on the amount of data being extracted, you may run short of RAM. This will free that RAM. Also, PowerShell seems to take a beating in terms of performance if there is a LOT of data, this prevents this performance degradation. Use this at your own discretion.

### Intune Device Data Extraction

You will notice here that I refresh the token for MS Graph extraction. This should not ask for auth or MFA again as we are simply renewing a current token. Without this section, data extraction fails. Nothing really fancy here, apart from the data transformation perhaps. I also sort the data for use later.

    $IntuneInterimArray = [System.Collections.ArrayList]@(Invoke-MSGraphOperation -Get -APIVersion "Beta" -Resource "deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'" -Headers $AuthenticationHeader -Verbose | 
    Where-Object { $_.azureADDeviceId -ne "00000000-0000-0000-0000-000000000000" } | 
    Select-Object @{Name = "azureADDeviceId"; Expression = { $_.azureADDeviceId.toString() } }, 
    @{Name = "IntuneDeviceID"; Expression = { $_.id.ToString() } }, 
    @{Name = "MSGraphDeviceName"; Expression = { $_.deviceName } }, 
    @{Name = "enrolledDateTime"; Expression = { (Get-Date -Date $_.enrolledDateTime -Format "yyyy/MM/dd HH:mm") } }, 
    @{Name = "MSGraphlastSyncDateTime"; Expression = { (Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") } }, 
    @{Name = "MSGraphLastSyncStale"; Expression = { if ((Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") -le $StaleDate) { "TRUE" } elseif ((Get-Date -Date $_.lastSyncDateTime -Format "yyyy/MM/dd HH:mm") -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, @{Name = "UserUPN"; Expression = { $_.userPrincipalName } } | Sort-Object IntuneDeviceID)    

Ok, so much the same with the export from the MSGraph API for the devices in Intune (I use "beta' because the attribute names differ between 'Beta' and 'v1.0' in some cases, this makes it quicker for me to share code between scripts, this can easily be convferted to 'v1.0' if the data is available, I have not checked).

The one exception in this section of code is that I had to convert the dates in the calculation, it would likely be quicker to convert the date once, then calculate on that, this could be tested in another iteration of the script. I then sort the data on 'IntuneDeviceID'

![](https://private-user-images.githubusercontent.com/47409333/293186353-8d01e400-694f-4de1-ba1e-6e7f24b16196.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNTMtOGQwMWU0MDAtNjk0Zi00ZGUxLWJhMWUtNmU3ZjI0YjE2MTk2LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPWQ5MzJmNGExYTNhYzNjMmVjNDUxZDMxZjI1YmRmOTYzYmY5ZGVhOTVhYjIzODZjNWQzM2E5ZTQ1YTdlNmM0YmMmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.XEXUF7CFXibLPO4MvjsQQvj2jUYRN0hFunEw-a1mUmc)

### On-Prem AD Data Extraction

This script has been written to extract all the details for all the Windows 10 devices in an AD forest. If you are wanting to specify only a specific domain, you will need to edit the 'Variables' section.

I added a section to specify a specific domain controller to target if needed. The code is pretty cool, so I left it here.

![](https://private-user-images.githubusercontent.com/47409333/293186336-e548b7b9-41c8-44dd-96a0-fb9c29db18a4.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjIxNjcsIm5iZiI6MTcwMzc2MTg2NywicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzMzYtZTU0OGI3YjktNDFjOC00NGRkLTk2YTAtZmI5YzI5ZGIxOGE0LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMTEwN1omWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPWVjZWIxYmMzOGVjZGM0ZjZjMzFkMzI2NDM3ODVhZDczMGFmMGRiNjc5ODY0YzAyMmFjZTg4NmU0MDJkODcwYTkmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.qgjFDjVseZOVa_PfZ_mtfmcVZkAjw0vQq_F4tmCkSw8)

![](https://private-user-images.githubusercontent.com/47409333/293186368-14d38b47-a2da-4e7c-94b2-17f32e45cbb5.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNjgtMTRkMzhiNDctYTJkYS00ZTdjLTk0YjItMTdmMzJlNDVjYmI1LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTUyZThiMGE5ZjgxNjA1MDRiYTNhNjY0ZjE5ZDViZGY2NDFmY2YwNzRjNTFiMTFkNjBlNTNjYmMxYTI1ZjRkZjkmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.NbimU4tqVXXviNis2I9qDAg6Bukg9hS7875hRgXE9Lo)

I had a number of issues with the extraction with timeouts for some reason. I assume this is some strange network latency or something similar. One needs to remember to pick ones battles.

![](https://private-user-images.githubusercontent.com/47409333/293186372-60dd21c0-ab66-4c5f-ad94-6d2528422594.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNzItNjBkZDIxYzAtYWI2Ni00YzVmLWFkOTQtNmQyNTI4NDIyNTk0LmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPWM2ZTVhYTFmNmRmZTk4MDY5ZmFlMzFmMmMyNDJmNWZiNjZhYjcxMWYzMzIyYjJjOWFjYzU0ZGY5Y2RmNTEwNTEmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.DdWvC3AUa6LJWgeanfK3GRtDzEIOb7J8VrH9lvh3XdE)

This has proven to be pretty reliable, thankfully.

#### Here is the on-prem data processing section

    $AllOPCompsArray = [System.Collections.ArrayList]@($RAWAllComps | Select-Object 
    @{Name = "azureADDeviceId"; Expression = { $_.objectGUID.toString() } }, 
    @{Name = "OPDeviceName"; Expression = { $_.CN } }, 
    @{Name = "OPDeviceFQDN"; Expression = { "$($_.CN).$($_.CanonicalName.Split('/')[0])" } }, 
    @{Name = "SourceDomain"; Expression = { "$($_.CanonicalName.Split('/')[0])" } }, 
    @{Name = "OPLastLogonTS"; Expression = { (Get-Date -Date $_.LastLogonDate -Format "yyyy/MM/dd HH:mm") } }, 
    @{Name = "OPSTALE"; Expression = { if ($_.LastLogonDate -le $StaleDate) { "TRUE" } elseif ($_.LastLogonDate -gt $StaleDate) { "FALSE" } else { "NoLoginDateFound" } } }, @{Name = "OPEnabled"; Expression = { $_.Enabled } } | Sort-Object azureADDeviceId )

I noticed that in our estate, the on-prem 'objectGUID' matches the 'AzureADDeviceID', so, this exists. I also specify it as a string value.
As Microsoft (as of the time of writing) has not provided an easily accessible method of finding the source domain data in AzureAD to enable splitting of the machines by source domain, so I had to create one in my reporting, so I create a field in the array called "OPDeviceFQDN", which is a calculated value using the 'CN' and the 'CanonicalName', but split and using only the first object, with a '.' in the middle. The "SourceDomain" field is much the same, but without the machine name ('CN').
The "OPLastLogonTS" is the 'LastLogonDate'
The "OPSTALE" section is a calculation again, much the same as the other calculations from Intune\AzureAD.

I sort the data by "azureADDeviceId".

![](https://private-user-images.githubusercontent.com/47409333/293186374-b460db4a-454d-44d1-a68b-8b73eea29b82.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNzQtYjQ2MGRiNGEtNDU0ZC00NGQxLWE2OGItOGI3M2VlYTI5YjgyLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPWU5NWVhY2U1NTJhMTg5MzE0MmYxNGEzNTRjZTRkYWMyMjA2NjMyMDk2ZjY5YWViNWZmYzg1NTdmYTEzNzZhYWQmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.AhrisRbX1qAdCLypPL49EOmMWXMGI50PJ8LVt5JOgwQ)

### Blending On-Prem AD Data with Intune Data

Here you will see that there is a section that if enabled, will import the exported data from the previous extractions, if the relevent export is enabled above. If you are testing, this section will test for the existance of the data in memory, if not present, will import from the 'interim' file\s.

![](https://private-user-images.githubusercontent.com/47409333/293186376-815b2597-a280-44ef-acc8-c116059b6360.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNzYtODE1YjI1OTctYTI4MC00NGVmLWFjYzgtYzExNjA1OWI2MzYwLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTk0M2QwODVkMzA3NTgwNzQ0MjUzM2ExZTJlODU0NDhkMmYzMDZjODdhZWRmNzA1MmY4MzcwOTNkNjA4YWMwYjgmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.gWi6WWwjvrSewKX8RnUoB6jQ_BaHEpd8OvJaIJqRgR8)

Below I use a term 'blend'. I dont realy know what to call this. It is similar to a SQL join function apparently. What is does is takes the 2 arrays, containing very different information, and joins the data based on a specific field, that is present in both arrays. I sort the data in the arrays in an attempt to speed up processing, I dont remember testing the performance though, it may not be needed. The sorting takes seconds so, I left it in.

Now things start to get interesting. The script in this section 'blends' the previously extracted data. The data is matched using the 'objectGUID' (now 'azureADDeviceID') from the on-prem data extraction with the 'AzureADDeviceID' from the Intune extract. Interestingly, the on-prem AD 'objectGUID' and the 'AzureADDeviceID' is the same (as stated earlier). At least if the devices are Hybrid joined. I am unable to comment on other environments though. Your mileage may vary.

I 'blend' the data in both 'directions' (first using the Intune extract as the left array and the on-prem as the right array then swopping the arrays). I noted that I got different numbers (object counts) so, for completeness, this process was born. This also duplicates a lot of the data, but ensures that the data is as complete as possible.

### Data Deduplication and AzureAD Data Blending

In this section, I deduplicate the data, then 'blend' the AzureAD device data with the previously 'blended' data. I deduplicate the data first before blending the next lot of data. 

![](https://private-user-images.githubusercontent.com/47409333/293186362-09bacbd2-7d45-49d6-b339-09dfbb0fc9f2.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzNjItMDliYWNiZDItN2Q0NS00OWQ2LWIzMzktMDlkZmJiMGZjOWYyLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTAyZjVhODUwOGVmMjg2ZWNjYThhZDA5MzRkZDVkODNmNTBmOWQyMmZmMzY4ZjQ3ZjU0ZjgyOGUzNDEzYTJiYWImWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.Ws4eiY-jM0MJ2UqPB_nrB_sX6vf7LRvegSSDQXQxj4I)

### Report Export - This is where the Magic happens!!!

This little section, is where the magic happens. This section will do the calculation on the OPStale, AADStale and MSGraphLastSyncStale fields. These are calculated fields higher up in the script. If a device is stale on-prem (likely if working remotely), but not in AzureAD, then the device is **NOT** stale\dormant. If the device is not matched to an AzureAD object, then the device **IS** classified as stale\dormant. In the same way, if the device is classified as stale\dormant in AzureAD, and not in on-prem AD, the device is **NOT** stale\dormant. If the AzureAD device is stale in in AzureAD but the device is not matched to an on-prem object, the device **IS** stale.

This is what the code looks like:

This is a snippet of the code on line 575. This code it what does all the analysis of the computer accounts: 

    $AllDevices = [System.Collections.ArrayList]@($DDAllDevProcArray | Select-Object AzureADDeviceID, IntuneDeviceID, ObjectID, AADDisplayName, MSGraphDeviceName, OPDeviceName, OPDeviceFQDN, SourceDomain, UserUPN, enrolledDateTime, AADApproximateLastLogonTimeStamp, MSGraphlastSyncDateTime, OPLastLogonTS, AADEnabled, OPEnabled, AADSTALE, OPSTALE, MSGraphLastSyncStale, 
    @{Name = "TrueStale"; Expression = { if ($_.AADStale -notlike "False" -and $_.OPStale -notlike "False" -and $_.MSGraphLastSyncStale -notlike "False") { "TRUE" } else { "FALSE" } } }, 
    @{Name = "AccountEnabled"; Expression = { if ($_.AADEnabled -notlike "False" -and $_.OPEnabled -notlike "False") { "TRUE" } else { "FALSE" } } })

The calculated field "TrueStale" will show "TRUE" only if the fields "AADSTALE", "OPSTALE" and "MSGraphLastSyncStale" are specifically "False", else this will show the device as truly stale. This means that if any of the fields are anything other than "False", like "TRUE", "NoLoginDateFound" or $null will be classified as "TrueStale".

I then, for dexterity, created a calculated field called "AccountEnabled". This field will only show as "TRUE" if both "AADEnabled" and "OPEnabled" are not "False" (Same rule as "TrueStale" above).

The export will export all devices in the report, both stale and active. This is easily switched. The code is in the script. There is also the 'remote' export if you would like to send the extract to another server\share.

![](https://private-user-images.githubusercontent.com/47409333/293186388-a189b4e4-4b1b-4484-aca7-0505d3cdff51.jpg?jwt=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJnaXRodWIuY29tIiwiYXVkIjoicmF3LmdpdGh1YnVzZXJjb250ZW50LmNvbSIsImtleSI6ImtleTEiLCJleHAiOjE3MDM3NjI4MzgsIm5iZiI6MTcwMzc2MjUzOCwicGF0aCI6Ii80NzQwOTMzMy8yOTMxODYzODgtYTE4OWI0ZTQtNGIxYi00NDg0LWFjYTctMDUwNWQzY2RmZjUxLmpwZz9YLUFtei1BbGdvcml0aG09QVdTNC1ITUFDLVNIQTI1NiZYLUFtei1DcmVkZW50aWFsPUFLSUFJV05KWUFYNENTVkVINTNBJTJGMjAyMzEyMjglMkZ1cy1lYXN0LTElMkZzMyUyRmF3czRfcmVxdWVzdCZYLUFtei1EYXRlPTIwMjMxMjI4VDExMjIxOFomWC1BbXotRXhwaXJlcz0zMDAmWC1BbXotU2lnbmF0dXJlPTVjNDQ2MmNlZDc1YzZjMjkwZDQ2MjkwM2M3MTJhODk1MGFkZjc4MjA0Mzk5MmRlZGU3MDVkMTA0NjBkNDhkZTkmWC1BbXotU2lnbmVkSGVhZGVycz1ob3N0JmFjdG9yX2lkPTAma2V5X2lkPTAmcmVwb19pZD0wIn0.L8MLLVObNX-R2Q3wGbW0wW7Kp7TgkMK1j6i_B7FT1uc)

I have uploaded a SAMPLE report output. Please take a look at it.

### Whats Next?

I have started work on the Disablement script, I am finding it very difficult to find the time to dig into this at the moment though. Business as Usual work takes precidence over all else, so this is on the backburner, hoping to spend a little time on this in the next few weeks.

Visitor Counter (Excuse the font, I have no idea what I'm doing)\
![Visitor Count](https://profile-counter.glitch.me/ReportOnStaleWorkstationsUsingAzureADandOnPrem/count.svg)
