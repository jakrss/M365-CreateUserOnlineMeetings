[CmdletBinding()]
    param (
        [parameter(Mandatory = $False,
        Position=0,
        HelpMessage="Do we delete created meetings? WARNING: THIS WILL DELETE ALL CREATED MEETINGS FROM THE CURRENT YEAR")]
        [Alias("Delete")]
        [Switch] $CleanUp
    )

Import-Module Microsoft.Graph.CloudCommunications

#### Permissions required: 
# OnlineMeetings.ReadWrite.All    = Ability to create online meetings for users. Also requires CS Policy https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy (ParentTeacherMeetings-AppPolicy)
# GroupMember.Read.All            = Reads all users from Azure groups
# User.Read.All                   = Ability to read user attributes in Azure. This is only required as we are reading additional properties such as OnPremisesSamAccountName, otherwise you can use User.ReadBasic.All for only basic account info.

#https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy
#New-CsApplicationAccessPolicy -Identity ParentTeacherMeetings-AppPolicy -AppIds "application-guid" -Description "This Policy allows the enterprise app access to every user"
#Grant-CsApplicationAccessPolicy -PolicyName ParentTeacherMeetings-AppPolicy -Global

##
## Constants 
##
$ApplicationID = "application-guid" 
$TenatDomainName = "tenant.onmicrosoft.com"
$AccessSecret = "application-access-secret"
$AADGroups=@(
    "aad-group-guid" #Target Azure AD Group GUID
    "aad-group-guid" #Target Azure AD Group GUID
)

# Variables
$CurrentYear =  $(Get-Date).Year
[System.DateTime]$MeetingStartTime =[System.DateTime]::Parse("$($CurrentYear)-01-01T05:00:00+08:00")
[System.DateTime]$MeetingEndTime = [System.DateTime]::Parse("$($CurrentYear)-12-31T09:30:00+08:00")

# Meeting Subject Prefix, Users name will be appended
$MeetingSubject = "Meeting Subject"
$CurrentList = "./Exports/Current-MeetingsList.csv"
$RunningList = "./Exports/$($CurrentYear)-MeetingsList.csv"


function main {
    # Get a token using our App Registration
    $token = Get-GraphToken "client_credentials" "https://graph.microsoft.com/.default" $ApplicationID $AccessSecret $TenatDomainName
    # Connect to Graph with our Token
    Connect-MgGraph -AccessToken $token
    
    if ($CleanUp) {
        #Call functions
        Invoke-Cleanup
    }

    # Get list of users from the defined AAD groups
    $UserList = Get-TeachersList -GroupIDs $AADGroups
   
    # Reporting on any duplicate entries
    #Write-Host "Found $(($UserList | Group-Object -Property Id | Where-Object { $_.Count -gt 1 }).Count) duplicated IDs"

    # Loop through all users found in AAD Groups
    foreach ($user in $UserList){
        Write-Host "Processing $($user.displayname)"
        $onPremSam = Get-MgUser -UserId $user.id -Property OnPremisesSamAccountName
        $meeting = Create-OnlineMeeting -User $user -MeetingSubject $MeetingSubject -StartDateTime $MeetingStartTime -EndDateTime $MeetingEndTime
        $user | Add-Member -MemberType NoteProperty -Name "StaffCode" -Value $($onPremSam.OnPremisesSamAccountName)
        $user | Add-Member -MemberType NoteProperty -Name "MeetingLink" -Value $($meeting.JoinWebUrl)
        $user | Add-Member -MemberType NoteProperty -Name "OnlineMeetingId" -Value $($meeting.Id)
    }
    $UserList | Export-Csv -Path $CurrentList -NoTypeInformation -Force -
    Export-RunningMeetingList -Users $UserList -CSVPath $RunningList
    #$UserList | Export-Csv -Path "./$($CurrentYear)-PTSI-Links.csv" -Append -NoTypeInformation
    Write-Host "INFO: Script finished running."
}

function Get-GraphToken {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [string]$Grant,
        [Parameter(Mandatory=$true, Position = 1)]
        [string]$Scope,
        [Parameter(Mandatory=$true, Position = 2)]
        [string]$Client_Id,
        [Parameter(Mandatory=$true, Position = 3)]
        [string]$Client_Secret,
        [Parameter(Mandatory=$true, Position = 4)]
        [string]$Tenant_Name
    )

    # Populate API Body
    $Body = @{
    Grant_Type = "client_credentials"
    Scope = "https://graph.microsoft.com/.default"
    client_Id = $Client_Id
    Client_Secret = $Client_Secret
    }

    try {
        $ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Tenant_Name/oauth2/v2.0/token" -Method POST -Body $Body
    } catch {
        Write-Host "ERROR: Failed to acquire token - exiting script"
        exit
    }
    $token = $ConnectGraph.access_token
    Write-Host "INFO: Acquired token for Graph API"
    return $token
}

function Get-TeachersList {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [PSObject]$GroupIDs
    )

    $membersArray=@()
    foreach ($group in $GroupIDs) {
        foreach ($member in $(Get-MgGroupMember -GroupId $group -All)) {
                $membersArray += [PSCustomObject]@{
				    Id = $member.id
				    DisplayName = $member.AdditionalProperties.displayName
				    EmailAddress = $member.AdditionalProperties.mail
			}
        }
    }
    Write-Host "INFO: Collected $($membersArray.Count) users from defined groups"
    return $membersArray
}

function Create-OnlineMeeting {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [PSObject]$User,
        [Parameter(Mandatory=$true, Position = 1)]
        [string]$MeetingSubject,
        [Parameter(Mandatory=$true, Position = 2)]
        [System.DateTime]$StartDateTime,
        [Parameter(Mandatory=$true, Position = 3)]
        [System.DateTime]$EndDateTime
    )

    $params = @{
        ExternalId = "PTSI-$($CurrentYear)-$($User.Id)"
        StartDateTime = $StartDateTime
        EndDateTime = $EndDateTime
        Subject = "$MeetingSubject - $($User.DisplayName)"
    }
    $timeout = 10
    # Invoke API Command to check if meeting exists, if not, create it. (Uses ExternalID as the Unique Identifier)
    $meeting = Invoke-MgCreateOrGetUserOnlineMeeting -UserId $User.Id -BodyParameter $params
    if ($meeting) {Write-Host "INFO: Meeting already existed for $($User.DisplayName), Grabbing info..."} else { Write-Host "INFO: Meeting does not exist for $($User.DisplayName), Creating meeting..." }
    while ($timeout -gt $i -and (-not $meeting)) {
        $i += 1
        # The invoke does not respond when creating a new meeting, run again until we see a confirmation - fail after 20 attempts (10s)
        $meeting = Invoke-MgCreateOrGetUserOnlineMeeting -UserId $User.Id -BodyParameter $params
        # Give the API time to breath
        sleep .5
    }

    # If $meeting is populated, an onlineMeeting exists - Lets set the rules of the meeting. 
    if ($meeting) {
        # Redeclare parameters for updating
        $params = @{
            Subject = "$MeetingSubject - $($User.DisplayName)"
            AllowMeetingChat = "disabled"
            LobbyBypassSettings = @{
                IsDialInBypassEnabled = $false
                Scope = "organizer"
            }
        }
        # Lets run an update to ensure we have the correct Team restrictions
        Update-MgUserOnlineMeeting -UserId $User.Id -OnlineMeetingId $meeting.Id -BodyParameter $params
        return $meeting
    } else {
        # If we couldn't create a meeting, return ERROR, this will appear in export. 
        return "ERROR: Failed to create an online meeting for $($User.DisplayName)"
    }
    
}

function Export-RunningMeetingList {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [PSObject]$Users,
        [Parameter(Mandatory=$true, Position = 1)]
        [string]$CSVPath
    )
    try {$Existing = Import-Csv -Path $CSVPath} catch {Write-Host "First run this year!"}
    $MissingEntries = Get-ArrayComparison -Array1 $Users -Array2 $Existing -Field "Id" -NoMatch
    $MissingEntries | Export-Csv -Path $CSVPath -Append -NoTypeInformation
}

function Get-ArrayComparison {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [PSObject]$Array1,
        [Parameter(Mandatory=$true, Position = 1)]
        [AllowNull()]
        [PSObject]$Array2,
        [Parameter(Mandatory=$true, Position = 2)]
        [string]$Field,
        [switch]$NoMatch
    )
    if ($Array1 -ne $null -and $Array2 -ne $null) {      
        if ($nomatch) {
            $results = $Array1 | Where-Object { $_.$($Field) -notin $Array2.$($Field) }
        } else {
            $results = $Array1 | Where-Object { $_.$($Field) -in $Array2.$($Field) }
        }
    } elseif ($Array1 -ne $null -and $Array2 -eq $null) {
        $results = $Array1
    } else { 
        throw "Array 1 is empty - cannot compare an empty array."
    }
    return $results
}

# Cleanup And Remove functions TBC. Needs further testing. 
# Issue with DELETE onlineMeeting Method
# https://github.com/microsoftgraph/microsoft-graph-docs/issues/17590
function Invoke-Cleanup {
        Write-Host -BackgroundColor Yellow -ForegroundColor Black "WARNING - You are about to enter a destructive action, this will delete all Online Meetings from the spreadsheet."
        $response = Read-Host "To remove all online meetings, please enter the App registration Application ID"
        if ($response -eq $ApplicationID){
            # Get Users from CSV
            $csvPath = Read-Host "Enter path to CSV"
            Write-Host -BackgroundColor Red "ATTENTION - ABOUT TO DELETE ALL ONLINE MEETINGS"
            if ((Read-Host "Continue? (y/N)") -imatch "Y") {
                foreach ($user in (Import-Csv -Path "$csvPath")){
                    Write-Host " Deleting $($user.DisplayName)"
                    Remove-OnlineMeetings -User $user
                }
            }
            exit
        } else {
            Write-Host "Please disable CleanUp variable before next run."
            pause
            exit
        }
}

function Remove-OnlineMeetings {
    [cmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position = 0)]
        [PSObject]$User
    )
    Remove-MgUserOnlineMeeting -UserId $User.Id -OnlineMeetingId $User.OnlineMeetingId -IfMatch
}


main
