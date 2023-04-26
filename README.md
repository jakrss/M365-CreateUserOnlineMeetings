This script creates an online meeting for all users who are members of the target groups. Allowing the users a virtual 'meeting room' for the entire year, where they are able to send a link to external guests and have them permitted into the meeting as required.  

This initial concept for this was around online parent-teacher meetings, where external users can be added and removed from the meeting as required by the meeting owner.  

An Azure enterprise application is required with the below permissions configured.

## Permissions required: 
OnlineMeetings.ReadWrite.All    = Ability to create online meetings for users. Also requires CS Policy https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy (ParentTeacherMeetings-AppPolicy)  
GroupMember.Read.All            = Reads all users from Azure groups  
User.Read.All                   = Ability to read user attributes in Azure. This is only required as we are reading additional properties such as OnPremisesSamAccountName, otherwise you can use User.ReadBasic.All for only basic account info.  

# Application Policy
- https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy  
New-CsApplicationAccessPolicy -Identity ParentTeacherMeetings-AppPolicy -AppIds "application-guid" -Description "This Policy allows the enterprise app access to every user"  
Grant-CsApplicationAccessPolicy -PolicyName ParentTeacherMeetings-AppPolicy -Global  