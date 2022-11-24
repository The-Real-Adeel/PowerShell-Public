Powershell for MS Graph

**Common verbs used in MS Graph API to MS Graph in PowerShell**
    HTTP    --->  PS VERB --->  Example
    GET     --->  Get     --->  Get-MgUser
    POST    --->  New     --->  New-MgUserMessage
    PUT	    --->  New     --->  New-MgTeam
    PATCH   --->  Update  --->  Update-MgUserEvent
    DELETE	--->  Remove  --->  Remove-MgDriveItem

**Common nouns consistent of mg to not confuse with other cmds from other modules the rest follows the path of where the object lies**
Get-MgUserMailFolderMessage for example is pulling data out of Users > MailFolder > Messages object

**Import Module**
Import-Module Microsoft.Graph 

**Find cmds for a URI**
Find-MgGraphCommand -Uri '/users/{id}'

**Find scope/permissions for a cmd, ie. new-mguser**
Find-MgGraphCommand -command new-MgUser | Select -First 1 -ExpandProperty Permissions  | Format-List

**Connect to graph with scopes for read/write users and groups**
connect-mggraph -Scopes "user.readwrite.all","Group.readwrite.all","Organization.Read.All"  

**list users**
Get-MgUser

**Create New User**
 $PasswordProfile = @{Password = 'P@word!!11!!'}
 new-mguser -DisplayName "Tony Singh" -AccountEnabled -MailNickname 'tsingh' -UserPrincipalName 'tsingh@corpocorp.onmicrosoft.com' -PasswordProfile $PasswordProfile 

**Store user as an object**
 $user = Get-MgUser -Filter "displayName eq 'Tony Singh"

**Use the object stored in $user to find ID when using other cmds that require it**
 get-mguserjoinedteam -userId &user.Id

Get-MgUser | Select-Object DisplayName, UserPrincipalName, LicenseDetails, MemberOf | Format-List