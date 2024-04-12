#-------------------------
# Send Email using MsGraph
#-------------------------
# Function that can be called or added to a script to ensure msgraph 
# Simplifies the mgraph email command to allow sending emails to a comma seperate list from a single command in length. 
# Handles the authN process as well (if you are storing this in automation account)
# Requires a Service Principal(AppRegistration)
# NOTE: REMOVED SP INFO for public. Ensure its added

function Send-EmailFunction {
    <#
    .SYNOPSIS
        Send emails via MSGraph as Service Principal
    .DESCRIPTION
        Utilize inline function in Azure Automation to send emails in your scripts without having to include many lines of code below each time. This reduces alot of the verbose aspects by simply invoking with a few parameters: Subject, To, CC, BCC, Body, HTML & Importance.
    .PARAMETER Subject
        MANDATORY. Enter the subject of the email
    .PARAMETER Body
        MANDATORY. Enter the body of the email you wish to sent.
    .PARAMETER To
        MANDATORY. Enter the recipients in quatations. If you wish to include multiple seperate them by comma in a single string
    .PARAMETER CC
        OPTIONAL. Enter the CC recipients in quatations. If you wish to include multiple seperate them by comma in a single string
    .PARAMETER BCC
        OPTIONAL. Enter the BCC recipients in quatations. If you wish to include multiple seperate them by comma in a single string
    .PARAMETER Attachment
        OPTIONAL. Enter the path for the attachment (ie, "c:\test\file1"). Must be able to be navigated to from the runbook. If you wish to include multiple seperate them by comma in a single string
    .PARAMETER HTML
        OPTIONAL. Default set to $false. Set as $true if you wish the email body to utilize HTML.
    .PARAMETER Importance
        OPTIONAL. Default set to $false. Set as $true if you wish the email to be flagged as Important "!"
    .EXAMPLE
        Send-EmailFunction -Subject "Where are you?" -To "Bob@contoso.com" -Body "You are late!"
        # Send a simple email with the least amount of parameters needed
    .EXAMPLE
        Send-EmailFunction -Subject "Todays agenda"  -To "johnsmith@contoso.com, coyotemike@contoso.com" -CC "bob@contoso.com" -BCC "accounting@contoso.com" -Body $Body -HTML $True -Importance $True
        # Send an email with a subject, TO, CC, BCC, Body, HTML and Importance
    .NOTES
        Author: Adeel
        Version: 
            1.0 - Created Function
    #>
        [cmdletBinding()]
        param (
            [Parameter(Mandatory=$True)]
            [String]$To,
            [Parameter(Mandatory=$True)]
            [String]$Subject,
            [Parameter(Mandatory=$True)]
            [String]$Body,
            [Parameter(Mandatory=$False)]
            [String]$CC,
            [Parameter(Mandatory=$False)]
            [String]$BCC,
            [Parameter(Mandatory=$False)]
            [String]$Attachment,
            [Parameter(Mandatory=$False)]
            [Bool]$HTML = $False,
            [Parameter(Mandatory=$False)]
            [Bool]$Importance = $False
        )
    # Set Variables that will be used to compose the email details.
        $ToObject = @()
        $CCObject = @()
        $BCCObject = @()
        $AttachmentObject = @()
        $BodyObject = @()
        $MessageHashtable = @{} # Create hashtable to place all objects into
        $ToItem = $null
  
    # Subject Composition. Create hashtable and add subject in to it
        $MessageHashtable +=  @{subject = $Subject}

    # To Field composition. Add the fields to the same hashtable.
        $ToArray = ($To.split(',')).Replace(' ','')
        foreach ($ToItem in $ToArray) {
            $ToObject += @{ 
                emailAddress = @{
                    address = $ToItem
                }
            }
        }
        $MessageHashtable += @{toRecipients = $ToObject}
    # CC Field composition. Add the fields to the same hashtable.
        if ($CC -ne ""){
            $CCArray = ($CC.split(',')).Replace(' ','')
            foreach ($CCItem in $CCArray) {
                $CCObject += @{ 
                    emailAddress = @{
                        address = $CCItem
                    }
                }
            }
            $MessageHashtable += @{ccRecipients = $CCObject}
        }
    # BCC Field composition. Add the fields to the same hashtable.
        if ($BCC -ne ""){
            $BCCArray = ($BCC.split(',')).Replace(' ','')
            foreach ($BCCItem in $BCCArray) {
                $BCCObject += @{ 
                    emailAddress = @{
                        address = $BCCItem
                    }
                }
            }
            $MessageHashtable += @{bccRecipients = $BCCObject}
        }
    # Attachment Composition
        if ($Attachment -ne ""){
            $AttachmentArray = ($Attachment.split(',')).Replace(' ','')
            foreach ($AttachmentItem in $AttachmentArray) {
                $FileContent = [System.IO.File]::ReadAllBytes($AttachmentItem) # Read the file and convert it to a byte array
                $FileBase64 = [System.Convert]::ToBase64String($FileContent) # Encode the byte array to base64
                $FileName = [System.IO.Path]::GetFileName($AttachmentItem) # Store File Name
                $AttachmentObject += @{ 
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    Name = $FileName
                    ContentType = "text/plain"
                    ContentBytes = $FileBase64
                }
                $FileBase65 = $null
                $FileContent = $null
            }
            $MessageHashtable += @{Attachments = $AttachmentObject}
        }
    # Content Type Composition
        if ($false -eq $HTML){
            $ContentType= "text"
        }else{$ContentType = "html"}
    
    # Body Composition. Take in both content type (html or text) and the content itself and add it to the hashtable
        $BodyObject = @{
            Content = $Body
            ContentType = $ContentType
        }
        $MessageHashtable +=  @{Body = $BodyObject}
    
    # Importance  Composition. Create hashtable and add subject in to it
        if ($true -eq $Importance){
            $MessageHashtable +=  @{Importance = "High"}
        }
    # Authenticate as a service principal that has the capabilties to send email to users in the organization only if a session already isn't active!
        $FromMailbox = "<EmailOfSender>"
        $AppName = '<SPNameInKV>'
        $KeyVaultName = '<kvName'
        $MailFunctionAuthTenantId = '<TenantID>'
        $MailFunctionAuthhClientId = '<ClientID>'
        $MailFunctionAuthClientSecret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name $AppName -AsPlainText
        $MailFunctionAuthbody =  @{
            Grant_Type    = "client_credentials"
            Scope         = "https://graph.microsoft.com/.default"
            Client_Id     = $MailFunctionAuthhClientId
            Client_Secret = $MailFunctionAuthClientSecret
        }
        $MailFunctionAuthconnection = Invoke-RestMethod `
            -Uri https://login.microsoftonline.com/$MailFunctionAuthTenantId/oauth2/v2.0/token `
            -Method POST `
            -Body $MailFunctionAuthbody
        $MailFunctionAuthtoken = $MailFunctionAuthconnection.access_token
        $MailFunctionAuthtoken = $MailFunctionAuthtoken | ConvertTo-SecureString -AsPlainText -Force ########## THIS LINE IS NEEDED for 2.0
        if (Connect-MgGraph -AccessToken $MailFunctionAuthtoken){
            write-output "Successfully connected as Email Function Service Principal"
        }
        else {
            Write-Output "Failed to authenticate as Email Function Service Principal"
            throw "Failed"
        }
    # Finalize data and send email
        $MailBodyParameters   = @{'message'         = $MessageHashtable}
        $MailBodyParameters  += @{'saveToSentItems' = $True}

        try {
            Send-MgUserMail -UserId $FromMailbox -BodyParameter $MailBodyParameters -ErrorAction Stop
            write-output "Sent Email as $($FromMailbox) to $To"
        }
        catch {
            Write-Output "Failed to send email from the function as $($FromMailbox) to $To"
            throw "Failed"
        }
}

function New-HTMLTable {
    <#
.SYNOPSIS
    Generate HTMLTables without having to enter all the HTML required. You can still inject HTML in to what is passed along with special switches for links/colors
.DESCRIPTION
    Utilize inline function in Azure Automation to generate HTML tables into your scripts without having to include many lines of code below each time. This reduces alot of the verbose aspects of creating such tables that are required and allow you to focus on the job
    Allows for a few styles/colors as well as using special 'switches' to set data as links or in specific colors
    Columns only accept arrays. Total column number is set by the amount of data added in the array
    Rows only accept objects as we need an array of data per row. # of cells that will be filled in the rows must match the number of columns provided in the columns parameter. $null/"" can be used when there is no data to be passed for the specific cell.
    Switches(inside row cell data): Inside the row variable if you enter data for a cell to be a specific color or be a link (with a title text inside of the URL) set them with custom switches using !!<switch>!!: Here are the options
        "!!GREEN!!$Data" # set color to green
        "!!RED!!$Data"  # set color to red
        "!!BLUE!!$Data" # set color to blue
        "!!AMBER!!$Data" # set color to Amber
        "!!LINK!!$URL!!TITLE!!Click Here."
    For any further modification, you can add HTML to the row/column data that you are adding. ie. if you want to bold the green text using the switch above, do the following: "!!GREEN!!<b>$Data</b>"
    Opening & Closing HTML can be added to include data before and after the table. These do use html styling so if you want to bold, highlight, line breaks, etc. 
    Styles & Themes allow you to set different styles for the tables.
    The output would be the full concatation of the HTML that can be set as the email body

.PARAMETER Columns
    MANDATORY. Enter an array of data that will be used to build the column headers. How many you add here will dictate the total width of the table cells. Ordered left to right
.PARAMETER Rows
    MANDATORY. Enter the objects for the row that you will place inside the columns listed above. Total number of data inside the object must match the column count. Name parameter is required for objects but not used. We just enter it numerically. The function only targets the values. Ordered first in.
.PARAMETER OpeningHTML
    OPTIONAL. Data you would like to have in HTML before the table. Great place to place your header text, introduction data, etc. Must use HTML format (ie for new lines you use line breaks)
.PARAMETER ClosingHTML
    OPTIONAL. Data you would like to have in HTML after the table. Great place to place your closing text, signatures, etc. Must use HTML format (ie for new lines you use line breaks)
.PARAMETER Style
    OPTIONAL. Default is Fancy but you can change it to other ones. Right now your options are "Fancy" & "Simple"
.PARAMETER Theme
    OPTIONAL. Theme refers to colours but only affects Fancy theme as of now. Default is green. Right now your options are "Green", "Red", "Blue", "Black"
.EXAMPLE
    $RG = Get-AzResourceGroup | Select-Object ResourceGroupName, Location, ProvisioningState, ResourceId
    $Columns = "Subscription", "ResourceGroupName", "Location", "ProvisioningState", "ResourceId"
    $Rows = New-Object System.Collections.Specialized.OrderedDictionary # This sets the $variable to a hash
    $RowCount = 1 # Use counts: this will be the Name for the name/value pair in the hashtable
    $OpeningHTML = "This is our text to see how all of this flows.<br> New Line is entered."
    $ClosingHTML = 'This is our <b>outro text</b><br>
    <p style="border-width:3px; border-style:solid; border-color:#FF0000; padding: 1em;"> <b>What do I need to do?</b><br><br>
    Contact: WEB AND IDENTITY TEAM.'
    foreach ($item in $rg) {
        $RowCount += 1
        $Row = "<b>$Subscription</b>", "$($item.ResourceGroupName)", "$($item.Location)", "!!GREEN!!<b>$($item.ProvisioningState)</b>", "!!LINK!!<domain>/resource$($item.ResourceId)!!TITLE!!$($item.ResourceId)"
        $Rows += @{$RowCount = $Row}
    }
    New-HTMLTable -Rows $Rows -Columns $Columns -OpeningHTML $OpeningHTML -ClosingHTML $ClosingHTML -Theme "Red" -Style $Style 
    # Set the columns and set the rows with the data. Notice in the rows html around $subscription (optional), Colors for $item.ProvisioningState and link for the last row with a title
.EXAMPLE
    New-HTMLTable -Rows $Rows -Columns $Columns
    # The bare minimum it will accept
.EXAMPLE
    New-HTMLTable -Rows $Rows -Columns $Columns -Theme "Black"
    # Since the default is to accept style as fancy you can just set the theme to the color you want
.NOTES
    Author: Adeel
    Version: 
        1.0 - Created Function
#>
[cmdletBinding()]
param (
    [Parameter(Mandatory=$True)]
    [Array]$Columns,
    [Parameter(Mandatory=$True)]
    [Object]$Rows,
    [Parameter(Mandatory=$False)]
    [String]$OpeningHTML,
    [Parameter(Mandatory=$False)]
    [String]$ClosingHTML,
    [Parameter(Mandatory=$False)]
    [String]$Theme, # Accepts: 'Blue', 'Red', 'Green'
    [Parameter(Mandatory=$False)]
    [String]$Style # Accepts: 'Simple' or 'Fancy'
)
$blank = '"_blank"' #Used for generating cells with links/title
$Enumerate = 0
if ($false -eq $Style){
    $Style = "Fancy"
}
if ($false -eq $Theme){
    $Theme = "Green"
}
if ($false -eq $OpeningHTML){
    $OpeningHTML = ""
}
if ($false -eq $ClosingHTML){
    $ClosingHTML = ""
}
#Assign Head data to HTML. Includes all the style options
$HTMLHEAD = '
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">     
<html xmlns="http://www.w3.org/1999/xhtml">
<head>'
if($Style -eq 'Fancy'){
    if($Theme -eq 'Green'){$ThemeColour = '#009879'}
    if($Theme -eq 'Blue'){$ThemeColour = '#4682b4'}
    if($Theme -eq 'Red'){$ThemeColour = '#a52a2a'}
    if($Theme -eq 'Black'){$ThemeColour = '#000000'}
    $HTMLHEAD += '
        <style type="text/css">
            * {
                font-family: sans-serif;
            }
            .content-table {
                border-collapse: collapse;
                margin: 5px 0;
                font-size: 0.9em;
                min-width: 400px;
                border-radius: 5px 5px 0 0;
            }'
    $HTMLHEAD += "
            .content-table thead tr {
                background-color: $ThemeColour;
                color: #ffffff;
                text-align: left;
                font-weight: bold;
            }
            .content-table th,
            .content-table td {
                padding: 1px 15px;
                border-width: 1px;
                border-color: $ThemeColour;
                border-style: none none solid none;
            }
            .content-table td.red {
                color: #BF0000;
            }
            .content-table td.green {
                color: #009900;
            }
            .content-table td.amber {
                color: #E0A800;
            }
            .content-table td.blue {
                color: #0066b2;
            }
            .content-table tr.grey {
                background-color: #f3f3f3;
            }
            .content-table tr.transparent {
                background-color: transparent;
            }
        </style>"
}
if($Style -eq 'Simple'){
    $HTMLHEAD += '
        <style type="text/css">
            * {
                font-family: sans-serif;
            }
            .content-table {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
            .content-table thead tr {
            }
            .content-table th {
                border-width: 1px; padding: 5px; border-style: solid; border-color: black;
            }
            .content-table td {
                border-width: 1px; padding: 5px; border-style: solid; border-color: black;
            }
            .content-table td.red {
                color: #BF0000;
            }
            .content-table td.green {
                color: #009900;
            }
            .content-table td.amber {
                color: #E0A800;
            }
            .content-table td.blue {
                color: #0066b2;
            }
        </style>'
}
$HTMLHEAD += '
    </head>
    <body>'
# Insert text that users want prior to table view
$HTMLHEAD += "
    <p>$OpeningHTML
    </p>"
$HTMLHEAD += '
    <table class="content-table">
        <thead>'

# HTML COLUMN CREATION
$HTMLBODYCOLUMN = '
            <tr>'
# Fill Column Body
foreach ($cell in $columns){
    $HTMLBODYCOLUMN += "
                <th>$cell</th>"
}         
$HTMLBODYCOLUMN += '   
            </tr>
        </thead>'

# HTML ROW CREATION
$HTMLBODYROW = '
        <tbody>'
# Fill Row Body
foreach ($row in $rows.values){
$Enumerate += 1
# Set Grey or Transparent row
if($Style = 'Fancy'){
    if($Enumerate % 2 -eq 0 ){
            $HTMLBODYROW += '
                    <tr class = "grey">'
        }
    if($Enumerate % 2 -eq 1 ){
            $HTMLBODYROW += '
                    <tr class = "transparent">'
        }
    }
    # Add in data for the rows
    foreach ($cell in $row){
            if ($cell -like '!!RED!!*') { # Set the cell to Red if flagged as Red
                $redtransform = $cell -replace '!!RED!!',""
                $HTMLBODYROW += "
                <td class = 'red'>$($redtransform)</td>"
            }
            elseif ($cell -like '!!AMBER!!*') { # Set the cell to Amber if flagged as Amber
                $ambertransform = $cell -replace '!!AMBER!!',""
                $HTMLBODYROW += "
                <td class = 'amber'>$($ambertransform)</td>"
            }
            elseif ($cell -like '!!GREEN!!*') { # Set the cell to Green if flagged as Green
                $greentransform = $cell -replace '!!GREEN!!',""
                $HTMLBODYROW += "
                <td class = 'green'>$($greentransform)</td>"
            }
            elseif ($cell -like '!!BLUE!!*') { # Set the cell to Blue if flagged as Blue
                $greentransform = $cell -replace '!!BLUE!!',""
                $HTMLBODYROW += "
                <td class = 'blue'>$($greentransform)</td>"
            }
            elseif ($cell -like '!!LINK!!*' -and $cell -like '*!!TITLE!!*') { # Set the cell to a link with title if flagged as link
                $titletransform = $cell -replace "^.*!!TITLE!!",""
                $linktransform = (($cell -replace "!!LINK!!","") -replace "!!TITLE!!.*$","")
                $HTMLBODYROW += "
                <td>&nbsp<a href=$linktransform target=$blank>$($titletransform)</a>&nbsp</td>"
            }
            else{
                $HTMLBODYROW += "
                <td>$($cell)</td>"
            }
    }
    # Close Table Rows
    $HTMLBODYROW += '
    </tr>'
} 
# Close Table Body
$HTMLBODYROW += ' 
        </tbody>'

# Close everything
$HTMLCLOSE = '
    </table>'
$HTMLCLOSE += "
    <p>$ClosingHTML
    </p>"
$HTMLCLOSE += '
    </body>
    </html>
'

$ReturnHTML = $HTMLHEAD + $HTMLBODYCOLUMN + $HTMLBODYROW +$HTMLCLOSE
return $ReturnHTML
}