<#
Script: Create-365-SharedMailbox.ps1
Author: Adeel Anwar
Date: September 1st, 2022
Duration: 10 minutes (est)

Description: 
	- Authenticate to Microsoft 365 using API triggering automation account (see other file) & create a shared mailbox using input parameters
		Mandatory: Shared Mailbox Display Name & Shared Mailbox Email
    	Optional: Permissions to the Shared Mailbox
	- Failure to properly name the shared mailbox will result in a fail
	- Failure to properlyly add permissions to the shared mailbox (User does not exist/Bad Formating) will result in those parameters being skipped
#>

####################
# INPUT PARAMETERS #
####################
# These are the parameters that need to be inputed. Some are mandatory and others are not (will accept $NULL)

param (
	[Parameter (Mandatory= $true)] # Mandatory
    $SharedMailbox = $null,
    [Parameter (Mandatory= $true)] # Mandatory
    $DisplayName = $null,
	[Parameter (Mandatory= $false)] # Optional
    $FullAccess = $null,
	[Parameter (Mandatory= $false)] # Optional
	$SendAs = $null,
    [Parameter (Mandatory= $false)] # Optional
	$SendBehalfOf = $null
)

############
# VALIDATE #
############
# This section validates the shared mailbox email and display name. To make sure the email is correctly formated and the display name/email isn't already in use
# Failure will result in stopping the script

# Validate format of Shared Mailbox to ensure it is infact an email
try {[ValidatePattern ('[a-zA-Z0-9]+@[a-zA-Z0-9]+\.[a-zA-Z0-9]')] $SharedMailboxValidate = $SharedMailbox}
catch {Write-Host "Invalid email format for $SharedMailbox Please try again." -ForegroundColor Yellow
    exit
}
# Validate existance of shared mailbox (Display Name & UPN) 
if ($null -ne (Get-mailbox -Identity $SharedMailbox -ErrorAction SilentlyContinue)){
    Write-Output "The email address: '$SharedMailbox' already exists"
    exit
}
if ($null -ne (Get-mailbox -Identity $DisplayName -ErrorAction SilentlyContinue)){
    Write-Output "The Display Name: '$SharedDisplayName' already exists"
    exit
}
Write-Output "Initial validation success. Proceding..."

#########################
# CREATE SHARED MAILBOX #
#########################
# This section creates the shared mailbox. It might still fail after passing the validation above. Errors will be caught via Try/Catch. 
# Failure will result in stopping the script

# Extract Alias from shared mailbox email.
$Alias = $SharedMailbox.Split("@")[0] # Split in to array and only store the section prior to the @ symbol in $Alias 

# Create shared mailbox
try {
    $CreationSuccess = New-Mailbox -Shared -Name $DisplayName -Alias $Alias -PrimarySmtpAddress $SharedMailbox -ErrorAction Stop
    if ($null -ne $CreationSuccess){
        Write-Output "CREATED: Shared mailbox '$SharedMailbox'."
    }
}
catch {
    Write-Output "The following error occured when creating the shared mailbox: $_"
    exit
}
Write-Output "Checking for permissions to be granted..."

#############################
# Assign Access Permissions #
#############################
# Assign permissions (all 3 types) if the parameters passed in values. Will check to see if users can be added. 
# Format must be email seperated by commas. It will split the data in an array from the commas & remove any empty spaces per array as it runs
# Failing to do this will result in the failed user being outputed in a report and will skip ahead to the next section
# Comments will be added only to the full access task as the other two follow the same procedure

# Assign Full Access if it's not null
if ($null -ne $FullAccess){
    $FAArrays = $FullAccess.Split(",") #Split at the commas to build an array of accounts
    ForEach ($FAArray in $FAArrays){ # For each email in the array
        $FAArray = $FAArray -replace " ","" #Remove any spaces
        if ($null -eq (Get-mailbox -Identity $FAArray -ErrorAction SilentlyContinue)) { # Fail to assign permission due to bad formating
            Write-Output "SKIP: Full Access permission request for ($FAArray) will not be added because the account does not exist or is an invalid format. Skipping"
        }
        elseif ($null -ne (Get-mailbox -Identity $FAArray -ErrorAction SilentlyContinue)) { # If user exists, add the permissions
            $CreationSuccess = $null
            try {
                $CreationSuccess = Add-MailboxPermission -Identity $SharedMailbox -User $FAArray -AccessRights FullAccess -InheritanceType All -ErrorAction Stop
                if ($null -ne $CreationSuccess){
                    Write-Output "ADDED: 'Full Access' permission request for ($FAArray) was granted!"
                }
            }
            catch {
                Write-Output "The following error occured when adding 'Full Access' permissions: $_"
                exit
            }
        }
    }
}
elseif ($null -eq $FullAccess){ # Skip if data was not provided
    Write-Output "No user was requested to be granted 'Full Access' permissions to the account, skipping"
}

# Assign Send As if it's not null
if ($null -ne $SendAs){
    $SAArrays = $SendAs.Split(",")
    ForEach ($SAArray in $SAArrays){ 
        $SAArray = $SAArray -replace " ","" 
        if ($null -eq (Get-mailbox -Identity $SAArray -ErrorAction SilentlyContinue)) { # Fail to assign permission due to bad formating
            Write-Output "SKIP: 'Send As' permission request for ($SAArray) will not be added because the account does not exist or is an invalid format. Skipping"
        }
        elseif ($null -ne (Get-mailbox -Identity $SAArray -ErrorAction SilentlyContinue)) { # If user exists, add the permissions
            $CreationSuccess = $null
            try {
                $CreationSuccess = Add-RecipientPermission -Identity $SharedMailbox -AccessRights SendAs -Trustee $SAArray -Confirm:$False -ErrorAction Stop
                if ($null -ne $CreationSuccess){
                    Write-Output "ADDED: 'Send As' permission request for ($SAArray) was granted!"
                }
            }
            catch {
                Write-Output "The following error occured when adding 'Send As' permissions: $_"
                exit
            }
        }
    }
}
elseif ($null -eq $SendAs){
    Write-Output "No user was requested to be granted 'Send As' permissions to the account, skipping"
}

# Assign "Send on Behalf" if it's not null
if ($null -ne $SendBehalfOf){
    $SBArrays = $SendBehalfOf.Split(",")
    ForEach ($SBArray in $SBArrays){
        $SBArray = $SBArray -replace " ",""
        if ($null -eq (Get-mailbox -Identity $SBArray -ErrorAction SilentlyContinue)) {
            Write-Output "SKIP: 'Send On Behalf Of' permission request for ($SBArray) will not be added because the account does not exist or is an invalid format. Skipping"
        }
        elseif ($null -ne (Get-mailbox -Identity $SBArray -ErrorAction SilentlyContinue)) {
            try {
                Set-Mailbox -Identity $SharedMailbox -GrantSendOnBehalfTo @{add=$SBArray} -ErrorAction Stop
                if($?){ #check if the last cmdlet executed successfully
                    Write-Output "ADDED: 'Send on Behalf of' permission request for ($SBArray) was granted!"
                }
            }
            catch {
                Write-Output "The following error occured when adding the SendonBehalfOf users: $_"
                exit
            }
        }
    }
}
elseif ($null -eq $SendBehalfOf){
    Write-Output "No user granted 'Send on Behalf Of' permissions, Skipping"
}

Write-Output "Job completed. Please review outputs to ensure you address any permission failures during the process that were skipped"
