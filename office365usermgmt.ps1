<#
.NOTES
    Version:    0.5
    Author:     Matvey Mirman
#>


<#
# ! This makes the program hang
if (!(Get-Module -ListAvaliable -Name -Connect-ExchangeOnline)) {
    Install-Module ExchangeOnlineManagement
}
#>

# Init Core Module:

Add-Type -AssemblyName System.Windows.Forms

# set core paramaters

$ProjectCSV = Import-Csv -Path "C:\\Path To File"


# Menu UI

#--------------------------------------------------#

function DrawFileSelectMenu {
    Write-Host "*************************"
    Write-Host "*       Welcome!        *"
    Write-Host "*************************"
    Write-Host
    Write-Host "With this script you will be able to create new emails, apply HTML signatures and more."
    Write-Host "Please provide .csv file with the following format (including headers!"
    Write-Host "| FirstName | LastName | Username | Password |"
    Write-Host "| Foo       | Bar      | foo.bar  | Qwerty12 |"
    Write-Host -NoNewline
}


function DrawProjectMenu($ProjectCSV) {
    Write-Host "*************************"
    Write-Host "*       Projects        *"
    Write-Host "*************************"
    Write-Host "Choose Project:"

    foreach ($item in $ProjectCSV) {
        Write-Host "[" $item.id "] " $item.ProjectName
    }
    Write-Host "[ 0 ] Quit."
    Write-Host
    Write-Host " Select an option and press Enter: " -NoNewline
}

function DrawActionMenu {
    Write-Host "*************************"
    Write-Host "*       Scripts         *"
    Write-Host "*************************"
    Write-Host
    Write-Host " [ 1 ] Create Accounts (Emails)"
    Write-Host " [ 2 ] Assing Signatures"
    Write-Host " [ 4 ] Quit"
    Write-Host
    Write-Host "Select an option and press Enter: " -NoNewline
}

function DrawPostActionMenu {
    Write-Host "Would you like to do a different action?"
    Write-Host
    Write-Host "1) Change Project"
    Write-Host "2) Change Script"
    Write-Host "3) Run Script Again"
    Write-Host "4) Quit"
    Write-Host
    Write-Host "Select an option and press Enter: " -NoNewline
}
function DrawOptionsMenu {
    Write-Host "*************************"
    Write-Host "*       Options         *"
    Write-Host "*************************"
    Write-Host
    Write-Host "1) Toggle file dialouge wait time: " $fileUiWait
    Write-Host "2) Toggle create email immidiatly after signature (can take long time!): " $createSigAfterEmail
    Write-Host "3) Credits"
    Write-Host "4) Quit"
    Write-Host
    Write-Host "Select an option and press Enter: " -NoNewline
}


function DelayWithSpinner ($DelayTime, $WaitText) {
    for ($i = 0; $i -le $DelayTime; $i++) {
        ## check only every five seconds so we don't spam the server
        Clear-Host
        Write-Host "`r $WaitText |" -NoNewline
        Start-Sleep -m 250
        Clear-Host
        Write-Host "`r $WaitText /" -NoNewline
        Start-Sleep -m 250
        Clear-Host
        Write-Host "`r $WaitText -" -NoNewline
        Start-Sleep -m 250
        Clear-Host
        Write-Host "`r $WaitText \" -NoNewline
        Start-Sleep -m 250
        Clear-Host
    }
}

#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~#

function Set-Credentials([securestring] $Credential, $Message) {
     Set-Variable $Credential -Scope global -Value (Get-Credential -Message $Message)
}
function GetFile {
    <#
    .SYNOPSIS
    Gets file using winforms file browser
    #>

    Clear-Host

    DrawFileSelectMenu

    Start-Sleep 2    
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Filter = 'CSV Files (*.csv)|*.csv' # Specified file types
        InitialDirectory = [System.Environment]::GetFolderPath('MyDocuments')
        RestoreDirectory = $true
    }
    $FileBrowser.ShowDialog()
    $script:FileSelected = Import-CSV -Path $FileBrowser.FileName
    Write-Host $FileBrowser.FileName
}

function ProjectMenu {
    Clear-Host
    DrawProjectMenu($ProjectCSV)

    # Read user input and store in var

    $selection = Read-Host

    # quits session if value is 0

    if($selection -eq 0) {
        Clear-Host
        return
    }

    # input handler - reads .csv file and sets var for project name.
    # * This sets the project for later use!

    $Project = $ProjectCSV | Where-Object { $_.id -eq $selection}
    if (!$Project) {
        Write-Host "Invalid Value!"
        Pause
        ProjectMenu
    }
    Write-Host "Project selected:" $Project.ProjectName

    Start-Sleep 2

    # Get Credentials

    $credMsg = "Enter credentials for " + $Project.ProjectName
    #Set-Variable "Credentials" -Scope global -Value (Get-Credential -Message $credMsg)

    Set-Credentials($Credential, $credMsg)

    ActionMenuFunc
}

function ActionMenuFunc($ActionSel) {
    Clear-Host
    if (!$ActionSel) {
        DrawActionMenu

        $ActionSel = Read-Host
        Write-Host "Initiating option " $ActionSel " ..."

        Start-Sleep 1
    }

    switch ($ActionSel) {
        1 {
            Clear-Host
            CreateEmails
        }
        2 {
            Clear-Host
            AssignSignatures
        }
        3 {
            Clear-Host
            # RemoveEmails
            Write-Host "This option will be ready Soon™"
            Pause
            ActionMenuFunc
        }
        4 {
            Clear-Host
            return
            Pause
        }
        Default {
            Write-Host "Please provide a valid value!"
            Pause
            Clear-Host
            ActionMenuFunc
        }
    }

    PostActioneMenuFunc
}

function PostActionMenuFunc($PostSel) {
    DrawPostActionMenu
    $PostSel = Read-Host

    switch ($PostSel) {
        1 {
            Clear-Host
            ProjectMenu
        }
        2 {
            Clear-Host
            ActionMenuFunc
        }
        3 {
            ActionMenuFunc($ActionSel)
        }
        4 { return }
        Default {
            Write-Host "Please provide a valid value!"
            Pause
            Clear-Host
            PostActionMenuFunc
        }
    }
}

# Logic
#---------------------------------#
# ! No real logic here

function CreateEmails() {
    # Connect to Office365

    Write-Host "Connecting to Exchange..."
    Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
    Write-Host "Connecting to Microsoft Outlook Service..."
    Connect-MsolService -Credential $Credential

    Clear-Host

    # GET users from CSV
    $script:FileSelected | ForEach-Object {

        # Some params
        $DisplayName = $_.FirstName + " " + $_.LastName
        $UserPrincipalName = $_.UserName + "@" + $Project.ProjectDomain

        # check if user exists
        if ([bool](Get-MsolUser -UserPrincipalName $UserPrincipalName -EA SilentlyContinue)) {
            Write-Host "User " $UserPrincipalName " Already exists!"
            return
        }

        Write-Host "Creating user: " $DisplayName " " $UserPrincipalName
        Write-Host

        # * CHANGE USAGE LOCATION TO WHATEVER YOU ARE GOING TO NEED
        New-MsolUser -DisplayName $DisplayName -FirstName $_.FirstName -LastName $_.LastName -UserPrincipalName $UserPrincipalName -UsageLocation "US" -LicenseAssigment $Project.ProjectLicense -Password $_.Password -ForceChangePassword $false -PasswordNeverExpires $true
    }
    # Done message

    Write-Host "Done!"
    Write-Host
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host
}

function AssignSignatures {
    <#
    .SYNOPSIS
    Set signature for org members
    
    .DESCRIPTION
    Set signature for organization members
    This must use an HTML template that is stored locally on the machine
    
    .NOTES
    Currently it can only change the name and not the department, but you can have multiple copies of signatures for each department


    * Signatures must be in accordance to template and contain those params written in plain text:
    * Name
    #>
        # Connect to Office365

        Write-Host "Connecting to Exchange..."
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
        
        # init department select menu
        Clear-Host

        DrawDeptSelect
        $selection = Read-Host

        # TODO: Get departments from JSON file
        switch ($selection) {
            1 { $department = "Sales Department" }
            2 { $department = "HR Department" }
        }

        Clear-Host
        $Source = "C:\Users\Public\Documents\Signatures" +$department +"\" + $Project.ProjectName + ".html"
        Write-Host "Using signature from: " $Source

        $script:FileSelected | ForEach-Object {

            # Variable definitions
            $File = (Get-Content -Path $Source -ReadCount 0) -join "`n"
            $SaveLocation  = "C:\Users\Public\Documents\Signatures\Export\Set"
            $DisplayName = $_.FirstName + " " + $_.LastName
            $OutputFile = $SaveLocation + $DisplayName + ".html"
            $UserPrincipalName = $_.UserName + "@" + $Project.ProjectDomain

            # Status output
            Write-Host "Creating signature for " $DisplayName " at " $Project.ProjectName

            $File -replace "Name", $DisplayName | Set-Content -Path $OutputFile

            Set-MailBoxConfiguration -Identity $UserPrincipalName -SignaturHtml (get-content -Path $OutputFile) -AutoAddSignature $true -AutoAddSignatureOnReply $true
            # ! Remove comment here for output of HTML and other maibox params
            # Get-MailboxMessageConfiguration -identity $UserPrincipalName
        }
    # Done message

    Write-Host "Done!"
    Write-Host
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host
}

# TODO Make the function send email to new owner regarding the new shared mailbox
# ! Currently doesn't appear for new owner
function RemoveEmails {
    <#
    .SYNOPSIS
    Remove email and free up license
    
    .DESCRIPTION
    Remove email by converting to shared mailbox and setting new owner (usually support)
    
    .NOTES
    # ! Currently this function doesn't work properly, do not use.
    #>

        # Connect to Office365

        Write-Host "Connecting to Exchange..."
        Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
        Write-Host "Connecting to Microsoft Outlook Service..."
        Connect-MsolService -Credential $Credential
    
        Clear-Host
    
        # GET users from CSV
        $script:FileSelected | ForEach-Object {
            $DisplayName = $_.FirstName + " " + $_.LastName
            $UserPrincipalName = $_.UserName + "@" + $Project.ProjectDomain
            $SupportEmail = "Support@" + $Project.ProjectDomain # * Change this if you want a different email to be the owner

            # Set autoreply message using the default MS exchange message

            $AutoReplyMessage = "Thank you for contacting " + $Project.ProjectName + ". We regret to inform you that " + $DisplayName + " is no longer employed here. Please direct any further correspondence to Support at " + $SupportEmail + ". This is an automated reply. For you convenience, this email has been automatically forwarder to " + $SupportEmail + "."

            # Check if user exists
            if ([bool]!(Get-MsolUser -UserPrincipalName $UserPrincipalName -EA SilentlyContinue)) {
                Write-Host "User " + $UserPrincipalName + " Doesn't exist!"
                return
            }
            else {
                Write-Host "Removing " $UserPrincipalName

                # Set autoreply message
                Set-MailboxAutoReplyConfiguration -Identity $UserPrincipalName -AutoReplyState Enabled -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage

                # Set mailbox owner
                "Converting Mailbox " + $UserPrincipalName + " to shared mailbox..."
                Set-Mailbox $UserPrincipalName -Type:Shared

                DelayWithSpinner -DelayTime 3 -WaitText "Converting mailbox " + $UserPrincipalName + " to shared mailbox"

                while (Get-Mailbox($UserPrincipalName) -eq $null) {
                    # loop until mailbox exists
                    DelayWithSpinner -DelayTime 3 -WaitText "Converting mailbox " + $UserPrincipalName + " to shared mailbox"
                }
                Clear-Host
                "`n Mailbox converted, proceeding..."
                "`n Giving " + $SupportEmail + " permissions to mailbox " + $UserPrincipalName

                Add-MailboxPermission -Identity $UserPrincipalName -User $SupportEmail -AccessRight FullAccess #give new owner access

                Get-EXOMailbox -Identity $UserPrincipalName | Format-Table Name, RecipientTypeDetails, Owner

                Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -RemoveLicenses $Project.ProjectLicense
            }
        }
            # Done message

    Write-Host "Done!"
    Write-Host
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host
}


#~~~~~$~~~~~$~~~~~$~~~~~#

# Run the program
GetFile
ProjectMenu