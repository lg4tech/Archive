<#
.SYNOPSIS
    This script will create users in Active Directory by importing data
    from a form with pre-filled data.

.DESCRIPTION
    When the proper fields are filled out in the form, this script 
    will:
        -Import users into AD with their details
        -Generate and output a random password
        -Create their home folder and enable folder sharing
        -Place them in the proper OU based on their city
        -Assign group membership (optional).

    Additionally, if the importer specifies TRUE for Mailbox and Lync, this
    script will build the user in Exchange and enable them in Lync. 
        *** Lync is not required, if only a mailbox is desired then select TRUE for
        'Mailbox' and FALSE for 'Skype'.

    There are two choices for password generation:
        $SecurePass is used for complex passwords, including special characters.
        $TempPass is used for easy communication to the end user.

.NOTES
    File name         - Create users in Active Directory with GUI.ps1
    Version           - 1.31
    Last Updated      - 04/10/2021
    Author            - Lou Garramone
    E-mail            - lou@lg4tech.com

    # Required setup prior to running this script on Windows 10 1809 or newer.,
    Add-WindowsCapability –online –Name “Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0”
#>

#Generated Form Function
function GenerateForm {

#region Import the Assemblies
[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
#endregion

#region Generated Form Objects
$mainForm = New-Object System.Windows.Forms.Form
$checkBoxMail = New-Object System.Windows.Forms.CheckBox
$checkBoxLync = New-Object System.Windows.Forms.CheckBox
$btnCreate = New-Object System.Windows.Forms.Button
$lblOU = New-Object System.Windows.Forms.Label
$comboOU = New-Object System.Windows.Forms.ComboBox
$textManager = New-Object System.Windows.Forms.TextBox
$lblManager = New-Object System.Windows.Forms.Label
$lblDisplayName = New-Object System.Windows.Forms.Label
$textDisplayName = New-Object System.Windows.Forms.TextBox
$lblUserNameExample = New-Object System.Windows.Forms.Label
$lblUserName = New-Object System.Windows.Forms.Label
$textUserName = New-Object System.Windows.Forms.TextBox
$textCopySecurity = New-Object System.Windows.Forms.TextBox
$lblCopySecurity = New-Object System.Windows.Forms.Label
$lblCompany = New-Object System.Windows.Forms.Label
$textCompany = New-Object System.Windows.Forms.TextBox
$lblDepartment = New-Object System.Windows.Forms.Label
$textDepartment = New-Object System.Windows.Forms.TextBox
$textMobilePhone = New-Object System.Windows.Forms.TextBox
$lblMobilePhone = New-Object System.Windows.Forms.Label
$lblJobTitle = New-Object System.Windows.Forms.Label
$textJobTitle = New-Object System.Windows.Forms.TextBox
$lblPhone = New-Object System.Windows.Forms.Label
$textPhone = New-Object System.Windows.Forms.TextBox
$comboLocation = New-Object System.Windows.Forms.ComboBox
$lblLocation = New-Object System.Windows.Forms.Label
$textLastName = New-Object System.Windows.Forms.TextBox
$lblLastName = New-Object System.Windows.Forms.Label
$lblFirstName = New-Object System.Windows.Forms.Label
$textFirstName = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion Generated Form Objects

#----------------------------------------------
#Generated Event Script Blocks
#----------------------------------------------
$handler_mainForm_Load= 
{
# Import modules necessary for proper script execution.
Try {
Import-Module ActiveDirectory
} Catch{write-host "Import Module(s) failure, terminating..."; Break}
}

$btnCreate_OnClick= 
{

function Create-Exchange {
# Start new session to Exchange.
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange `
-ConnectionUri http://$ExchangeServer/PowerShell/ -Authentication Kerberos

# Import the newly created PS Session to the Exchange server.
Import-PSSession $ExchangeSession -AllowClobber -DisableNameChecking
            
# Create user mailbox.
New-Mailbox `
-DomainController $DC `
-Name $DisplayName `
-UserPrincipalName $UserPrincipalName `
-Alias $SamAccountName `
-OrganizationalUnit $OU `
-SamAccountName $SamAccountName `
-FirstName $FirstName `
-LastName $LastName `
-Password (ConvertTo-SecureString -AsPlainText $TempPass -Force)`
-ResetPasswordOnNextLogon $TRUE | Out-Null
<#-Database $MailboxDatabase  *Removed for testing with new Exchange DB Auto Sorting*#>

}

function Set-UserInfo {
<#
.SYNOPSIS
    This function will set user information in AD.

.DESCRIPTION
    The information is taken from the .csv and set in
    the fields below in AD.
#>
Set-ADUser `
-Server $DC `
-Identity $SamAccountName `
-DisplayName $DisplayName `
-SamAccountName $SamAccountName `
-UserPrincipalName $UserPrincipalName `
-Company $Company `
-Department $Department `
-Title $Title `
-Manager $Manager `
-StreetAddress $StreetAddress `
-City $City `
-Office $City `
-State $State `
-PostalCode $PostalCode `
-Country $Country `
-OfficePhone $Telephone `
-MobilePhone $MobilePhone `
-Enabled $True `
-ScriptPath "logon.bat" 
}

# Defining variables.
$Date = Get-Date -Format g
$Domain = "domain.com"
$DC = "dc.domain.com"
$ExchangeServer = "exchange.domain.com"
$SkypeServer = "skype.domain.com"
$ScriptPath = "logon.bat"
$BrowseDirectory = "$Env:USERPROFILE\Desktop"

############################## Script Start ##############################

# The following block(s) will be executed for each user (or row) imported.
ForEach-Object {

    # Define variables for user details (from the form).
    $DisplayName = $textDisplayName.Text
    $FirstName = $textFirstName.Text
    $LastName = $textLastName.Text
    $SamAccountName = $textUserName.Text
    $Company = $textCompany.Text
    $Manager = $textManager.Text
    $Department = $textDepartment.Text
    $Title = $textJobTitle.Text
    $City = $comboLocation.SelectedItem
    $Country = "US"
    $MobilePhone = $textMobilePhone.Text
    $OU = $comboOU.SelectedItem + ",DC=domain,DC=domain,DC=com" #---------------------------------------------------------------------------------------#
    $CopySecurity = $textCopySecurity.Text
    $UserPrincipalName = $SamAccountName + "@" + $Domain
    $HomeDirectory = "<NETWORK PATH>\$SamAccountName" #---------------------------------------------------------------------------------------#

    # Prompt the user for their password, username is gathered from current logon.
    $UserCredential = Get-Credential -Credential $ENV:USERNAME

    # if user exists, display error and terminate.
    if(dsquery user -samid $SamAccountName) {
    [System.Windows.Forms.MessageBox]::Show("User " + $SamAccountName + " already exists. Click OK to exit.","Duplicate User",0) | Out-Null
    $mainForm.Close();
    }

    # Assign a state and address to the user based on their City selected.
    switch($City){
        "CityName"    {$StreetAddress = "123 Main Street" #---------------------------------------------------------------------------------------#
                          $State = "NY"
                          $PostalCode = "12345"
                          $Telephone = "(123) 456-7890"}
    }

    # Generate a complex random password, 8 characters long, non/alphanumeric.
    # TO USE: Replace TempPass with SecurePass during user creation below.
    #[Reflection.Assembly]::LoadWithPartialName(“System.Web”) | Out-Null
    #$SecurePass = [System.Web.Security.Membership]::GeneratePassword(8,1)
    # Array of characters/integers for random password generation.
    $Chars = [char[]]"abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    # Generate a random password, 8 characters long, alphanumeric.
    $TempPass = ($Chars | Get-Random -Count 7) + (Get-Random -Minimum 1 -Maximum 100) -join ""

    # Mailbox and Lync are TRUE, create mailbox, set AD details and enable Skype.
    if($checkboxMail.Checked -eq $TRUE -and $checkboxLync.Checked -eq $TRUE) {
        Try {
        # Start an Exchange PS Session and call Set-UserInfo.
        Create-Exchange -UserCredential $UserCredential
        Start-Sleep -Seconds 3
        Set-UserInfo
        # Start new PS session to the Lync server.  
        Import-Module Lync                     
        $SkypeSession = New-PSSession -ConnectionUri https://$SkypeServer/OcsPowerShell -Credential $UserCredential
            
        # Import the new PS session to the Skype server.
        Import-PSSession $SkypeSession -AllowClobber -DisableNameChecking | Out-Null
        # Enable the user in Lync.
        Get-Mailbox $SamAccountName -DomainController $DC | Enable-CsUser -DomainController $DC -RegistrarPool $SkypeServer `
        -SipAddressType EmailAddress

        # An exception will be output to the screen and to a txt file.
        } Catch {Write-Host "$DisplayName - Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        "$DisplayName - Exception Type: $($_.Exception.GetType().FullName) $Date" | Out-File -FilePath "$BrowseDirectory\New User Errors.txt" -Append
        $mainForm.Close()}

     }
    # Create mailbox and AD account, excluding Skype.
    elseif($checkboxMail.Checked -eq $TRUE) {
        Try {
        # Start an Exchange PS Session and call Set-UserInfo.
        Create-Exchange -UserCredential $UserCredential
        Set-UserInfo

        # An exception will be output to the screen and to a txt file.
        } Catch {Write-Host "$DisplayName - Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        "$DisplayName - Exception Type: $($_.Exception.GetType().FullName) $Date" | Out-File -FilePath "$BrowseDirectory\New User Errors.txt" -Append
        $mainForm.Close()}

    }
    # Create an AD 'shell' account if Exchange and Skype boxes are not selected.
    else
    {
        Try {
        New-ADUser `
        -Name $DisplayName `
        -SamAccountName $SamAccountName `
        -UserPrincipalName $UserPrincipalName `
        -Company $Company `
        -Department $Department `
        -Manager $Manager `
        -Title $Title `
        -StreetAddress $StreetAddress `
        -City $City `
        -Office $City `
        -State $State `
        -Country $Country `
        -PostalCode $PostalCode `
        -OfficePhone $Telephone `
        -MobilePhone $MobilePhone `
        -DisplayName $DisplayName `
        -GivenName $FirstName `
        -Surname $LastName `
        -AccountPassword (ConvertTo-SecureString -AsPlainText $TempPass -Force) `
        -Enabled $TRUE `
        -ChangePasswordAtLogon $TRUE `
        -ScriptPath $ScriptPath `
        -Path $OU | Start-Sleep 2

        # An exception will be output to the screen and to a txt file. 
        } Catch {Write-Host "$DisplayName - Exception Message: $($_.Exception.Message)" -ForegroundColor Red
        "$DisplayName - Exception Type: $($_.Exception.GetType().FullName) $Date" | Out-File -FilePath "$BrowseDirectory\New User Errors.txt" -Append
        $mainForm.Close()}
   }

# User creation successful, create home folder.
if($SamAccountName) {
    # Create H drive, assign the security values to the home folder ACL.
    New-Item $HomeDirectory -Type Directory | Out-Null
    $Acl = Get-Acl $HomeDirectory
    $AccessRights = ($SamAccountName, "Modify", "ContainerInherit, ObjectInherit", "None", "Allow")
    $AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $AccessRights
    $Acl.AddAccessRule($AccessRule)     
    $Acl | Set-Acl $HomeDirectory
    
    if($CopySecurity) {    
    # Copies user permissions from the name listed in CopyPermissionsFrom field.
    foreach ($Group in (Get-ADUser $CopySecurity -Properties MemberOf).MemberOf) {   
        Add-ADGroupMember $Group -Members $SamAccountName 
     }
    }
    # Outputs the newly created associate name and temporary password to the screen and txt file.
    $DisplayName + "`t" + $TempPass + "`t`t" + $Date | Out-File -FilePath "$BrowseDirectory\New User PW.txt" -Append | Invoke-Item
  }
  # Sleep for 2 seconds after creating the user before checking for success.
  Start-sleep 2

  # Check to see if the user was created successfully, if not return creation failed error.
  if(dsquery user -samid $SamAccountName) {
    [System.Windows.Forms.MessageBox]::Show("User " + $SamAccountName + " created.","User Creation Successful",0) | Out-Null
    } else {[System.Windows.Forms.MessageBox]::Show("User " + $SamAccountName + " creation failed.","User Creation Failure",0) | Out-Null}
 }
$mainForm.Close()
}

$OnLoadForm_StateCorrection=
{#Correct the initial state of the form to prevent the .Net maximized form issue
	$mainForm.WindowState = $InitialFormWindowState
}

#----------------------------------------------
#region Generated Form Code
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 329
$System_Drawing_Size.Width = 323
$mainForm.ClientSize = $System_Drawing_Size
$mainForm.DataBindings.DefaultDataSourceUpdateMode = 0
$mainForm.Name = "mainForm"
$mainForm.SizeGripStyle = 2
$mainForm.Text = "New Associate Creation"
$mainForm.add_Load($handler_mainForm_Load)


$checkBoxMail.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 237
$System_Drawing_Point.Y = 11
$checkBoxMail.Location = $System_Drawing_Point
$checkBoxMail.Name = "checkBoxMail"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 74
$checkBoxMail.Size = $System_Drawing_Size
$checkBoxMail.TabIndex = 16
$checkBoxMail.Text = "Mailbox"
$checkBoxMail.UseVisualStyleBackColor = $True

$mainForm.Controls.Add($checkBoxMail)


$checkBoxLync.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 237
$System_Drawing_Point.Y = 37
$checkBoxLync.Location = $System_Drawing_Point
$checkBoxLync.Name = "checkBoxLync"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 74
$checkBoxLync.Size = $System_Drawing_Size
$checkBoxLync.TabIndex = 17
$checkBoxLync.Text = "Skype"
$checkBoxLync.UseVisualStyleBackColor = $True

$mainForm.Controls.Add($checkBoxLync)


$btnCreate.DataBindings.DefaultDataSourceUpdateMode = 0
$btnCreate.Font = New-Object System.Drawing.Font("Microsoft Sans Serif",8.25,1,3,1)

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 237
$System_Drawing_Point.Y = 63
$btnCreate.Location = $System_Drawing_Point
$btnCreate.Name = "btnCreate"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 46
$System_Drawing_Size.Width = 74
$btnCreate.Size = $System_Drawing_Size
$btnCreate.TabIndex = 18
$btnCreate.Text = "Create"
$btnCreate.UseVisualStyleBackColor = $True
$btnCreate.add_Click($btnCreate_OnClick)

$mainForm.Controls.Add($btnCreate)

$lblOU.AutoSize = $True
$lblOU.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 1
$System_Drawing_Point.Y = 303
$lblOU.Location = $System_Drawing_Point
$lblOU.Name = "lblOU"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 99
$lblOU.Size = $System_Drawing_Size
$lblOU.TabIndex = 38
$lblOU.Text = "Organizational Unit:"
$lblOU.add_Click($handler_label2_Click)

$mainForm.Controls.Add($lblOU)

$comboOU.DataBindings.DefaultDataSourceUpdateMode = 0
$comboOU.DropDownHeight = 200
$comboOU.DropDownWidth = 400
$comboOU.FormattingEnabled = $True
$comboOU.IntegralHeight = $False

$comboOU.Items.Add("OU=<OU NAME>,OU=<OU NAME>,OU=<OU NAME>")|Out-Null #---------------------------------------------------------------------------------------#
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 300
$comboOU.Location = $System_Drawing_Point
$comboOU.Name = "comboOU"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 124
$comboOU.Size = $System_Drawing_Size
$comboOU.TabIndex = 15
$comboOU.Text = "Select..."

$mainForm.Controls.Add($comboOU)

$textManager.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 249
$textManager.Location = $System_Drawing_Point
$textManager.Name = "textManager"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textManager.Size = $System_Drawing_Size
$textManager.TabIndex = 13
#$textManager.Text = "Skip this"

$mainForm.Controls.Add($textManager)

$lblManager.AutoSize = $True
$lblManager.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 42
$System_Drawing_Point.Y = 252
$lblManager.Location = $System_Drawing_Point
$lblManager.Name = "lblManager"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 59
$lblManager.Size = $System_Drawing_Size
$lblManager.TabIndex = 36
$lblManager.Text = " *Manager:"

$mainForm.Controls.Add($lblManager)

$lblDisplayName.AutoSize = $True
$lblDisplayName.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 92
$lblDisplayName.Location = $System_Drawing_Point
$lblDisplayName.Name = "lblDisplayName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 75
$lblDisplayName.Size = $System_Drawing_Size
$lblDisplayName.TabIndex = 33
$lblDisplayName.Text = "Display Name:"

$mainForm.Controls.Add($lblDisplayName)

$textDisplayName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 89
$textDisplayName.Location = $System_Drawing_Point
$textDisplayName.Name = "textDisplayName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textDisplayName.Size = $System_Drawing_Size
$textDisplayName.TabIndex = 3

$mainForm.Controls.Add($textDisplayName)

$lblUserNameExample.AutoSize = $True
$lblUserNameExample.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 237
$System_Drawing_Point.Y = 265
$lblUserNameExample.Location = $System_Drawing_Point
$lblUserNameExample.Name = "lblUserNameExample"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 81
$lblUserNameExample.Size = $System_Drawing_Size
$lblUserNameExample.TabIndex = 31
$lblUserNameExample.Text = "*ex. User Name"

$mainForm.Controls.Add($lblUserNameExample)

$lblUserName.AutoSize = $True
$lblUserName.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 38
$System_Drawing_Point.Y = 66
$lblUserName.Location = $System_Drawing_Point
$lblUserName.Name = "lblUserName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 63
$lblUserName.Size = $System_Drawing_Size
$lblUserName.TabIndex = 29
$lblUserName.Text = "User Name:"

$mainForm.Controls.Add($lblUserName)

$textUserName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 63
$textUserName.Location = $System_Drawing_Point
$textUserName.Name = "textUserName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textUserName.Size = $System_Drawing_Size
$textUserName.TabIndex = 2

$mainForm.Controls.Add($textUserName)

$textCopySecurity.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 274
$textCopySecurity.Location = $System_Drawing_Point
$textCopySecurity.Name = "textCopySecurity"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textCopySecurity.Size = $System_Drawing_Size
$textCopySecurity.TabIndex = 14

$mainForm.Controls.Add($textCopySecurity)

$lblCopySecurity.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 20
$System_Drawing_Point.Y = 277
$lblCopySecurity.Location = $System_Drawing_Point
$lblCopySecurity.Name = "lblCopySecurity"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 82
$lblCopySecurity.Size = $System_Drawing_Size
$lblCopySecurity.TabIndex = 27
$lblCopySecurity.Text = "*Copy Security:"

$mainForm.Controls.Add($lblCopySecurity)

$lblCompany.AutoSize = $True
$lblCompany.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 47
$System_Drawing_Point.Y = 225
$lblCompany.Location = $System_Drawing_Point
$lblCompany.Name = "lblCompany"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 54
$lblCompany.Size = $System_Drawing_Size
$lblCompany.TabIndex = 24
$lblCompany.Text = "Company:"

$mainForm.Controls.Add($lblCompany)

$textCompany.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 222
$textCompany.Location = $System_Drawing_Point
$textCompany.Name = "textCompany"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textCompany.Size = $System_Drawing_Size
$textCompany.TabIndex = 10
$TextCompany.Text = "Company Name Here"                         #---------------------------------------------------------------------------------------#

$mainForm.Controls.Add($textCompany)

$lblDepartment.AutoSize = $True
$lblDepartment.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 36
$System_Drawing_Point.Y = 199
$lblDepartment.Location = $System_Drawing_Point
$lblDepartment.Name = "lblDepartment"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 65
$lblDepartment.Size = $System_Drawing_Size
$lblDepartment.TabIndex = 22
$lblDepartment.Text = "Department:"

$mainForm.Controls.Add($lblDepartment)

$textDepartment.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 196
$textDepartment.Location = $System_Drawing_Point
$textDepartment.Name = "textDepartment"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textDepartment.Size = $System_Drawing_Size
$textDepartment.TabIndex = 9
#$textDepartment.Text = "Skip this"

$mainForm.Controls.Add($textDepartment)

$textMobilePhone.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 142
$textMobilePhone.Location = $System_Drawing_Point
$textMobilePhone.Name = "textMobilePhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textMobilePhone.Size = $System_Drawing_Size
$textMobilePhone.TabIndex = 7
$TextMobilePhone.Text = "None"

$mainForm.Controls.Add($textMobilePhone)

$lblMobilePhone.AutoSize = $True
$lblMobilePhone.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 26
$System_Drawing_Point.Y = 145
$lblMobilePhone.Location = $System_Drawing_Point
$lblMobilePhone.Name = "lblMobilePhone"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 75
$lblMobilePhone.Size = $System_Drawing_Size
$lblMobilePhone.TabIndex = 19
$lblMobilePhone.Text = "Mobile Phone:"

$mainForm.Controls.Add($lblMobilePhone)

$lblJobTitle.AutoSize = $True
$lblJobTitle.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 51
$System_Drawing_Point.Y = 173
$lblJobTitle.Location = $System_Drawing_Point
$lblJobTitle.Name = "lblJobTitle"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 50
$lblJobTitle.Size = $System_Drawing_Size
$lblJobTitle.TabIndex = 14
$lblJobTitle.Text = "Job Title:"
$lblJobTitle.add_Click($handler_lblTitle_Click)

$mainForm.Controls.Add($lblJobTitle)

$textJobTitle.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 170
$textJobTitle.Location = $System_Drawing_Point
$textJobTitle.Name = "textJobTitle"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textJobTitle.Size = $System_Drawing_Size
$textJobTitle.TabIndex = 8
#$textJobTitle.Text = "Skip this"

$mainForm.Controls.Add($textJobTitle)

$comboLocation.DataBindings.DefaultDataSourceUpdateMode = 0
$comboLocation.Items.Add("<CityName>")|Out-Null                 #---------------------------------------------------------------------------------------------------------#
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 115
$comboLocation.Location = $System_Drawing_Point
$comboLocation.Name = "comboLocation"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 21
$System_Drawing_Size.Width = 124
$comboLocation.Size = $System_Drawing_Size
$comboLocation.Sorted = $True
$comboLocation.TabIndex = 4
$comboLocation.Text = "Select..."

$mainForm.Controls.Add($comboLocation)

$lblLocation.AutoSize = $True
$lblLocation.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 13
$System_Drawing_Point.Y = 118
$lblLocation.Location = $System_Drawing_Point
$lblLocation.Name = "lblLocation"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 88
$lblLocation.Size = $System_Drawing_Size
$lblLocation.TabIndex = 7
$lblLocation.Text = "Branch Location:"

$mainForm.Controls.Add($lblLocation)

$textLastName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 37
$textLastName.Location = $System_Drawing_Point
$textLastName.Name = "textLastName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textLastName.Size = $System_Drawing_Size
$textLastName.TabIndex = 1

$mainForm.Controls.Add($textLastName)

$lblLastName.AutoSize = $True
$lblLastName.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 40
$System_Drawing_Point.Y = 40
$lblLastName.Location = $System_Drawing_Point
$lblLastName.Name = "lblLastName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 61
$lblLastName.Size = $System_Drawing_Size
$lblLastName.TabIndex = 4
$lblLastName.Text = "Last Name:"

$mainForm.Controls.Add($lblLastName)

$lblFirstName.AutoSize = $True
$lblFirstName.DataBindings.DefaultDataSourceUpdateMode = 0

$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 40
$System_Drawing_Point.Y = 14
$lblFirstName.Location = $System_Drawing_Point
$lblFirstName.Name = "lblFirstName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 13
$System_Drawing_Size.Width = 60
$lblFirstName.Size = $System_Drawing_Size
$lblFirstName.TabIndex = 3
$lblFirstName.Text = "First Name:"
$lblFirstName.add_Click($handler_label1_Click)

$mainForm.Controls.Add($lblFirstName)

$textFirstName.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 107
$System_Drawing_Point.Y = 11
$textFirstName.Location = $System_Drawing_Point
$textFirstName.Name = "textFirstName"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 20
$System_Drawing_Size.Width = 124
$textFirstName.Size = $System_Drawing_Size
$textFirstName.TabIndex = 0

$mainForm.Controls.Add($textFirstName)

#endregion Generated Form Code

#Save the initial state of the form
$InitialFormWindowState = $mainForm.WindowState
#Init the OnLoad event to correct the initial state of the form
$mainForm.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$mainForm.ShowDialog()| Out-Null

} #End Function

$showWindowAsync = Add-Type –memberDefinition @” 
[DllImport("user32.dll")] 
public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow); 
“@ -name “Win32ShowWindowAsync” -namespace Win32Functions –passThru
 
function Show-PowerShell() { 
     [void]$showWindowAsync::ShowWindowAsync((Get-Process -Id $PId).MainWindowHandle, 10) 
}
 
function Hide-PowerShell() { 
    [void]$showWindowAsync::ShowWindowAsync((Get-Process -Id $PId).MainWindowHandle, 2) 
}

# Hide the console window
#Hide-PowerShell

#Call the Function
GenerateForm