Add-Type -AssemblyName System.Drawing
# Define the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "הרשאות"
$form.Size = New-Object System.Drawing.Size(1000, 850)  # Adjust the size here
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
# Set the background image
$imagePath = "G:\ג. תחום מרכז שירותי תקשוב והרשאות\HelpDesk\Adan\new.jpeg"  # Replace with the correct path
if (Test-Path $imagePath) {
    try {
        $backgroundImage = New-Object System.Drawing.Bitmap($imagePath)
        $form.BackgroundImage = $backgroundImage
        Write-Host "Image file found at: $imagePath"
    } catch {
        Write-Host "Error loading image: $_"
    }
} else {
    Write-Host "Image file not found at: $imagePath"
}
# Set the background color
$form.BackColor = "orange"
# Show the form

$buttonM = New-Object System.Windows.Forms.Button
$buttonM.Location = New-Object System.Drawing.Point(357, 120)  # Adjust the positioning as needed
$buttonM.Size = New-Object System.Drawing.Size(65, 25)  # Adjust the size as needed
$buttonM.Text = "חחי"
$form.Controls.Add($buttonM)
# Define button for 'b'
$buttonB = New-Object System.Windows.Forms.Button
$buttonB.Location = New-Object System.Drawing.Point(425, 120)  # Adjust the positioning as needed
$buttonB.Size = New-Object System.Drawing.Size(65, 25)  # Adjust the size as needed
$buttonB.Text = "פרטי"
$form.Controls.Add($buttonB)



$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Location = New-Object System.Drawing.Point(10, 10)
$labelUsername.Size = New-Object System.Drawing.Size(350, 20)
$labelUsername.Text = "Enter the username:"
$labelUsername.ForeColor = [System.Drawing.Color]::Black
$labelUsername.Font = New-Object System.Drawing.Font("Arial",16, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($labelUsername)
# Define textbox for username input
$textBoxUsername = New-Object System.Windows.Forms.TextBox
$textBoxUsername.Location = New-Object System.Drawing.Point(20, 45)
$textBoxUsername.Size = New-Object System.Drawing.Size(280, 20)
$form.Controls.Add($textBoxUsername)
# Define button to check Ofek permissions
$buttonCheckOfek = New-Object System.Windows.Forms.Button
$buttonCheckOfek.Location = New-Object System.Drawing.Point(20, 75)
$buttonCheckOfek.Size = New-Object System.Drawing.Size(150, 30)
$buttonCheckOfek.Text = "Ofek"
$buttonCheckOfek.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckOfek.BackColor = "orange"
$buttonCheckOfek.ForeColor = "White"
$buttonCheckOfek.FlatStyle = "Flat"
$form.Controls.Add($buttonCheckOfek)
# Define button to check Intune permissions
$buttonCheckIntune = New-Object System.Windows.Forms.Button
$buttonCheckIntune.Location = New-Object System.Drawing.Point(180, 75)
$buttonCheckIntune.Size = New-Object System.Drawing.Size(150, 30)
$buttonCheckIntune.Text = "Intune"
$buttonCheckIntune.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckIntune.BackColor = "orange"
$buttonCheckIntune.ForeColor = "White"
$buttonCheckIntune.FlatStyle = "Flat"
$form.Controls.Add($buttonCheckIntune)
# Define button to check remote applications permissions
$buttonCheckRemoteApps = New-Object System.Windows.Forms.Button
$buttonCheckRemoteApps.Location = New-Object System.Drawing.Point(340, 75)
$buttonCheckRemoteApps.Size = New-Object System.Drawing.Size(150, 30)
$buttonCheckRemoteApps.Text = "Remote Apps"
$buttonCheckRemoteApps.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckRemoteApps.BackColor = "orange"
$buttonCheckRemoteApps.ForeColor = "White"
$buttonCheckRemoteApps.FlatStyle = "Flat"
$form.Controls.Add($buttonCheckRemoteApps)
# Define button to check Outlook permissions
$buttonCheckOutlook = New-Object System.Windows.Forms.Button
$buttonCheckOutlook.Location = New-Object System.Drawing.Point(500, 75)
$buttonCheckOutlook.Size = New-Object System.Drawing.Size(150, 30)
$buttonCheckOutlook.Text = "Outlook"
$buttonCheckOutlook.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckOutlook.BackColor = "orange"
$buttonCheckOutlook.ForeColor = "White"
$buttonCheckOutlook.FlatStyle = "Flat"
$form.Controls.Add($buttonCheckOutlook)
# Define button to check Ad permissions
$buttonCheckinfo = New-Object System.Windows.Forms.Button
$buttonCheckinfo.Location = New-Object System.Drawing.Point(660, 75)
$buttonCheckinfo.Size = New-Object System.Drawing.Size(150, 30)
$buttonCheckinfo.Text = "Info"
$buttonCheckinfo.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonCheckinfo.BackColor = "orange"
$buttonCheckinfo.ForeColor = "White"
$buttonCheckinfo.FlatStyle = "Flat"
$form.Controls.Add($buttonCheckinfo)

$buttonGeneratePassword = New-Object System.Windows.Forms.Button
$buttonGeneratePassword.Location = New-Object System.Drawing.Point(820, 75)  # Adjust the positioning as needed
$buttonGeneratePassword.Size = New-Object System.Drawing.Size(150, 30)  # Adjust the size as needed
$buttonGeneratePassword.Text = "Generate Password"
$buttonGeneratePassword.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonGeneratePassword.BackColor = "orange"
$buttonGeneratePassword.ForeColor = "White"
$buttonGeneratePassword.FlatStyle = "Flat"
$form.Controls.Add($buttonGeneratePassword)



   
function Generate-RandomPassword {
    # Define arrays of characters for letters, numbers, and special characters
    $letters = -join ((97..122) | Get-Random -Count 2 | ForEach-Object { [char]$_ })
    $capitalLetter = -join ((65..90) | Get-Random -Count 1 | ForEach-Object { [char]$_ })
    $numbers = Get-Random -Minimum 100 -Maximum 999
    $specialChars = '!@#$%*+.?'
    $specialCharIndex = Get-Random -Minimum 0 -Maximum ($specialChars.Length - 1)
    $specialChar = $specialChars[$specialCharIndex]
    
    # Shuffle the characters
    $characters = $capitalLetter + $letters + $numbers + $specialChar -split '(?<=.)' | Sort-Object {Get-Random}
    
    # Create the password
    $password = -join $characters
    
    return $password
}

$buttonCheckinfo.Add_Click({
    
    $USER = $textboxUsername.Text
    $UserProperties = Get-ADUser -Identity $USER -Properties "PasswordLastSet", "CanonicalName", "msDS-UserPasswordExpiryTimeComputed" , "PasswordNeverExpires" , "MobilePhone", "GivenName", "Surname" , "UserPrincipalName", "PasswordNeverExpires", "Enabled", "Created", "pwdLastSet", "LastLogonDate", "MemberOf", "LockedOut"
    $PasswordExpirationDate = [datetime]::FromFileTime($UserProperties."msDS-UserPasswordExpiryTimeComputed")
   # $TimeLeft = ($PasswordExpirationDate - (Get-Date)).Days
     $TimeLeft = 5
    # Check if the user must change password at next logon
    $mustChangePassword = $UserProperties.pwdLastSet -eq 0
    # Get the last logon date
    $lastLogonDate = $UserProperties.LastLogonDate
    # Get the groups the user is a member of that are "Microsoft_M365F3_Users-G" or "Microsoft_M365E5_Users-G"
    $memberOf = $UserProperties.MemberOf | Where-Object {$_ -like "Microsoft_M365F3_Users-G" -or $_ -like "Microsoft_M365E5_Users-G"}
    # Check if the user is a member of "FireGlassUsersBlockInternet-G"
    $isMemberOfFireGlassUsersBlockInternet = $UserProperties.MemberOf -contains "FireGlassUsersBlockInternet-G"
    # Convert the result to a string ("Yes" or "No")
    # $isMemberOfFireGlassUsersBlockInternetString = If ($isMemberOfFireGlassUsersBlockInternet) {"No"} 
    # Else {"Yes"}
    # Format the dates to Day/Month/Year
    $formattedPasswordExpired = $UserProperties.PasswordLastSet.ToString("dd/MM/yyyy")
    $formattedPasswordChanged = $UserProperties.PasswordLastSet.ToString("dd/MM/yyyy")
    $formattedUserCreatedDate = $UserProperties.Created.ToString("dd/MM/yyyy")
    $formattedPasswordExpirationDate = $PasswordExpirationDate.ToString("dd/MM/yyyy")
    $formattedLastLogonDate = $lastLogonDate.ToString("dd/MM/yyyy")
    # Citrix User Information
    $CTXUserInfo = Get-BrokerMachine -AdminAddress winsrvp0118:80 -Filter "(((SessionUserName -like `"NET\$User`")) -and (SessionSupport -eq `"SingleSession`"))" -MaxRecordCount 50000 | select -Property *
    $CTXMachineName = $CTXUserInfo.MachineName
    $CTXCatalogName = $CTXUserInfo.CatalogName
    $global:SessionStartTime = $CTXUserInfo.SessionStartTime
    $CTXLastLogonOnThisMachine = $CTXUserInfo.LastRegistrationTime.ToString("dd/MM/yyyy, HH:MM") # in case of multiple connections, this will be an F'with.
    $CTXSessionState = $CTXUserInfo.SessionState #what if the user is connected to several machines?
    
    
    $textBoxResult.Text = (
        "UserName:  $USER`r`n" +
        "-------------`r`n" +
        "When Password Changed:   $formattedPasswordChanged`r`n`n" +
        "What's the Password Expiration Date:   $formattedPasswordExpirationDate`r`n`n" +
        "Time Left to Expiration:   $TimeLeft days`r`n`n"+
        +"   $($UserProperties.MemberOf | Where-Object {$_ -like "Microsoft_M365F3_Users-G" -or $_ -like "Microsoft_M365E5_Users-G"})`r`n`n" +
        "Mobile Number:   $($UserProperties.MobilePhone)`r`n`n" +
        "When User Created:   $formattedUserCreatedDate`r`n`n" +
        "Full Name:   $($UserProperties.GivenName)" +
        "  $($UserProperties.Surname)`r`n`n" +
        "User Logon Name (UPN):   $($UserProperties.UserPrincipalName)`r`n`n" +
        "The Is User Locked: $($UserProperties.LockedOut)`r`n`n" +
    
        "---------------------------------`r`n`n"+
        "Director :   `r`n`n"+
        "Machine Name: $CTXMachineName`r`n`n"+
        "Catalog Name: $CTXCatalogName`r`n`n"+
        
        "Session Start Time $global:SessionStartTime`r`n`n"  
    )
    $textBoxResult.ForeColor = [System.Drawing.Color]::white
    $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
    $textBoxResult.BackColor = "black"
 
}) 

    
    # Values to check for Intune applications
 


$buttonCheckIntune.Add_Click({
    $username = $textBoxUsername.Text
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("Enter Username!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Values to check for Intune applications
    $valuesToCheckIntune = @(
      
        "EMM_Application_Authenticator",
        "EMM_Application_Edge",
        "EMM_Application_Office",
        "EMM_Application_Outlook",
        "EMM_Enrollment_BYOD",
        "EMM_Web_Default"
        "Microsoft_M365E5_Users-G",
        "Microsoft_M365F3_Users-G"
        
     
    )
    
    try {
        # Get the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties MemberOf -Server "DC04"
        
        # Check if the user is a member of any groups for Intune applications
        if ($user.MemberOf) {
            # Extract group names
            $groupNames = $user.MemberOf | ForEach-Object { ($_ -split ",", 2)[0] -replace '^CN=', '' }
            $notFoundPermissions = @()
            foreach ($value in $valuesToCheckIntune) {
                if (!($groupNames -contains $value)) {
                    $notFoundPermissions += $value
                }
            }
             if ($notFoundPermissions.Count -eq 1 -and $notFoundPermissions -eq "Microsoft_M365F3_Users-G" ){
                $textBoxResult.Text = " כל ההרשאות נמצאות ללקוח."
                $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += "  E5 לא משפיע- יש רשיון , F3  חסר לו רישיון*"
                $textBoxResult.ForeColor = [System.Drawing.Color]::Green
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)}
           elseif ($notFoundPermissions.Count -eq 1 -and $notFoundPermissions -eq "Microsoft_M365E5_Users-G" ){
             
                $textBoxResult.Text = " כל ההרשאות נמצאות ללקוח."
              $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += [Environment]::NewLine
                 $textBoxResult.Text += " F3 לא משפיע- יש רשיון , E5  חסר לו רישיון*"
               $textBoxResult.ForeColor = [System.Drawing.Color]::Green
               $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)}
             else {
                $textBoxResult.Text = "ההרשאות האלה לא קיימות:`n"
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text +=($notFoundPermissions -join "`r`n")  # Join the rest with newline characters
                $textBoxResult.ForeColor = [System.Drawing.Color]::Red
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
            }
        } else {
            $textBoxResult.Text = "User $($user.SamAccountName) is not a member of any groups for Intune applications."
            $textBoxResult.ForeColor = [System.Drawing.Color]::Red
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        }
    } catch {
        if ($_.Exception.Message -match "Cannot find an object with identity") {
            $textBoxResult.Text = "User not found."
        } else {
            $textBoxResult.Text = "Error: $_"
        }
        $textBoxResult.ForeColor = [System.Drawing.Color]::Red
        $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    }
})
#This code will now include the missing permissions in the result text box if any are found, excluding "Microsoft_M365E5_Users-G" and "Microsoft_M365F3_Users-G" when both are missing. Let me know if you need further adjustments!
# Event handler for checking remote applications permissions


$buttonB.Add_Click({
    $username = $textBoxUsername.Text
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("Enter Username!", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Values to check for permissions
    $permissionsToCheck = @(
        "Intune_Remote_Citrix-G",
        "InTune_AutoEnroll",
        "Intune_Workspace_Install-G",
        "SSLVPN_Remote"
    )
    
    # Initialize counters
    $trueCount = 0
    $falseCount = 0
    
    try {
        # Get the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties MemberOf -Server "DC04"
        
        # Check if the user is a member of any groups for the permissions
        if ($user.MemberOf) {
            # Extract group names
            $groupNames = $user.MemberOf | ForEach-Object { ($_ -split ",", 2)[0] -replace '^CN=', '' }
            
            # Iterate over permissions to check
            foreach ($permission in $permissionsToCheck) {
                $isMember = $groupNames -contains $permission
                
                # Update counters based on condition
                if ($isMember) {
                    $trueCount++
                } else {
                    $falseCount++
                }
            }
            
            # Check if the conditions are met
            if ($trueCount -eq 1 -and $falseCount -eq 3) {
                
                $textBoxResult.Text = "יש ללקוח את כל ההרשאות  "
                $textBoxResult.ForeColor = [System.Drawing.Color]::green
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
            } else {
                $textBoxResult.Text = " SSLVPN_Remote :חסר ללקוח הרשאת "
                $textBoxResult.ForeColor = [System.Drawing.Color]::red
                 $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
            }
            
      
       
        } else {
            $textBoxResult.Text = "User $($user.SamAccountName) is not a member of any groups."
            $textBoxResult.ForeColor = [System.Drawing.Color]::Red
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        }
    } catch {
        if ($_.Exception.Message -match "Cannot find an object with identity") {
            $textBoxResult.Text = "User not found."
        } else {
            $textBoxResult.Text = "Error: $_"
        }
        $textBoxResult.ForeColor = [System.Drawing.Color]::Red
        $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    }
})
$buttonm.add_click({
    $username = $textboxusername.text
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("Enter Username !  ! ! !", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Values to check for remote applications
    $valuesToCheck = @(
        
        "Intune_Remote_Citrix-G",
        "InTune_AutoEnroll",
        "Intune_Workspace_Install-G",
        "SSLVPN_Remote"
        
    )
    
    try {
        # Get the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties MemberOf -Server "DC04"
        
        # Check if the user is a member of any groups for remote applications
        if ($user.MemberOf) {
            # Extract group names
            $groupNames = $user.MemberOf | ForEach-Object { ($_ -split ",", 2)[0] -replace '^CN=', '' }
            $notFoundPermissions = @()
            foreach ($value in $valuesToCheck) {
                if (!($groupNames -contains $value)) {
                    $notFoundPermissions += $value
                }
            }
            if ($notFoundPermissions.Count -eq 0) {
                $textBoxResult.Text = "כל ההרשאות של החיבור מרחוק נמצאות אצל הלקוח."
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.ForeColor = [System.Drawing.Color]::Green
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
            } else {
                $textBoxResult.Text = "ההרשאות האלה לא קיימות:`n"
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text +=($notFoundPermissions -join "`r`n")  # Join the rest with newline characters
                #$textBoxResult.Text += $message
                $textBoxResult.ForeColor = [System.Drawing.Color]::Red
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
            }
        } else {
            $textBoxResult.Text = "User $($user.SamAccountName) is not a member of any groups for remote applications."
            $textBoxResult.ForeColor = [System.Drawing.Color]::Red
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        }
    } catch {
        if ($_.Exception.Message -match "Cannot find an object with identity") {
            $textBoxResult.Text = "User not found."
        } else {
            $textBoxResult.Text = "Error: $_"
        }
        $textBoxResult.ForeColor = [System.Drawing.Color]::Red
        $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    }
})
# Event handler for checking Ofek permissions
$buttonCheckOfek.Add_Click({
    $username = $textBoxUsername.Text
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("Enter Username !  ! ! !", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    
    # Values to check for Ofek applications
    $valuesToCheckOfek = @(
        "OfekOfficeUsers-G",
        "OfekUsersApp-G",
        "GoogleChromeNoSaveAs-G",
        "IfnusersApp-g"
    )
    
    try {
        # Get the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties MemberOf -Server "DC04"
        
        # Check if the user is a member of any groups for Ofek applications
        if ($user.MemberOf) {
            # Extract group names
            $groupNames = $user.MemberOf | ForEach-Object { ($_ -split ",", 2)[0] -replace '^CN=', '' }
            $notFoundPermissions = @()
            foreach ($value in $valuesToCheckOfek) {
                if (!($groupNames -contains $value)) {
                    $notFoundPermissions += $value
                }
            }
            if ($notFoundPermissions.Count -eq 0) {
                $textBoxResult.Text = "כל ההרשאות לאופק קיימות."
                $textBoxResult.ForeColor = [System.Drawing.Color]::Green
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
                
            } else {
                $textBoxResult.Text = "ההרשאות האלה לא קיימות:`n"
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text += [Environment]::NewLine
                $textBoxResult.Text +=($notFoundPermissions -join "`r`n")  # Join the rest with newline characters
                #$textBoxResult.Text += $message
                $textBoxResult.ForeColor = [System.Drawing.Color]::Red
                $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
            }
        } else {
            $textBoxResult.Text = "User $($user.SamAccountName) is not a member of any groups for Ofek applications."
            $textBoxResult.ForeColor = [System.Drawing.Color]::Red
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        }
    } catch {
        if ($_.Exception.Message -match "Cannot find an object with identity") {
            $textBoxResult.Text = "User not found."
        } else {
            $textBoxResult.Text = "Error: $_"
        }
        $textBoxResult.ForeColor = [System.Drawing.Color]::Red
        $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    }
})
$buttonCheckOutlook.Add_Click({
    $username = $textBoxUsername.Text
    if ([string]::IsNullOrEmpty($username)) {
        [System.Windows.Forms.MessageBox]::Show("Enter Username !  ! ! !", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
$outlookPermissions = @(
        "Microsoft_M365E5_Users-G",
        "Microsoft_O365E5_Exchange_M365E5_users-G",
        "Microsoft_M365F3_Users-G"
    )
    
    try {
        # Get the user object from Active Directory
        $user = Get-ADUser -Identity $username -Properties MemberOf -Server "DC04"
        
        # Check if the user is a member of any groups
        if ($user.MemberOf) {
            # Extract group names
            $groupNames = $user.MemberOf | ForEach-Object { ($_ -split ",", 2)[0] -replace '^CN=', '' }
            $outlookLicense = $null
            
            foreach ($permission in $outlookPermissions) {
                if ($groupNames -contains $permission) {
                    if ($permission -eq "Microsoft_M365E5_Users-G" -or $permission -eq "Microsoft_O365E5_Exchange_M365E5_users-G") {
                        $outlookLicense = "E5"
                    } elseif ($permission -eq "Microsoft_M365F3_Users-G") {
                        $outlookLicense = "F3"
                    }
                    break
                }
            }
            
            if ($outlookLicense) {
                if ($outlookLicense -eq "E5") {
                    $textBoxResult.Text = "E5=יש ללקוח הרשאה בדפדפן ובמחשב"
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += "`n Microsoft_O365E5_Exchange_M365E5_users-G"
                    $textBoxResult.Text += [Environment]::NewLine
                   
                    $textBoxResult.Text += "`n Microsoft_M365E5_Users-G"
       
                   
                    $textBoxResult.Text += [Environment]::NewLine
                } elseif ($outlookLicense -eq "F3") {
                    $textBoxResult.Text = "F3=יש ללקוח הרשאה רק בדפדפן"
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                    $textBoxResult.Text += [Environment]::NewLine
                
                }    
            } else {
                $textBoxResult.Text = "User does not have Outlook license."
            }
            
            $mailOnPremStatus = if ($groupNames -contains "MailOnPrem") {"מקומי."} else {"ענן."}
            $textBoxResult.Text += [Environment]::NewLine
            $textBoxResult.Text += [Environment]::NewLine
            $textBoxResult.Text += "`n$mailOnPremStatus"
            $textBoxResult.ForeColor = [System.Drawing.Color]::Green  # Set text color to black
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        } else {
            $textBoxResult.Text = "User $($user.SamAccountName) is not a member of any groups."
            $textBoxResult.ForeColor = [System.Drawing.Color]::Red
            $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
        }
    } 
    
    
    catch {
        if ($_.Exception.Message -match "Cannot find an object with identity") {
            $textBoxResult.Text = "User not found."
        } else {
            $textBoxResult.Text = "Error: $_"
        }
        $textBoxResult.ForeColor = [System.Drawing.Color]::Red
        $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    }
    }

  )  
$buttonGeneratePassword.Add_Click({
    $randomPassword = Generate-RandomPassword
    $textBoxResult.Text = "Generated Password: $randomPassword"
    $textBoxResult.ForeColor = [System.Drawing.Color]::white
    $textBoxResult.Font = New-Object System.Drawing.Font("Arial", 18, [System.Drawing.FontStyle]::Bold)
})
# Define button to check Intune permissions
# Define the remaining buttons and functionalities here...
# Define textbox for result display





$textBoxResult = New-Object System.Windows.Forms.TextBox
$textBoxResult.Location = New-Object System.Drawing.Point(40, 160)
$textBoxResult.Size = New-Object System.Drawing.Size(800, 650)  # Adjust the size here
$textBoxResult.Multiline = $true
$textBoxResult.ReadOnly = $true
$textBoxResult.BackColor = [System.Drawing.Color]::Black  # Set background color to white
$form.Controls.Add($textBoxResult)
# Show the form
$form.ShowDialog() | Out-Null