
#########################################################################################################
####################################-----Tab 1 Contents-----#############################################
#########################################################################################################

#####---For SC Use---#####

####################-----Objective-----##################################################################
#
#  Allows SC's to change an Employee's Password and/Or Unlock an Employee's account
#
#########################################################################################################

############################################################################
#####################-----Tab 1 Code-----###################################
############################################################################

## Password reset
function reset {
    $displayname = $UserList1.Text
    $Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
    set-adaccountpassword $Username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Avantguard1" -Force)
    Set-ADUser -ChangePasswordAtLogon $true -Identity $Username -Confirm:$false
    [System.Windows.Forms.MessageBox]::Show('Password has been Reset')
}

##Unlock accounts
function Unlock {
    $displayname = $UserList1.Text
    $Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
    Unlock-adaccount -Identity $Username
    $Output.Text = "The account $Username has been unlocked."
    [System.Windows.Forms.MessageBox]::Show("The account for " + $Username + " has been unlocked")
}

$Tab1 = New-object System.Windows.Forms.Tabpage
$Tab1.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab1.UseVisualStyleBackColor = $True 
$Tab1.Name = "Tab1" 
$Tab1.Text = "Password Reset/Unlock Account” 

$InstLabel1 = New-Object System.Windows.Forms.label
$InstLabel1.Location = New-Object System.Drawing.Size(10,30)
$InstLabel1.Size = New-Object System.Drawing.Size(500,20)
$InstLabel1.BackColor = "TRANSPARENT"
$InstLabel1.ForeColor = "BLACK"
$InstLabel1.Text = "Passwords will always be reset to Avantguard1"


$Tab1.Controls.Add($UserList1)

$button1 = New-Object System.Windows.Forms.Button
$button1.BackColor = "#d0caca"
$button1.Location = New-Object System.Drawing.Point(300,120)
$button1.Size = New-Object System.Drawing.Size (200,30)
$button1.Text = "Reset Password"
$button1.Add_Click({ Reset })


if ($newTempUserName.substring(0,5) -eq "admin" -Or $vTitle -eq "HR Generalist" -Or $vTitle -eq "Human Resources Manager"){   
    $Tab1.Controls.Add($button1)
    $Tab1.Controls.Add($InstLabel1)
}
$unlockAccountButton = New-Object System.Windows.Forms.Button
$unlockAccountButton.BackColor = "#d0caca"
$unlockAccountButton.Location = New-Object System.Drawing.Point(300,170)
$unlockAccountButton.Size = New-Object System.Drawing.Size (200,30)
$unlockAccountButton.Text = "Unlock Account"
$unlockAccountButton.Add_Click({ Unlock })
$Tab1.Controls.Add($unlockAccountButton)

