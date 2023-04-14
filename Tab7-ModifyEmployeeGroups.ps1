#########################################################################################################
####################################-----Tab 7 Contents-----#############################################
#########################################################################################################

#####-----For IT Use-----#####

####################-----Objective-----##################################################################
#
#  Allows for easy updating of user groups
#
#########################################################################################################


#########################################################################################################
######################################-----Tab 7 Code-----###############################################
#########################################################################################################

$objTextBox7 = New-Object System.Windows.Forms.TextBox 
$objTextBox7.Multiline = $True;
$objTextBox7.Location = New-Object System.Drawing.Size(10,130) 
$objTextBox7.Size = New-Object System.Drawing.Size(250,130)
$objTextBox7.Scrollbars = "Vertical" 
$objTextBox7.ReadOnly=$true
$objTextBox7.ForeColor = "BLACK"

$objTextBoxAuthOrig7 = New-Object System.Windows.Forms.TextBox 
$objTextBoxAuthOrig7.Multiline = $True;
$objTextBoxAuthOrig7.Location = New-Object System.Drawing.Size(350,130) 
$objTextBoxAuthOrig7.Size = New-Object System.Drawing.Size(380,130)
$objTextBoxAuthOrig7.Scrollbars = "Vertical"
$objTextBoxAuthOrig7.WordWrap = $true
$objTextBoxAuthOrig7.ReadOnly=$true
$objTextBoxAuthOrig7.ForeColor = "BLACK"


##Check Groups of Current User
function CheckGroups {
        $objTextBox7.Text = ""
    $displayname = $UserList7.Text
#$Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
$Username = (Get-ADUser -Filter {name -eq $displayname}).samaccountname
$tempAllGroups = (Get-ADUser $Username –Properties MemberOf).memberof | Get-ADGroup | Select-Object name
    #$objTextBox7.Text += $tempAllGroups
    foreach ($group in $tempAllGroups)
        {
        $objTextBox7.Text += "$group `r`n"
    }
#Get-ADGroup -Identity $groupname | Set-ADGroup -Add @{authOrig=@($temp)}
    #[System.Windows.Forms.MessageBox]::Show($displayname + ' can now email ' + $groupname)
}

##Check AuthOrig of Current Group
function CheckAuthOrig {

    $objTextBoxAuthOrig7.Text = ""
    $displayname = $GroupList.Text
   $tempAllAuthOrig = (Get-ADGroup -Identity $displayname -Property authOrig | Select -Expand authorig)

    foreach ($value in $tempAllAuthOrig)
        {
            $separator = $value.IndexOf(',')
            $instance = $value.Substring(0, $separator)
            $name = $instance.Replace('CN=', '')
    #Write-Host $name
       #Get-ADGroupMember -Identity $group | ForEach-Object {
        $objTextBoxAuthOrig7.Text += $name + "`r`n"#.authorig #+ " `r`n"
    }#}
#Get-ADGroup -Identity $groupname | Set-ADGroup -Add @{authOrig=@($temp)}
    #[System.Windows.Forms.MessageBox]::Show($displayname + ' can now email ' + $groupname)
}


##Recieve Emails
function SendEmail {
    $displayname = $UserList7.Text
    $groupname = $GroupList.Text
$Username = (Get-ADUser -Filter {name -eq $displayname}).samaccountname
$temp = (Get-ADUser -Identity $Username).DistinguishedName
Get-ADGroup -Identity $groupname | Set-ADGroup -Add @{authOrig=@($temp)}
    [System.Windows.Forms.MessageBox]::Show($displayname + ' can now email ' + $groupname)
}
##Receive Emails and Add to Groups in AD
function AddToGroup {
    $displayname = $UserList7.Text
    $groupname = $GroupList.Text
    $Username = (Get-ADUser -Filter {name -eq $displayname}).samaccountname
    Get-ADGroup -Identity $groupname | Add-ADGroupMember -Members $Username
    [System.Windows.Forms.MessageBox]::Show($displayname + ' is now a member of ' + $groupname)
}
##Remove From Group
function RemoveFromGroup {
    $displayname = $UserList7.Text
    $groupname = $GroupList.Text
    $Username = (Get-ADUser -Filter {name -eq $displayname}).samaccountname
    Remove-ADGroupMember -Identity $groupname -Members $Username -confirm:$false
    [System.Windows.Forms.MessageBox]::Show($displayname + ' has been removed from ' + $groupname)
}

$GroupList = New-Object system.Windows.Forms.ComboBox
$GroupList.text = “Group List”
$GroupList.width = 170
$GroupList.autosize = $true
$GroupList.location = New-Object System.Drawing.Point(350,50)
$GroupList.Font = ‘Microsoft Sans Serif,10’

$AllGroups = Get-ADGroup -filter * -properties * | Sort-Object |select SAMAccountName
$AllGroups | ForEach-Object {
    $GroupList.Items.Add($_.SAMAccountName)
}

$Tab7 = New-object System.Windows.Forms.Tabpage
$Tab7.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab7.UseVisualStyleBackColor = $True 
$Tab7.Name = "Tab7" 
$Tab7.Text = "Group Mod” 

$CheckGroupButton7 = New-Object System.Windows.Forms.Button
$CheckGroupButton7.BackColor = "#d0caca"
$CheckGroupButton7.Location = New-Object System.Drawing.Point(190,50)
$CheckGroupButton7.Size = New-Object System.Drawing.Size (120,30)
$CheckGroupButton7.Text = "Check Groups"
$Tab7.Controls.Add($CheckGroupButton7)
$CheckGroupButton7.Add_Click({ CheckGroups })

$CheckAuthOrigButton7 = New-Object System.Windows.Forms.Button
$CheckAuthOrigButton7.BackColor = "#d0caca"
$CheckAuthOrigButton7.Location = New-Object System.Drawing.Point(530,50)
$CheckAuthOrigButton7.Size = New-Object System.Drawing.Size (120,30)
$CheckAuthOrigButton7.Text = "Check AuthOrig"
$Tab7.Controls.Add($CheckAuthOrigButton7)
$CheckAuthOrigButton7.Add_Click({ CheckAuthOrig })


$EmailReceiveButton7 = New-Object System.Windows.Forms.Button
$EmailReceiveButton7.BackColor = "#d0caca"
$EmailReceiveButton7.Location = New-Object System.Drawing.Point(10,320)
$EmailReceiveButton7.Size = New-Object System.Drawing.Size (150,30)
$EmailReceiveButton7.Text = "Send Emails"
##$Tab7.Controls.Add($EmailReceiveButton7)
$EmailReceiveButton7.Add_Click({ SendEmail })

$addButton7 = New-Object System.Windows.Forms.Button
$addButton7.BackColor = "#d0caca"
$addButton7.Location = New-Object System.Drawing.Point(150,320)
$addButton7.Size = New-Object System.Drawing.Size (150,30)
$addButton7.Text = "Add to Group"
$Tab7.Controls.Add($addButton7)
$addButton7.Add_Click({ AddToGroup })

$RemoveButton7 = New-Object System.Windows.Forms.Button
$RemoveButton7.BackColor = "#d0caca"
$RemoveButton7.Location = New-Object System.Drawing.Point(300,320)
$RemoveButton7.Size = New-Object System.Drawing.Size (150,30)
$RemoveButton7.Text = "Remove From Group"
$Tab7.Controls.Add($RemoveButton7)
$RemoveButton7.Add_Click({ RemoveFromGroup })

$Tab7.Controls.Add($UserList7)
$Tab7.Controls.Add($GroupList)
$Tab7.Controls.Add($objTextBox7)
$Tab7.Controls.Add($objTextBoxAuthOrig7)

