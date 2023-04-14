#. (Join-Path $PSScriptRoot 'PowerShellFormProject1.designer.ps1')

##$PowerShellFormProject1.ShowDialog()

[void][System.Reflection.Assembly]::LoadWithPartialName( "Microsoft.VisualBasic")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
##Main Box
$ApplicationForm = New-Object System.Windows.Forms.Form
$ApplicationForm.StartPosition = "CenterScreen"
$ApplicationForm.Topmost = $false 
$ApplicationForm.Size = "800,600"
$ApplicationForm.Text = "AD User Manager MCP"

$objIcon = New-Object system.drawing.icon("MCP.ico")
$ApplicationForm.Icon = $objIcon

$FormTabControl = New-object System.Windows.Forms.TabControl 
$FormTabControl.Size = "755,475" 
$FormTabControl.Location = "25,25" 
$ApplicationForm.Controls.Add($FormTabControl)

function VerifyJobTitle{
    $tempUserName = $env:UserName
    $JT = (Get-AdUser -Identity $tempUserName -Property Title | Select-Object Title).title
 
    return $JT
    }

$vTitle = VerifyJobTitle
$newTempUserName = $env:UserName
$tempuserlocation = (Get-AdUser -Identity $newTempUserName -Property l).l

#########################################################################################################
###################################----- Refresh The Form-----###########################################
#########################################################################################################
##Function MakeNewForm {
	#$Form.Close()
	#$Form.Dispose()
	#MakeForm
#}

#########################################################################################################
####################################----- Features to Add-----###########################################
#########################################################################################################
#  Ability to clear cache - both browsers and "Deep Cache" if possible - just a button press and username search 
#  Add a link to the phone Number on Tab 6, so that when they click it, it'll take them to Softphone and dial
#
#
#


###############################
#Logic to account For:
#
#   1) Trainers assigning per location
#

#########################################################################################################
###############################----- Check For Expired Accounts-----#####################################
#########################################################################################################
$ExpiredUsers = Search-ADAccount -Server ag-ul -AccountExpired -UsersOnly -ResultPageSize 2000 -resultSetSize $null| Select-Object Name, SamAccountName, DistinguishedName, AccountExpirationDate

$ExpiredUsers | ForEach-Object {
   Remove-ADUser -Identity $ExpiredUsers.SamAccountName -confirm:$false
}


#########################################################################################################
##################################----- UserList Declaration-----########################################
#########################################################################################################

$UserList1 = New-Object system.Windows.Forms.ComboBox
$UserList1.text = “User List”
$UserList1.width = 170
$UserList1.autosize = $true
$UserList1.location = New-Object System.Drawing.Point(10,50)
$UserList1.Font = ‘Microsoft Sans Serif,10’

$UserList2 = New-Object system.Windows.Forms.ComboBox
$UserList2.text = “User List”
$UserList2.width = 170
$UserList2.autosize = $true
$UserList2.location = New-Object System.Drawing.Point(10,50)
$UserList2.Font = ‘Microsoft Sans Serif,10’

$UserList3 = New-Object system.Windows.Forms.ComboBox
$UserList3.text = “User List”
$UserList3.width = 170
$UserList3.autosize = $true
$UserList3.location = New-Object System.Drawing.Point(10,50)
$UserList3.Font = ‘Microsoft Sans Serif,10’

$UserList4 = New-Object system.Windows.Forms.ComboBox
$UserList4.text = “User List”
$UserList4.width = 170
$UserList4.autosize = $true
$UserList4.location = New-Object System.Drawing.Point(10,50)
$UserList4.Font = ‘Microsoft Sans Serif,10’

$UserList5 = New-Object system.Windows.Forms.ComboBox
$UserList5.text = “User List”
$UserList5.width = 170
$UserList5.autosize = $true
$UserList5.location = New-Object System.Drawing.Point(10,50)
$UserList5.Font = ‘Microsoft Sans Serif,10’

$UserList7 = New-Object system.Windows.Forms.ComboBox
$UserList7.text = “User List”
$UserList7.width = 170
$UserList7.autosize = $true
$UserList7.location = New-Object System.Drawing.Point(10,50)
$UserList7.Font = ‘Microsoft Sans Serif,10’


#########################################################################################################
#################################----- UserList PowerShell Cmds-----#####################################
#########################################################################################################
function UserLists{
    $UserList1.Items.Clear()
    $UserList2.Items.Clear()
    $UserList3.Items.Clear()
    $UserList4.Items.Clear()
    $UserList5.Items.Clear()
    $UserList7.Items.Clear()
# Add the items in the dropdown list
$AllUsers = Get-ADUser -Filter "(Title -like '*CS Team Operator*') -Or (Title -like '*Rexburg Team Mentor*') -Or (Title -like '*Ogden Team Mentor*') -Or (Title -like '*Cedar Team Mentor*')"| Sort-Object | Select-Object name

$AllUsers | ForEach-Object {
    $UserList1.Items.Add($_.name)
    $UserList2.Items.Add($_.name)
    $UserList3.Items.Add($_.name)
    
    #$UserList5.Items.Add($_.name)
}

#$AllUsers2 = Get-ADUser -Filter "(Title -like '*CS Team Operator*') -Or (Title -like '*Rexburg Team Mentor*') -Or (Title -like '*Ogden Team Mentor*') -Or (Title -like '*Cedar Team Mentor*') -Or (Title -like '*Rexburg Team Coach*') -Or (Title -like '*Ogden Team Coach*') -Or (Title -like '*Cedar Team Coach*')"| Sort-Object | Select-Object name

#$AllUsers2 | ForEach-Object {
    
#}

#$AllUsers3 = Get-ADUsers -Filter * -SearchBase "OU=Users"| Sort-Object | Select-Object name
$AllUsers3 = Get-ADUser -Filter "(Title -like '*')"| Sort-Object | Select name
$AllUsers3 | ForEach-Object {
    $UserList4.Items.Add($_.name)
    $UserList5.Items.Add($_.name)
    $UserList7.Items.Add($_.name)
}
}

UserLists
#########################################################################################################
###################################----- X-Tab PS Scripts-----###########################################
#########################################################################################################

###Get-ADUser -Filter "Title -like '*CS Team Operator*'" | Sort-Object | Select name   ## Used for listing all operators || can change to Select sAMAccountName


#########################################################################################################
###################################----- X-Tab WebSources-----###########################################
#########################################################################################################

####https://geekdudes.wordpress.com/2021/04/21/powershell-import-csv-file-to-combobox-and-get-selected-value-into-variable/
<#
https://social.technet.microsoft.com/Forums/en-US/42a7d43d-942b-485e-94ce-b10d497b7966/powershell-quotclear-or-remove-text-boxquot-button?forum=winserverpowershell
https://stackoverflow.com/questions/11500962/how-to-set-keyboard-focus-to-a-textbox-in-powershell
#>

. (Join-Path $PSScriptRoot 'Tab1-PasswordReset-AccountUnlock.ps1')
. (Join-Path $PSScriptRoot 'Tab2-KeyCardReset.ps1')
. (Join-Path $PSScriptRoot 'Tab3-CreateUser.ps1')
. (Join-Path $PSScriptRoot 'Tab4-PromoteEmployee.ps1')
. (Join-Path $PSScriptRoot 'Tab5-TerminateEmployee.ps1')
. (Join-Path $PSScriptRoot 'Tab6-HelpTab')
. (Join-Path $PSScriptRoot 'Tab7-ModifyEmployeeGroups')
. (Join-Path $PSScriptRoot 'DefaultTab-BadTab')


 





#########################################################################################################
######################################-----Form Control-----#############################################
#########################################################################################################

if ($newTempUserName.substring(0,5) -eq "admin" ){   
    $FormTabControl.Controls.Add($Tab1)
    $FormTabControl.Controls.Add($Tab2)
    $FormTabControl.Controls.Add($Tab3)
    $FormTabControl.Controls.Add($Tab4)
    $FormTabControl.Controls.Add($Tab5)
    $FormTabControl.Controls.Add($Tab7)
    $FormTabControl.Controls.Add($Tab6)
 
}
elseif ($vTitle -eq "$tempuserlocation Team Coach" -Or $vTitle -eq "$tempuserlocation Shift Coordinator" -Or $vTitle -eq "WFM Real Time Analyst"-Or $vTitle -eq "Team Coach"){
    $FormTabControl.Controls.Add($Tab1)
    $FormTabControl.Controls.Add($Tab2)
    $FormTabControl.Controls.Add($Tab6)
}
elseif ($vTitle -eq "Grave Shift Supervisor" -Or $vTitle -eq "Graves Shift Supervisor" -Or $vTitle -eq "Swings Shift Supervisor"){
    $FormTabControl.Controls.Add($Tab1)
    $FormTabControl.Controls.Add($Tab2)
    $FormTabControl.Controls.Add($Tab6)
}
elseif ($vTitle -eq "HR Generalist" -Or $vTitle -eq "Corporate Recruiter" -Or $vTitle -eq "Campus Recruiter" -Or $vTitle -eq "Human Resources Manager"){
    $FormTabControl.Controls.Add($Tab3)
    $FormTabControl.Controls.Add($Tab4)
    $FormTabControl.Controls.Add($Tab7)
    $FormTabControl.Controls.Add($Tab6)
}
elseif ($vTitle -eq "Operations Manager" -Or $vTitle -eq "$tempuserlocation Operations Manager"){
    $FormTabControl.Controls.Add($Tab1)
    $FormTabControl.Controls.Add($Tab2)
    $FormTabControl.Controls.Add($Tab4)
    $FormTabControl.Controls.Add($Tab6)
}

else{

    $FormTabControl.Controls.Add($BadTab)
    $FormTabControl.Controls.Add($Tab6)
 
}

#deleteUsers

# Initlize the form
$ApplicationForm.Add_Shown({$ApplicationForm.Activate()})
[void] $ApplicationForm.ShowDialog()



