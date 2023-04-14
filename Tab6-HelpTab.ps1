#########################################################################################################
####################################-----Tab 6 Contents-----#############################################
#########################################################################################################

#####-----For All Use-----#####

####################-----Objective-----##################################################################
#
#  Adds Links for Any users and a guide for each of the tabs.
#
#########################################################################################################
$Tab6 = New-object System.Windows.Forms.Tabpage
$Tab6.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab6.UseVisualStyleBackColor = $True 
$Tab6.Name = "Tab6" 
$Tab6.Text = "Help” 
#$FormTabControl.Controls.Add($Tab6)

$LinkLabel1 = New-Object System.Windows.Forms.LinkLabel
$LinkLabel1.Location = New-Object System.Drawing.Size(30,50)
$LinkLabel1.Size = New-Object System.Drawing.Size(150,20)
$LinkLabel1.LinkColor = "BLUE"
$LinkLabel1.ActiveLinkColor = "RED"
$LinkLabel1.Text = "Submit A Ticket"
$LinkLabel1.add_Click({[system.Diagnostics.Process]::start("https://agmonitoring-ism.ivanticloud.com/ Jump ")})
$Tab6.Controls.Add($LinkLabel1)

$LinkLabel1 = New-Object System.Windows.Forms.LinkLabel
$LinkLabel1.Location = New-Object System.Drawing.Size(30,80)
$LinkLabel1.Size = New-Object System.Drawing.Size(150,20)
$LinkLabel1.LinkColor = "BLUE"
$LinkLabel1.ActiveLinkColor = "RED"
$LinkLabel1.Text = "IT Help Guide"
$LinkLabel1.add_Click({[system.Diagnostics.Process]::start("https://becklar.sharepoint.com/:w:/r/sites/AGCommon/_layouts/15/Doc.aspx?sourcedoc=%7B13998379-2F4A-4FDC-9B48-3ED08B8F9CF5%7D&file=Welcome%20to%20the%20IT%20help%20guide.docx&action=default&mobileredirect=true
 Jump ")})
$Tab6.Controls.Add($LinkLabel1)

$objLabel = New-Object System.Windows.Forms.label
$objLabel.Location = New-Object System.Drawing.Size(30,110)
$objLabel.Size = New-Object System.Drawing.Size(180,20)
$objLabel.BackColor = "TRANSPARENT"
$objLabel.ForeColor = "BLACK"
$objLabel.Text = "(801)781-1300 or (385)432-8919"
$Tab6.Controls.Add($objLabel)

$objTextBox6 = New-Object System.Windows.Forms.TextBox 
$objTextBox6.Multiline = $True;
$objTextBox6.Location = New-Object System.Drawing.Size(30,130) 
$objTextBox6.Size = New-Object System.Drawing.Size(700,200)
$objTextBox6.Scrollbars = "Vertical" 
$objTextBox6.ReadOnly=$true
$objTextBox6.ForeColor = "BLACK"

$HelpTabOpeningText = "General Guide for Usage: `r`n"

$HelpTab1Text = "`r`n Tab1: 
`n`t To Reset a Password: 
`r`n`t 1. Scroll down to the name of the Operator or Team Mentor and select their name. 
`r`n`t 2. Then click the 'Reset Password' button. The password will be changed to 'Avantguard1'. 
`r`n`t 3. Have the Operator then login to their account with their username and the password 'Avantguard1'.
`r`n`t 4. The Operator will then be prompted to change their password.
`r`n`t`t (Notes: Passwords must have: An upper-case letter, a lower-case letter, a number, and a special
`r`n`t`t charachter, such as '!', '#', '&', '@' or '$'. Passwords cannot contain the name of the Operator nor
`r`n`t`t names of locations, such as state or city names.)
`r`n `r`n 
`n`t To Unlock an Account:
`r`n`t 1. Scroll down to the name of the Operator or Team Mentor and select their name.
`r`n`t 2. Then click the 'Unlock Account' button. Wait approximately one minute and then have the operator login
`r`n`t 3. The account of the operator should now be unlocked. If not, wait 15 minutes, and the Account should unlock on its own
`r`t `r`n "

$HelpTab2Text = "`r`n Tab2: 
`n`t To Reset the Pincode for an Operator's Keycard: 
`r`n`t 1. Scroll down to the name of the Operator or Team Mentor and select their name.
`r`n`t 2. Enter in the empty box the 4 digit pin that the operator wants their pincode to be.
`r`n`t 3. Then click the 'Reset Pin' button. Please wait 5 min for the new pin to sync with our system
`r`n`t`t (Notes: Our keyard system can be a bit finicky. Please avoid using pins that start with '1' or '0') `r`n "

$HelpTab3Text = "`r`n Tab3: 
`n`t To Create a New Employee in Active Directory: 
`r`n`t 1. Sign in with your admin account
`r`n`t 2. Enter the information for the New Employee in the respective boxes
`r`n`t 3. Double Check to make sure all information is accurate, including location, pin, training instructor, and Key Card #.
`r`n`t 4. Click the 'Add Employee' button
`r`n`t 5. Double Check in Active Directory that all information is accurate.
`r`n`t`t (Notes: The error checking on this program is currently not fully fleshed out yet, so some errors that pop up may not 
`r`n`t`t be accurate as to the error. For instance, previously usernames were automatically generated from the First and Last name fields,  
`r`n`t`t and if anyone had the same username as a new user, the error would simply state that the user already exists. If such an error
`r`n`t`t occurs, or any other error occurs with this tool when creating an Employee account, manually create the user in  
`r`n`t`t Active Directory and submit a ticket to Daniel Thomsen, Avantguard I.T.)
`r`n "

$HelpTab4Text = "`r`n Tab4: 
`n`t To Promote an Employee: 
`r`n`t 1. Sign in with your admin account
`r`n`t 2. Scroll down to the name of the Operator, Team Mentor, or Team Coach and select their name.
`r`n`t 3. Select all information as accurate to their position.
`r`n`t 4. Then click the 'Promote Employee' button.
`r`n`t`t (Notes: Leaving any drop down box may result in an error that you may or may not be notified of, including, 
`r`n`t`t but not limited to: No Promotion occuring, Innacurate information in Active Directory, or incomplete 
`r`n`t`t information in their accounts)
`r`n "

$HelpTab5Text = "`r`n Tab5:
`n`t To Terminate an Employee
`r`n`t 1. Sign in with your admin account
`r`n`t 2. Scroll down to the name of the Employee
`r`n`t 3. Select the 'Mark for Termination/Deletion' button. This will disable the account and set them as ready for termination
`r`n`t 4. This Process can be undone, but requires accessing the 'me.csv' file and deleting the name manually from there.
`r`n`t`t (Notes: This process will only work with this tool. Disabling the account outside of this tool won't auto delete the account)
`r`n
`r`n "

$HelpTabAppendixText = "`r`n Other Notes:
`r`n`t Please reach out to I.T. by submitting a ticket if you are having any issues at all. If you experience 
`r`n`t any issues with this tool specifically, please direct tickets to Daniel Thomsen, Avantguard I.T."

if ($newTempUserName.substring(0,5) -eq "admin" -Or $vTitle -eq "HR Generalist" -Or $vTitle -eq "Corporate Recruiter" -Or $vTitle -eq "Campus Recruiter" -Or $vTitle -eq "Operations Manager"){
$objTextBox6.Text = $HelpTabOpeningText + $HelpTab1Text + $HelpTab2Text + $HelpTab3Text + $HelpTab4Text + $HelpTab5Text + $HelpTabAppendixText
	}

elseif($vTitle -eq "$tempuserlocation Team Coach" -Or $vTitle -eq "$tempuserlocation Shift Coordinator" -Or $vTitle -eq "Grave Shift Supervisor" -Or $vTitle -eq "Graves Shift Supervisor" -Or $vTitle -eq "Operations Manager" -Or $vTitle -eq "Swings Shift Supervisor" -Or $vTitle -eq "WFM Real Time Analyst"){
	$objTextBox6.Text =  $HelpTabOpeningText + $HelpTab1Text + $HelpTab2Text + $HelpTabAppendixText
}

else{
    $objTextBox6.Text = "Call IT"
}
$Tab6.Controls.Add($objTextBox6)

$VersionLabel = New-Object System.Windows.Forms.label
$VersionLabel.Location = New-Object System.Drawing.Size(30,400)
$VersionLabel.Size = New-Object System.Drawing.Size(500,20)
$VersionLabel.BackColor = "TRANSPARENT"
$VersionLabel.ForeColor = "BLACK"
$VersionLabel.Text = "Ver. 3.1 - 10/17/2022 - Beta - By Daniel Thomsen - ©Avantguard | ©Becklar"
$VersionLabel.Font = ‘Microsoft Sans Serif,6’
$Tab6.Controls.Add($VersionLabel)
