#########################################################################################################
####################################-----Tab 5 Contents-----#############################################
#########################################################################################################

#####-----For HR Use-----#####

####################-----Objective-----##################################################################
#
#  Terminate Employees
#
#########################################################################################################

#
# Add delete keycard info to the disable step
#

#########################################################################################################
######################################-----Tab 5 Code-----###############################################
#########################################################################################################

$objTextBox5 = New-Object System.Windows.Forms.TextBox 
$objTextBox5.Multiline = $True;
$objTextBox5.Location = New-Object System.Drawing.Size(300,130) 
$objTextBox5.Size = New-Object System.Drawing.Size(400,200)
$objTextBox5.Scrollbars = "Vertical" 
$objTextBox5.ReadOnly=$true
$objTextBox5.ForeColor = "BLACK"


#$ExpiredUsers | ForEach-Object {
#   $objTextBox5.Text += $ExpiredUsers.SamAccountName
#}

##function markedforTermList{
   ## $MarkedForTermCSV = Import-Csv ".\me.csv"
   ## foreach($LINE in $MarkedForTermCSV)
   ## {
  ##      $objTextBox5.Text += "$($LINE.UserName), $($LINE.DisplayName), $($LINE.CurrentDate) `r`n"
  ##  }

##}

##markedforTermList


function disableUsers{

    $displayname = $UserList5.Text
    $Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
   # set-aduser -Identity $Username -replace @{otherFacsimileTelephoneNumber = $textBox2.Text}
    
    Set-ADUser -Identity $Username -Clear @('otherFacsimileTelephoneNumber','otherHomePhone','otherIpPhone')
    
    Disable-ADAccount -Identity $Username
   
    # Set the user's account to expire in one week
    $expirationDate = (Get-Date).AddDays(7)
    Set-ADUser -Identity $username -AccountExpirationDate $expirationDate

    $objTextBox5.Text += "$username, `r`n"

   }

#########################################################################################################
####################################-----Tab 5 display-----##############################################
#########################################################################################################
$Tab5 = New-object System.Windows.Forms.Tabpage
$Tab5.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab5.UseVisualStyleBackColor = $True 
$Tab5.Name = "Tab5" 
$Tab5.Text = "Terminate User” 

$InstLabel5 = New-Object System.Windows.Forms.label
$InstLabel5.Location = New-Object System.Drawing.Size(10,30)
$InstLabel5.Size = New-Object System.Drawing.Size(500,20)
$InstLabel5.BackColor = "TRANSPARENT"
$InstLabel5.ForeColor = "BLACK"
$InstLabel5.Text = "Currently In Beta - Please Use with Extreme Caution"
$Tab5.Controls.Add($InstLabel5)

$TermLabel = New-Object System.Windows.Forms.label
$TermLabel.Location = New-Object System.Drawing.Size(300,110)
$TermLabel.Size = New-Object System.Drawing.Size(500,20)
$TermLabel.BackColor = "TRANSPARENT"
$TermLabel.ForeColor = "BLACK"
$TermLabel.Text = "Disabled, awaiting termination"
$Tab5.Controls.Add($TermLabel)


$Tab5.Controls.Add($objTextBox5)

$button5 = New-Object System.Windows.Forms.Button
$button5.BackColor = "#d0caca"
$button5.Location = New-Object System.Drawing.Point(10,120)
$button5.Size = New-Object System.Drawing.Size (200,30)
$button5.Text = "Mark for Termination/Deletion"
$button5.Add_Click({ disableUsers })
$Tab5.Controls.Add($button5)

$refreshButton5 = New-Object System.Windows.Forms.Button
$refreshButton5.BackColor = "#d0caca"
$refreshButton5.Location = New-Object System.Drawing.Point(200,50)
$refreshButton5.Size = New-Object System.Drawing.Size (200,30)
$refreshButton5.Text = "Refresh List"
$refreshButton5.Add_Click({ UserLists })
##$Tab5.Controls.Add($refreshButton5)


$Tab5.Controls.Add($UserList5)