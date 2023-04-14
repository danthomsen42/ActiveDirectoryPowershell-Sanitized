#########################################################################################################
####################################-----Tab 2 Contents-----#############################################
#########################################################################################################

#####-----For SC Use-----#####

####################-----Objective-----##################################################################
#
#  Allows SC's to change an Employee's pin number if an employee forgets it
#
#########################################################################################################


#########################################################################################################
######################################-----Tab 2 Code-----###############################################
#########################################################################################################

##Pin Code Reset
function PinReset {
    $displayname = $UserList2.Text
    $Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
    set-aduser -Identity $Username -replace @{otherFacsimileTelephoneNumber = $textBox2.Text}
    [System.Windows.Forms.MessageBox]::Show('Pin has been Reset')
}


$Tab2 = New-object System.Windows.Forms.Tabpage
$Tab2.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab2.UseVisualStyleBackColor = $True 
$Tab2.Name = "Tab2" 
$Tab2.Text = "Key Card Pin Reset” 

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,70) ### Location of the text box
$textBox2.Size = New-Object System.Drawing.Size(200,50) ### Size of the text box
$textBox2.Multiline = $false ### Allows multiple lines of data
$textBox2.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$textBox2.ReadOnly=$false
$textBox2.MaxLength = 4
$Tab2.Controls.Add($textBox2)

$button2 = New-Object System.Windows.Forms.Button
$button2.BackColor = "#d0caca"
$button2.Location = New-Object System.Drawing.Point(10,120)
$button2.Size = New-Object System.Drawing.Size (200,30)
$button2.Text = "Reset Pin"
$Tab2.Controls.Add($button2)

$Tab2.Controls.Add($UserList2)
$button2.Add_Click({ PinReset })

