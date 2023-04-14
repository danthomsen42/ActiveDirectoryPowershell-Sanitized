#########################################################################################################
########################################-----Bad Tab-----################################################
#########################################################################################################

####################-----Objective-----##################################################################
#
#  A Tab that only appears to tell you you don't have permission to use this tool
#
#########################################################################################################

$BadTab = New-object System.Windows.Forms.Tabpage
$BadTab.DataBindings.DefaultDataSourceUpdateMode = 0 
$BadTab.UseVisualStyleBackColor = $True 
$BadTab.Name = "Bad Tab" 
$BadTab.Text = "Insufficient Permissions” 

$BadTabWarningLabel = New-Object System.Windows.Forms.label
$BadTabWarningLabel.Location = New-Object System.Drawing.Size(30,30)
$BadTabWarningLabel.Size = New-Object System.Drawing.Size(500,500)
$BadTabWarningLabel.BackColor = "TRANSPARENT" 
$BadTabWarningLabel.ForeColor = "BLACK"
$BadTabWarningLabel.Text = " You do not have sufficient permissions to use this tool. Please close out and contact your local IT member. 
`n If you believe you should have permissions, make sure you are logged in to the correct account and have RSAT installed."
$BadTabWarningLabel.Font = ‘Microsoft Sans Serif,20’
$BadTab.Controls.Add($BadTabWarningLabel)