#########################################################################################################
####################################-----Tab 3 Contents-----#############################################
#########################################################################################################

#####-----For HR Use-----#####

####################-----Objective-----##################################################################
#
#  Allows HR to add a new employee into the system without the need for IT to get involved and without 
#  accessing Active Directory directly
#
#########################################################################################################

#########################################################################################################
######################################-----Tab 3 Code-----###############################################
#########################################################################################################

$Tab3 = New-object System.Windows.Forms.Tabpage
$Tab3.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab3.UseVisualStyleBackColor = $True 
$Tab3.Name = "Tab3" 
$Tab3.Text = "New Employee” 

function AddUser {
    $userPath = "OU=Operators,OU=Users,OU=Profiles,DC=ag-ul,DC=local"
    [SecureString] $defaultPass = ConvertTo-SecureString -AsPlainText "Avantguard1" -Force

    $Username = $FN.Text.substring(0,1).toLower() + $LN.Text.toLower()#################################################CONTINUE HERE####################################
    ###Check Username here###
    $i = 1;
   while (Get-aduser $Username){
       $Username = $FN.Text.substring(0,$i).toLower() + $MiddleInitial.Text.toLower() + $LN.Text.toLower()
        ++$i
    #Get-aduser $Username
   }
    $email = $Username+"@agmonitoring.com"
    $upn = $Username + "@agmonitoring.com"
    $MI = $MiddleInitial.Text
    
    Write-Output "KCIn: " + $KCIn.Text
    
    $TextInfo = (Get-Culture).TextInfo
    if($MI.length -gt 0){
	    $MI = $MiddleInitial.Text.substring(0,1).toUpper()
	    $name = $TextInfo.ToTitleCase($FN.Text + " " + $MiddleInitial.Text + ". " + $LN.Text)
    }
    else {$name = $TextInfo.ToTitleCase($FN.Text + " " + $LN.Text)}
    
    Switch ($LocationDropDown.Text) {
    'Ogden' {
            $office = "OGD"
            $street = "4699 Harrison Blvd. Suite 100"
            $city = "Ogden"
            $state = "UT"
            $zip = "84430"
            $country = "US"
            $fax = "(801) 399-4803"
			$phone = "(801) 781-6100"
			$company = "Avantguard Monitoring Centers"
			$mgr = Get-ADUser -Filter "DisplayName -eq '$($TrainerList1.Text)'" | Select SamAccountName
			$groups = @("Ogden Team Operator", "GPO_DefaultPrinter_OGDCentral", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "DL_BE_PartTime", "Spark_Operators")
            $department = "Monitoring Center-Ogden"
            $title = "CS Team Operator"
            }
    'Rexburg' {
            $office = "REX" 
			$street = "366 Grand Loop"
			$city = "Rexburg"
			$zip = "83440"
			$state = "ID"
			$country = "US"
			$fax = "(801) 781-6133"
			$phone = "(801) 781-6100"
			$company = "Avantguard Monitoring Centers"
            $mgr = Get-ADUser -Filter "DisplayName -eq '$($TrainerList1.Text)'" | Select SamAccountName
			$groups = @("Rexburg Team Operator", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "DL_BE_PartTime", "Spark_Operators")
            $department = "Monitoring Center-Rexburg"
            $title = "CS Team Operator"
            }
    'Cedar' {
            $office = "CED" 
			$street = "888 Sage Drive"
			$city = "Cedar City"
			$zip = "84720"
			$state = "UT"
			$country = "US"
			$fax = "(801) 399-4803"
			$phone = "(801) 781-6100"
			$company = "Avantguard Monitoring Centers"
			$mgr = Get-ADUser -Filter "DisplayName -eq '$($TrainerList1.Text)'" | Select SamAccountName
			$groups = @("Cedar Team Operator", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "DL_BE_PartTime", "Spark_Operators")
            $department = "Monitoring Center-Cedar City"
            $title = "CS Team Operator"
            }
    'Sarasota' {
            $office = "SAR" ##
			$street = "888 Sage Drive" ##
			$city = "Cedar City"##
			$zip = "84720"##
			$state = "FL"
			$country = "US"
			$fax = "(801) 399-4803"##
			$phone = "(801) 781-6100"##
			$company = "Avantguard Monitoring Centers" ##
			$mgr = Get-ADUser -Filter "DisplayName -eq '$($TrainerList1.Text)'" | Select SamAccountName
			$groups = @("Sarasota Team Operator", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "DL_BE_PartTime", "Spark_Operators") ###Job Title Group
            $department = "Monitoring Center-Sarasota"
            $title = "CS Team Operator"
            }
    'Armstrong' {
            $office = "ARM" ##
			$street = "888 Sage Drive" ##
			$city = "Cedar City"##
			$zip = "84720"##
			$state = "BC"
			$country = "CA"
			$fax = "(801) 399-4803"##
			$phone = "(801) 781-6100"##
			$company = "Armstrong" ##
			$mgr = Get-ADUser -Filter "DisplayName -eq '$($TrainerList1.Text)'" | Select SamAccountName
			$groups = @("Domain Users", "AZ_365_F1Lic") ###Job Title Group
            $department = "Monitoring Center-Sarasota"
            $title = "CS Team Operator"
            }
        }
    ### Update this in case of future new Facility codes.
    ### Notes for Future - change to switch/case instead of If/else
    if ($KCIn.Text -eq $null){
        #$KCIn.Text = ""
        #$FacilityCode = "";
    }
    elseif ($KCIn.text -eq ""){
       # $KCIn.Text = ""
       # $FacilityCode = "";
    }
    elseif ($KCIn.text.substring(0,2) -eq "09"){
        $FacilityCode = 15;
    }
    elseif ($KCIn.text.substring(0,1) -eq "9"){
      $FacilityCode = 15;
    }
    elseif ($KCIn.text.substring(0,2) -eq "33"){
      $FacilityCode = 140;
    }
    elseif ($KCIn.text.substring(0,3) -eq "104"){
      $FacilityCode = 15;
    }
    elseif ($KCIn.text.substring(0,3) -eq "101"){
      $FacilityCode = 26;
    }
    elseif ($KCIn.text.substring(0,3) -eq "102"){
      $FacilityCode = 26;
    }
    else{
        $FacilityCode = 0;
    }
    
try {    New-ADUser 	-Name 						$name 				`
                        -DisplayName				$name				`
                        -SamAccountName 		 	$username 			`
                        -UserPrincipalName			$upn				`
                        -Givenname				 	$FN.Text		`
                        -Surname				 	$LN.Text		`
                        -Initials				 	$MI        		`
                        -Path					 	$userPath			`
                        -AccountPassword		 	$defaultPass		`
                        -Office					 	$office				`
                        -City					 	$city				`
                        -Company				 	$company			`
                        -Country				 	$country			`
                        -Department				 	$department			`
                        -EmailAddress			 	$email				`
                        -Fax						$fax				`
                        -Manager				 	$mgr			    `
                        -OfficePhone			 	$phone				`
                        -PostalCode 			 	$zip				`
                        -State					 	$state				`
                        -StreetAddress				$street				`
                        -Title					 	$title				`
                        -Enabled				 	$True				`
                        -ChangePasswordAtLogon	 	$True				`
                        -PassThru
    
    if ($KCIn.Text -eq $null -or $KCIn.Text -eq ""){
        #Do nothing
    }
    else{
        set-aduser $Username -add @{otherIpPhone = $FacilityCode}
        set-aduser $Username -add @{otherHomePhone = $KCIn.Text}
    }
 if ($PinIn.Text -eq $null -or $PinIn.Text -eq ""){
     #Do nothing
 }
    else{
        set-aduser $Username -add @{otherFacsimileTelephoneNumber = $PinIn.Text}
    }

    set-aduser $Username -add @{proxyAddresses = "SMTP:"+$email}
        foreach ($group in $groups)
        {
        Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
        [System.Windows.Forms.MessageBox]::Show('User account ' + $Username + ' has been created')
        $objTextBox3.Text += "$Username - $name `r`n"
        $completed = $True
    }
catch #[Microsoft.ActiveDirectory.Management.ADIdentityAlreadyExistsException]
{ 
    [System.Windows.Forms.MessageBox]::Show('User already exists')
    New-ADUser 	-Name 						$name 				`
                -DisplayName				$name				`
                -SamAccountName 		 	$username 			`
                -UserPrincipalName			$upn				`
                -Givenname				 	$FN.Text		`
                -Surname				 	$LN.Text		`
                -Initials				 	$MI         		`
                -Path					 	$userPath			`
                -AccountPassword		 	$defaultPass		`
                -Office					 	$office				`
                -City					 	$city				`
                -Company				 	$company			`
                -Country				 	$country			`
                -Department				 	$department			`
                -EmailAddress			 	$email				`
                -Fax						$fax				`
                -Manager				 	$mgr			    `
                -OfficePhone			 	$phone				`
                -PostalCode 			 	$zip				`
                -State					 	$state				`
                -StreetAddress				$street				`
                -Title					 	$title				`
                -Enabled				 	$True				`
                -ChangePasswordAtLogon	 	$True				`
                 -PassThru

    set-aduser $Username -add @{otherIpPhone = $FacilityCode}
    set-aduser $Username -add @{otherHomePhone = $TextBoxCardNumber.Text}
    set-aduser $Username -add @{otherFacsimileTelephoneNumber = $TextBoxPIN.Text}
    set-aduser $Username -add @{proxyAddresses = "SMTP:"+$email}
        foreach ($group in $groups)
        {
        Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
        [System.Windows.Forms.MessageBox]::Show('User account has been created2')
        $completed = $True
    }
}


####################-----Details-----####################################################################
#  TODO Maybe:
#   1) Take out trainer and have it automatically decided based on Location.
#  TODO Definately:
#   1) Make Sure there aren't duplicate anything (names aside) 
#       a)No Duplicate Keycards #'s        
#   2) Add Department based on location to Organization Tab
#########################################################################################################



$InstLabel3 = New-Object System.Windows.Forms.label
$InstLabel3.Location = New-Object System.Drawing.Size(10,30)
$InstLabel3.Size = New-Object System.Drawing.Size(500,20)
$InstLabel3.BackColor = "TRANSPARENT"
$InstLabel3.ForeColor = "BLACK"
$InstLabel3.Text = "Currently in Beta - Please Double Check for accuracy"
$Tab3.Controls.Add($InstLabel3)

$FNLabel = New-Object System.Windows.Forms.label
$FNLabel.Location = New-Object System.Drawing.Size(10,70)
$FNLabel.Size = New-Object System.Drawing.Size(100,20)
$FNLabel.BackColor = "TRANSPARENT"
$FNLabel.ForeColor = "BLACK"
$FNLabel.Text = "First Name"
$Tab3.Controls.Add($FNLabel)

$FN = New-Object System.Windows.Forms.TextBox
$FN.Text = ""
$FN.Location = New-Object System.Drawing.Point(130,70) ### Location of the text box
$FN.Size = New-Object System.Drawing.Size(200,50) ### Size of the text box
$FN.Multiline = $false ### Allows multiple lines of data
$FN.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$FN.ReadOnly=$false
$Tab3.Controls.Add($FN)

$MILabel = New-Object System.Windows.Forms.label
$MILabel.Location = New-Object System.Drawing.Size(10,100)
$MILabel.Size = New-Object System.Drawing.Size(100,20)
$MILabel.BackColor = "TRANSPARENT"
$MILabel.ForeColor = "BLACK"
$MILabel.Text = "Middle Initial"
$Tab3.Controls.Add($MILabel)

$MiddleInitial = New-Object System.Windows.Forms.TextBox
$MiddleInitial.Text = ""
$MiddleInitial.Location = New-Object System.Drawing.Point(130,100) ### Location of the text box
$MiddleInitial.Size = New-Object System.Drawing.Size(50,50) ### Size of the text box
$MiddleInitial.Multiline = $false ### Allows multiple lines of data
$MiddleInitial.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$MiddleInitial.ReadOnly=$false
$MiddleInitial.MaxLength = 1
$Tab3.Controls.Add($MiddleInitial)

$LNLabel = New-Object System.Windows.Forms.label
$LNLabel.Location = New-Object System.Drawing.Size(10,130)
$LNLabel.Size = New-Object System.Drawing.Size(100,20)
$LNLabel.BackColor = "TRANSPARENT"
$LNLabel.ForeColor = "BLACK"
$LNLabel.Text = "Last Name"
$Tab3.Controls.Add($LNLabel)

$LN = New-Object System.Windows.Forms.TextBox
$LN.Text = ""
$LN.Location = New-Object System.Drawing.Point(130,130) ### Location of the text box
$LN.Size = New-Object System.Drawing.Size(200,50) ### Size of the text box
$LN.Multiline = $false ### Allows multiple lines of data
$LN.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$LN.ReadOnly=$false
$Tab3.Controls.Add($LN)

$LocationLabel = New-Object System.Windows.Forms.label
$LocationLabel.Location = New-Object System.Drawing.Size(10,160)
$LocationLabel.Size = New-Object System.Drawing.Size(100,20)
$LocationLabel.BackColor = "TRANSPARENT"
$LocationLabel.ForeColor = "BLACK"
$LocationLabel.Text = "Location"
$Tab3.Controls.Add($LocationLabel)

$LocationDropDown = New-Object system.Windows.Forms.ComboBox
$LocationDropDown.text = “Location”
$LocationDropDown.location = New-Object System.Drawing.Point(130,160)
$LocationDropDown.Size = New-Object System.Drawing.Size(200,50)

$LocationDropDown.Items.Add('Ogden')
$LocationDropDown.Items.Add('Rexburg')
$LocationDropDown.Items.Add('Cedar')
$LocationDropDown.Items.Add('Sarasota')
$LocationDropDown.Items.Add('Armstrong')

$Tab3.Controls.Add($LocationDropDown)

$PinLabel = New-Object System.Windows.Forms.label
$PinLabel.Location = New-Object System.Drawing.Size(10,190)
$PinLabel.Size = New-Object System.Drawing.Size(100,20)
$PinLabel.BackColor = "TRANSPARENT"
$PinLabel.ForeColor = "BLACK"
$PinLabel.Text = "PIN"
$Tab3.Controls.Add($PinLabel)

$PinIn = New-Object System.Windows.Forms.TextBox
$PinIn.Text = ""
$PinIn.Location = New-Object System.Drawing.Point(130,190) ### Location of the text box
$PinIn.Size = New-Object System.Drawing.Size(80,50) ### Size of the text box
$PinIn.Multiline = $false ### Allows multiple lines of data
$PinIn.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$PinIn.ReadOnly=$false
$PinIn.MaxLength = 4
$Tab3.Controls.Add($PinIn)

$KeyCardLabel = New-Object System.Windows.Forms.label
$KeyCardLabel.Location = New-Object System.Drawing.Size(10,220)
$KeyCardLabel.Size = New-Object System.Drawing.Size(100,20)
$KeyCardLabel.BackColor = "TRANSPARENT"
$KeyCardLabel.ForeColor = "BLACK"
$KeyCardLabel.Text = "Key Card #"
$Tab3.Controls.Add($KeyCardLabel)

$KCIn = New-Object System.Windows.Forms.TextBox
$KCIn.Text = ""
$KCIn.Location = New-Object System.Drawing.Point(130,220) ### Location of the text box
$KCIn.Size = New-Object System.Drawing.Size(100,50) ### Size of the text box
$KCIn.Multiline = $false ### Allows multiple lines of data
$KCIn.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$KCIn.ReadOnly=$false
$KCIn.MaxLength = 5
$Tab3.Controls.Add($KCIn)

$EmailLabel = New-Object System.Windows.Forms.label
$EmailLabel.Location = New-Object System.Drawing.Size(10,250)
$EmailLabel.Size = New-Object System.Drawing.Size(100,20)
$EmailLabel.BackColor = "TRANSPARENT"
$EmailLabel.ForeColor = "BLACK"
$EmailLabel.Text = "Email"

$EmailIn = New-Object System.Windows.Forms.TextBox
$EmailIn.Text = ""
$EmailIn.Location = New-Object System.Drawing.Point(130,250) ### Location of the text box
$EmailIn.Size = New-Object System.Drawing.Size(200,50) ### Size of the text box
$EmailIn.Multiline = $false ### Allows multiple lines of data
$EmailIn.Font = New-Object System.Drawing.Font("Consolas",10,[System.Drawing.FontStyle]::Regular)
$EmailIn.ReadOnly=$false

$TrainerLabel = New-Object System.Windows.Forms.label
$TrainerLabel.Location = New-Object System.Drawing.Size(10,280)
$TrainerLabel.Size = New-Object System.Drawing.Size(100,20)
$TrainerLabel.BackColor = "TRANSPARENT"
$TrainerLabel.ForeColor = "BLACK"
$TrainerLabel.Text = "Training Instructor"
$Tab3.Controls.Add($TrainerLabel)

$TrainerList1 = New-Object system.Windows.Forms.ComboBox
$TrainerList1.text = “Lead Trainers”
$TrainerList1.width = 170
$TrainerList1.autosize = $true
$TrainerList1.location = New-Object System.Drawing.Point(130,280)
$TrainerList1.Font = ‘Microsoft Sans Serif,10’

$Trainers = Get-ADUser -Filter "(Title -like '*Lead Training Instructor*')-Or (Title -like '*Training Instructor*')-Or (Title -like '*Training and Quality Manager*')-Or (Title -like '*Training Intern*')-Or (Title -like '*Operations Manager*')"| Sort-Object | Select-Object name
$Trainers | ForEach-Object {
    $TrainerList1.Items.Add($_.name)
}


$Tab3.Controls.Add($TrainerList1)

$JobTitleDropDown = New-Object system.Windows.Forms.ComboBox
$JobTitleDropDown.text = “Job Title”
$JobTitleDropDown.location = New-Object System.Drawing.Point(130,310)
$JobTitleDropDown.Size = New-Object System.Drawing.Size(200,50)

$JobTitleDropDown.Items.Add('CS Team Operator')
$JobTitleDropDown.Items.Add('Other')

$Tab3.Controls.Add($JobTitleDropDown)

$button3 = New-Object System.Windows.Forms.Button
$button3.BackColor = "#d0caca"
$button3.Location = New-Object System.Drawing.Point(400,130)
$button3.Size = New-Object System.Drawing.Size (200,30)
$button3.Text = "Add Employee"
$button3.Add_Click({AddUser; $FN.Clear(), $MiddleInitial.Clear(), $LN.Clear(), $PinIn.Clear(), $KCIn.Clear(), $FN.Focus()})
$Tab3.Controls.Add($button3)

#$objTextBox3.Text += "$group `r`n"

$objTextBox3Label = New-Object System.Windows.Forms.label
$objTextBox3Label.Location = New-Object System.Drawing.Size(400,210)
$objTextBox3Label.Size = New-Object System.Drawing.Size(200,20)
$objTextBox3Label.BackColor = "TRANSPARENT"
$objTextBox3Label.ForeColor = "BLACK"
$objTextBox3Label.Text = "Users added this session:"
$Tab3.Controls.Add($objTextBox3Label)

$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Text = ""
$objTextBox3.Multiline = $True;
$objTextBox3.Location = New-Object System.Drawing.Size(400,230) 
$objTextBox3.Size = New-Object System.Drawing.Size(250,130)
$objTextBox3.Scrollbars = "Vertical" 
$objTextBox3.ReadOnly=$true
$objTextBox3.ForeColor = "BLACK"
$Tab3.Controls.Add($objTextBox3)

$ClearButton = New-Object System.Windows.Forms.Button;
$ClearButton.BackColor = "#d0caca"
$ClearButton.Location = New-Object System.Drawing.Size(400,310)
$ClearButton.Size = New-Object System.Drawing.Size(200,30)
$ClearButton.Text = "Clear"
$ClearButton.Add_Click{$FN.Clear(), $MiddleInitial.Clear(), $LN.Clear(), $PinIn.Clear(), $KCIn.Clear(), $FN.Focus()}
#$Tab3.Controls.Add($ClearButton)

$RefreshButton = New-Object System.Windows.Forms.Button;
$RefreshButton.BackColor = "#d0caca"
$RefreshButton.Location = New-Object System.Drawing.Size(130,400)
$RefreshButton.Size = New-Object System.Drawing.Size(200,30)
$RefreshButton.Text = "Refresh"
$RefreshButton.Add_Click{UserLists;}
$Tab3.Controls.Add($RefreshButton)
