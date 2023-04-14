#########################################################################################################
####################################-----Tab 4 Contents-----#############################################
#########################################################################################################

#####-----For HR Use-----#####

####################-----Objective-----##################################################################
#
#  Allows HR To Promote or change an employee's job title within the monitoring center.
#
#########################################################################################################

#########################################################################################################
######################################-----Tab 4 Code-----###############################################
#########################################################################################################

##Promote Function - used to change job roles within the company while doing numerous checks to verify
##the jobtitle in relation to the shift and location
function Promote {
    $displayname = $UserList4.Text
    $Username = (Get-ADUser -Filter {DisplayName -eq $displayname}).samaccountname
    $DN = (get-aduser $Username).distinguishedname
    $BK = $false ##Is this a Becklar Job title
    $string = $PromoteLocationDropDown.Text
    $string2 = "Undefined"
    $otherManager = Get-ADUser -Filter "DisplayName -eq '$($ManagerList1.Text)'" | Select SamAccountName
    Set-aduser $Username -Manager $otherManager
    Switch ($string){
            'Rexburg'{
                $string2 = "REX"
            }
            'Ogden'{
                $string2 = "OGD"
            }
            'Cedar'{
                $string2 = "CED"
            }
            'Sarasota'{
                $string2 = "SAR"
            }
        }
        
    (get-aduser $Username -properties memberof).memberof|remove-adgroupmember -member $Username -confirm:$false




    Switch ($JobTitleDropDown.text){
      'CS Team Operator'{
        #(get-aduser $Username -properties memberof).memberof|remove-adgroupmember -member $Username -confirm:$false
        #Set-aduser $Username -Title "CS Team Operator"
        Set-aduser $Username -Department "Monitoring Center-$string"
        Move-ADObject -identity $DN -TargetPath "OU=Operators,OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("$string Team Operator", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
        
      }

      'Team Mentor'{
        Set-aduser $Username -Department "Monitoring Center-$string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Operators,OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("$string Team Mentor", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
      'Lead Team Mentor'{
        Set-aduser $Username -Department "Monitoring Center-$string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("$string Team Mentor", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }

      'Team Coach'{
        Set-aduser $Username -Department "Monitoring Center-$string"
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("$string Team Coach", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
       Switch ($string){
           'Ogden'{Add-ADGroupMember -Identity "DL_OgdTeamCoach" -Members $Username -confirm:$false}
           'Rexburg'{Add-ADGroupMember -Identity "DL_RexTeamCoach" -Members $Username -confirm:$false}
           'Cedar'{Add-ADGroupMember -Identity "DL_CEDShiftCoordinator" -Members $Username -confirm:$false}

      }  
    }
        'Dealer Care Representative'{
        Set-aduser $Username -Department "Monitoring Center-$string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("Dealer Care Representative", "AZ_365_BizLic", "Genetec_Cardholder_Ops", "Spark_Operators", "$string Team Operator")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
        'Training Instructor'{
        Set-aduser $Username -Department "Training $string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("DL_$string", "AZ_365_BizLic", "Genetec_Cardholder_Ops", "Spark_Operators", "Trainer", "DL_Training", "$string2 Roaming Users", "$string2 Users")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }

        'Accountant'{ 
        Set-aduser $Username -Department "Accounting"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("Accountant", "AZ_365_BizLic", "Genetec_Cardholder_Ops", "DL_AGCharity")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }

        'Video Specialist'{ 
        Set-aduser $Username -Department "Video Monitoring Team"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Operators,OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("Video Specialist", "AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
        'Workforce Management Manager'{ 
        Set-aduser $Username -Department "Monitoring Center $string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("WFM Manager", "AZ_365_BizLic", "Genetec_Cardholder_Ops", "Spark_Operators", "DL_CSOgden", "DL_CSRexburg", "DL_Rexburg", "DL_5k")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
        'WFM Scheduler'{ 
        Set-aduser $Username -Department "Monitoring Center $string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("Scheduler", "AZ_365_BizLic", "Genetec_Cardholder_Ops", "Spark_Operators", "DL_CSOgden", "DL_CSRexburg", "DL_Rexburg", "DL_5k", "DL_Ogden", "GPO_DefaultPrinter_Accounting", "UAT_Users", "DL_AGCaresDirector", "OGD Users")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
        'WFM Forecaster'{ 
        Set-aduser $Username -Department "Monitoring Center $string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("AZ_365_BizLic", "Genetec_Cardholder_Ops", "Spark_Operators", "OGD Users")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
      }
        'WFM Real Time Analyst'{ 
        Set-aduser $Username -Department "Monitoring Center-$string"
        #Set-aduser $Username -Manager 
        Move-ADObject -identity $DN -TargetPath "OU=Users,OU=Profiles,DC=ag-ul,DC=local"
        $groups = @("AZ_365_F1Lic", "Genetec_Cardholder_Ops", "Spark_Operators", "$string Shift Coordinator")
        foreach ($group in $groups){
            Get-ADGroup -Identity $group | Add-ADGroupMember -Members $Username
        }
             Switch ($string){
           'Ogden'{Add-ADGroupMember -Identity "DL_OgdShiftCoordinator" -Members $Username -confirm:$false}
           'Rexburg'{Add-ADGroupMember -Identity "DL_RexShiftCoordinator" -Members $Username -confirm:$false}
           'Cedar'{Add-ADGroupMember -Identity "DL_CEDShiftCoordinator" -Members $Username -confirm:$false}

      }  
      }
    }##End of Switch Statement

    if ($TimeList1.text -eq "Part Time")
    {
        Remove-ADGroupMember -Identity "DL_BE_FullTime" -Members $Username -confirm:$false
        Add-ADGroupMember -Identity "DL_BE_PartTime" -Members $Username -confirm:$false
    }
    else{
        Remove-ADGroupMember -Identity "DL_BE_PartTime" -Members $Username -confirm:$false
        Add-ADGroupMember -Identity "DL_BE_FullTime" -Members $Username -confirm:$false
    }


    set-aduser $Username -replace @{physicalDeliveryOfficeName = $string2}
    Set-aduser $Username -Title $JobTitleDropDown.text

    #if ($PromoteLocationDropDown.Text -eq "Cedar"){
    #    Add-ADGroupMember -Identity "DL_CSCedar" -Members $Username -confirm:$false
    #}

    [System.Windows.Forms.MessageBox]::Show('Employee Role Changed')

}

#########################################################################################################
#####################################-----Tab 4 Display-----#############################################
#########################################################################################################

$Tab4 = New-object System.Windows.Forms.Tabpage
$Tab4.DataBindings.DefaultDataSourceUpdateMode = 0 
$Tab4.UseVisualStyleBackColor = $True 
$Tab4.Name = "Tab4" 
$Tab4.Text = "Promote Employee” 

$InstLabel4 = New-Object System.Windows.Forms.label
$InstLabel4.Location = New-Object System.Drawing.Size(10,30)
$InstLabel4.Size = New-Object System.Drawing.Size(500,20)
$InstLabel4.BackColor = "TRANSPARENT"
$InstLabel4.ForeColor = "BLACK"
$InstLabel4.Text = "Currently in Beta - Please Double Check for accuracy"
$Tab4.Controls.Add($InstLabel4)


$PromoteLocationLabel = New-Object System.Windows.Forms.label
$PromoteLocationLabel.Location = New-Object System.Drawing.Size(10,130)
$PromoteLocationLabel.Size = New-Object System.Drawing.Size(100,20)
$PromoteLocationLabel.BackColor = "TRANSPARENT"
$PromoteLocationLabel.ForeColor = "BLACK"
$PromoteLocationLabel.Text = "Location"
$Tab4.Controls.Add($PromoteLocationLabel)

$PromoteLocationDropDown = New-Object system.Windows.Forms.ComboBox
$PromoteLocationDropDown.text = “Location”
$PromoteLocationDropDown.location = New-Object System.Drawing.Point(130,130)
$PromoteLocationDropDown.Size = New-Object System.Drawing.Size(200,50)


$PromoteLocationDropDown.Items.Add('Ogden')
$PromoteLocationDropDown.Items.Add('Rexburg')
$PromoteLocationDropDown.Items.Add('Cedar')
$PromoteLocationDropDown.Items.Add('Sarasota')
$PromoteLocationDropDown.Items.Add('Armstrong')

$Tab4.Controls.Add($PromoteLocationDropDown)


$JobTitleLabel = New-Object System.Windows.Forms.label
$JobTitleLabel.Location = New-Object System.Drawing.Size(10,160)
$JobTitleLabel.Size = New-Object System.Drawing.Size(100,20)
$JobTitleLabel.BackColor = "TRANSPARENT"
$JobTitleLabel.ForeColor = "BLACK"
$JobTitleLabel.Text = "New Job Title"
$Tab4.Controls.Add($JobTitleLabel)

$JobTitleDropDown = New-Object system.Windows.Forms.ComboBox
$JobTitleDropDown.text = “Job Title”
$JobTitleDropDown.location = New-Object System.Drawing.Point(130,160)
$JobTitleDropDown.Size = New-Object System.Drawing.Size(200,50)

####Add Job Titles to Drop Down -- Make sure they match the Switch Statement above
$JobTitleDropDown.Items.Add('CS Team Operator')
$JobTitleDropDown.Items.Add('Team Mentor')
$JobTitleDropDown.Items.Add('Lead Team Mentor')
$JobTitleDropDown.Items.Add('Team Coach')
$JobTitleDropDown.Items.Add('Dealer Care Representative')
$JobTitleDropDown.Items.Add('Training Instructor')
$JobTitleDropDown.Items.Add('Video Specialist')
$JobTitleDropDown.Items.Add('Workforce Management Manager')
$JobTitleDropDown.Items.Add('WFM Scheduler')
$JobTitleDropDown.Items.Add('WFM Forecaster')
$JobTitleDropDown.Items.Add('Accountant')


$Tab4.Controls.Add($JobTitleDropDown)

$ShiftLabel = New-Object System.Windows.Forms.label
$ShiftLabel.Location = New-Object System.Drawing.Size(10,190)
$ShiftLabel.Size = New-Object System.Drawing.Size(100,20)
$ShiftLabel.BackColor = "TRANSPARENT"
$ShiftLabel.ForeColor = "BLACK"
$ShiftLabel.Text = "Shift"


$ShiftDropDown = New-Object system.Windows.Forms.ComboBox
$ShiftDropDown.text = “”
$ShiftDropDown.location = New-Object System.Drawing.Point(130,190)
$ShiftDropDown.Size = New-Object System.Drawing.Size(200,50)


$ShiftDropDown.Items.Add('Days')
$ShiftDropDown.Items.Add('Swings')
$ShiftDropDown.Items.Add('Graves')


$ManagerLabel = New-Object System.Windows.Forms.label
$ManagerLabel.Location = New-Object System.Drawing.Size(10,280)
$ManagerLabel.Size = New-Object System.Drawing.Size(100,20)
$ManagerLabel.BackColor = "TRANSPARENT"
$ManagerLabel.ForeColor = "BLACK"
$ManagerLabel.Text = "New Manager"
$Tab4.Controls.Add($ManagerLabel)

$ManagerList1 = New-Object system.Windows.Forms.ComboBox
$ManagerList1.text = “Managers”
$ManagerList1.width = 170
$ManagerList1.autosize = $true
$ManagerList1.location = New-Object System.Drawing.Point(130,280)
$ManagerList1.Font = ‘Microsoft Sans Serif,10’


###Add Manager titles to List
$Managers1 = Get-ADUser -Filter "(Title -like '*Rexburg Team Coach*')-Or (Title -like '*Ogden Team Coach*')-Or
(Title -like '*Cedar Team Coach*')-Or (Title -like '*Operations Manager*')-Or (Title -like '*Grave Shift Supervisor*')-Or
(Title -like '*Graves Shift Supervisor*')-Or (Title -like '*Swing Shift Supervisor*')-Or (Title -like '*Swings Shift Supervisor*')-Or
(Title -like '*Rexburg Shift Coordinator*')-Or (Title -like '*Ogden Shift Coordinator*')-Or (Title -like '*Cedar Shift Coordinator*')-Or
(Title -like '*Customer Success Manager*')-Or (Title -like '*Assistant Controller*')-Or (Title -like '*Manager of Video Monitoring*')"| Sort-Object | Select-Object name
$Managers1 | ForEach-Object {
    $ManagerList1.Items.Add($_.name)
}

$Tab4.Controls.Add($ManagerList1)

$TimeLabel = New-Object System.Windows.Forms.label
$TimeLabel.Location = New-Object System.Drawing.Size(10,310)
$TimeLabel.Size = New-Object System.Drawing.Size(100,20)
$TimeLabel.BackColor = "TRANSPARENT"
$TimeLabel.ForeColor = "BLACK"
$TimeLabel.Text = "Full or Part Time"
$Tab4.Controls.Add($TimeLabel)

$TimeList1 = New-Object system.Windows.Forms.ComboBox
$TimeList1.text = “Time”
$TimeList1.width = 170
$TimeList1.autosize = $true
$TimeList1.location = New-Object System.Drawing.Point(130,310)
$TimeList1.Font = ‘Microsoft Sans Serif,10’
$TimeList1.Items.Add('Part Time')
$TimeList1.Items.Add('Full Time')
$Tab4.Controls.Add($TimeList1)




$button4 = New-Object System.Windows.Forms.Button
$button4.BackColor = "#d0caca"
$button4.Location = New-Object System.Drawing.Point(400,130)
$button4.Size = New-Object System.Drawing.Size (200,30)
$button4.Text = "Promote Employee"
$button4.Add_Click({Promote})
$Tab4.Controls.Add($button4)

$Tab4.Controls.Add($UserList4)




