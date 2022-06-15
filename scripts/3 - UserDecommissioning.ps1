#Begin Functions
#Good 5/5/2021
#


Function SourceMailboxCheck ($Global:RMBX)
    {
    Write-host "Enter Source mailbox check with " $Global:RMBX
    
    If ($Global:RMBX -eq $null)
        {
        Write-Host "Could not get remote mailbox " $Global:SourceUser.distinguishedname "Exiting script" -ForegroundColor red -BackgroundColor Yellow
        Exit
        }

    If ($Global:RMBX -is [Array])
        {
        Write-Host "Multiple remote mailboxes returned for " $Global:SourceUser.distinguishedname "Exiting script" -ForegroundColor red -BackgroundColor Yellow
        Exit
        }
      
     write-host "Single remote mailbox returned - " $Global:RMBX     
     Write-host "Exiting Source mailbox check function" 
          
     }

Function ConvertToShare ($Global:ConvertToShareObject)
    {
    Write-Host "Enter convert to Shared with " $Global:ConvertToShareObject

    Set-Remotemailbox -Identity $Global:ConvertToShareObject -Type Shared

    $DescriptionSTR = $Global:SourceUser.Description + " Disabled on " + $Global:currentdate
    Write-Host "New description will be " $DescriptionSTR
    Set-ADuser -Identity $Global:ConvertToShareObject -Description $DescriptionSTR  -server $DC -credential $Global:UserCredential # -HomeDirectory $null remarked 2/21 to accomidate not everyone having OD
    Disable-ADAccount -Identity $Global:ConvertToShareObject  -server $DC -credential $Global:UserCredential
    Set-ADAccountPassword -Identity $Global:ConvertToShareObject -Reset -NewPassword $Spassword -Server $dc -Credential $Global:UserCredential

    #Add all Shared mailbox to a lic group
    #Write-Host "Adding " $Global:ConvertToShareObject " to Shared Mailbox license Group " $Global:SharedMBXLicGroup
    #Add-ADGroupMember -Identity $Global:SharedMBXLicGroup -Members $Global:ConvertToShareObject -Server $dc -Credential $Global:UserCredential -Confirm:$False
    #Adding to the ShareMBX lic group is in group removal function
    }

Function AttributeCalculation ($GlobalAttribCalObject)
    {
    #6/4
    If($Global:FinalTemplate.Department -like ("*Regional Office*"))  

        {
        Write-host "Regional Office samAccountname logic" 
        $Global:samaccountNamePrefix = $Global:FinalTemplate.samAccountname.SubString(0,4)
        }
        else
        {
        Write-Host "Central Office sammaccountname logic in use"
        $Global:samaccountNamePrefix = $Global:FinalTemplate.samAccountname.SubString(0,2)
        }
    
    Write-Host "This is the samaccountNamePrefix " $Global:samaccountNamePrefix

    

    if($GlobalAttribCalObject.initials -eq $null)
        {
        $Global:NewsamAccountName = $Global:samaccountNamePrefix + $GlobalAttribCalObject.givenname.SubString(0,1) + "X" + $GlobalAttribCalObject.SurName.SubString(0,1)        
        }
        else
        {
        $Global:NewsamAccountName = $Global:samaccountNamePrefix + $GlobalAttribCalObject.givenname.SubString(0,1) + $GlobalAttribCalObject.initials  + $GlobalAttribCalObject.SurName.SubString(0,1) 
        }

    $Global:NewsamAccountName = $Global:NewsamAccountName.toUpper()        
    Write-host "The new with prefix samaccountname " $Global:NewsamAccountName    
    Write-host "Template HomeDir " $Global:FinalTemplate.HomeDirectory
    $Global:HomeDriveFolder = $Global:FinalTemplate.HomeDirectory
    #This needs to be determined based on department for Central Office
    If($Global:HomeDriveFolder -eq $null)      
        {
        Write-Host "The template has no HomeDrive, OneDrive logic in use"
        #$blnHasOneDrive = $true     
        }
        else
        {
            #**Updated 8/5/2021**
            #$Global:HomeDrivePath = $Global:HomeDriveFolder.Substring(0,$Global:HomeDriveFolder.Length -7)
            $Global:HomeDrivePath = $Global:HomeDriveFolder.Substring(0,$Global:HomeDriveFolder.LastIndexOf("\"))
            Write-Host "The homedrive path after cleanup is: " $Global:HomeDrivePath
            $Global:HomeDrivePath =  $Global:HomeDrivePath + "\" + $Global:NewsamAccountName
            Write-host "This is the Homedrive " $Global:HomeDrivePath
        }
    


    }

Function RemoveSecondaryAddresses($Global:MBXtoBeCleaned)
    {
    Write-Host "Enter email address removal function with " $Global:MBXtoBeCleaned.distinguishedname " there are " $Global:MBXtoBeCleaned.emailaddresses.Count " emailaddresses"
    $addresses = @()
    Foreach($addr in $Global:MBXtoBeCleaned.emailaddresses)
        {
        Write-host "Analyzing " $addr 
        #Remove all secondary addresses based on the case of the SMTP header type
        If($addr -clike "*smtp:*")
            {
            Write-Host "Removing " $addr "  for " $Global:MBXtoBeCleaned.primarysmtpaddress
            #$addresses.Remove($addr)
            #Write-Host "The remaining email addresses are " $addresses
            } 
            else
            {
            $addresses += $addr
            }      
        
        }
   
    
    Write-host "The final email addressses are " $addresses " a total of " $addresses.Count
    Set-RemoteMailbox -Identity $Global:MBXtoBeCleaned.distinguishedname -emailaddresses $addresses
    }

Function AddGroupMembers ($Global:GroupAddObject)
    {
    Write-Host "Begin Group Membership population"
     $Tempuser = $null
    $Tempuser = Get-ADUser $Global:FinalTemplate.distinguishedname  -Properties MemberOf -Server $dc -Credential $Global:UserCredential
    $GroupsToAdd = $null
    $GroupsToAdd = New-Object System.Collections.ArrayList
        foreach ($group in $Tempuser.MemberOf) 
        {
            $GroupsToAdd.Add((Get-ADGroup $group -Server $dc -Credential $Global:UserCredential).SamAccountName)
        }
    Foreach($Group in $GroupsToAdd)
        {
            If($Group -notlike "*Domain Users*" -and $Group -notlike "*Domain Admins*")
            {
            Write-Host "Making " $objTextBoxAgencyPrimarySMTP.Text " a member of " $Group
            Add-ADGroupMember -identity $Group -Members $Global:GroupAddObject -Server $dc -Credential $Global:UserCredential
            }
        }

    }
    

Function RemoveGroupMembers($Global:GroupRemoveObject)
    {
                     
                     Write-Host "Enter Remove GroupMembers with " $Global:GroupRemoveObject 
                     Sleep -Seconds 10
                     #These must be the DN of the groups to NOT remove in quotes separated by a ,
                     $ArrGroupsNotToRemove = "CN=Domain Users,CN=Users,DC=yellow,DC=local"
                     $Tempuser = $null
                     $Tempuser = Get-ADUser $Global:GroupRemoveObject -Properties MemberOf -Server $dc -Credential $Global:UserCredential
                     $RGroupsToRemove = $null
                     $RGroupsToRemove = New-Object System.Collections.ArrayList
                         foreach ($group in $Tempuser.MemberOf)
                         {
                                 if($ArrGroupsNotToRemove -notcontains $group)
                                 {
                                     Write-Host "Adding " $group " to the remove list"
                                     $RGroupsToRemove.Add((Get-ADGroup $group -Server $dc -Credential $Global:UserCredential).SamAccountName)
                                 }
                             else
                                 {
                                     Write-host $group " was not added to the groups to remove list"
                                 }
                         }
             
                         Write-Host "Begin Group removal"
                         Foreach($RGroup in $RGroupsToRemove)
                             {
                             Write-Host "Removing " $RGroup " from " $Global:GroupRemoveObject 
                             Remove-ADGroupMember -Identity $RGroup -member $Global:GroupRemoveObject  -Server $dc -Credential $Global:UserCredential -Confirm:$false
                             }
             

    
    }

Function TranferUser {
Write-Host "Enter User Transfer Function"
            
             #Null CA2 and CA4 to clear out any previous transfers
            set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute2 "" -customattribute4 ""
            set-remotemailbox -identity $Global:SourceUser.distinguishedname -EmailAddressPolicyEnabled:$false
            $GlobalAttribCalObject = $Global:SourceUser
            AttributeCalculation $GlobalAttribCalObject
            
            Set-ADuser -identity $Global:SourceUser.distinguishedname`
                -Description $Global:FinalTemplate.Description`
                -streetaddress $Global:FinalTemplate.streetaddress`
                -City $Global:FinalTemplate.City`
                -PostalCode $Global:FinalTemplate.postalcode`
                -Office $Global:FinalTemplate.Office`
                -company $Global:FinalTemplate.company`
                -department $Global:FinalTemplate.department`
                -fax $Global:FinalTemplate.fax`
                -Server $dc -Credential $Global:UserCredential

           # if($blnHasOneDrive -ne $true)
                #{
                    Write-Host "Setting home drive on target user " $Global:SourceUser.distinguishedname
                    Set-ADuser -identity $Global:SourceUser.distinguishedname -HomeDirectory $Global:HomeDrivePath -Server $dc -Credential $Global:UserCredential
                #}
            Set-ADuser -identity $Global:SourceUser.distinguishedname -ScriptPath $Global:FinalTemplate.ScriptPath -Server $dc -Credential $Global:UserCredential
            Set-ADuser -identity $Global:SourceUser.distinguishedname -Title $Global:FinalTemplate.title -Server $dc -Credential $Global:UserCredential
            Set-ADuser -identity $Global:SourceUser.distinguishedname -SamAccountName $Global:NewsamAccountName -Server $dc -Credential $Global:UserCredential
            

           #12/2/21 
           $neweAppsProxyAddress = "smtp:" + $Global:NewsamAccountName + "@jorenby.us"
           Write-Host "Adding new eApps proxy address of " $neweAppsProxyAddress
           set-aduser -identity $Global:SourceUser.distinguishedname  -Add @{Proxyaddresses=$neweAppsProxyAddress} -Server $dc -Credential $Global:UserCredential

           $newRRAAddress = $Global:NewsamAccountName + "@M365x81231015.mail.onmicrosoft.com"
           Write-Host "Changing remote routing address to " $neweAppsProxyAddress
           set-remotemailbox -identity $Global:SourceUser.distinguishedname -RemoteRoutingAddress $newRRAAddress


            $Global:GroupRemoveObject = $Global:SourceUser.distinguishedname
            RemoveGroupMembers $Global:GroupRemoveObject
          
            $Global:GroupAddObject = $Global:SourceUser.distinguishedname
            AddGroupMembers $Global:GroupAddObject
            <#
            #Set CA2 for the supervisior if one is entered
            #This will be used by the o365 processor to copy the transfering user OneDrive to the supervisors OneDrive in a folder
            #IF CA2 -ne $null  
            #Copy OneDrive to CA2's One Drive in a sub folder
            If($objTextBoxSupervisorEmail.Text -ne "")
                {
                    Write-host "Setting customattribute 2 on transfering user " $Global:SourceUser.distinguishedname " to " $objTextBoxSupervisorEmail.Text
                    #Set CA2 and CA4. CA4 is the marker of whether the OD content has been copied to the supervisor (CA2 value)
                    #UT-OD-0 = user transfer (UT) | One Drive (OD) | 0 not done, 1 done
                    set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute2 $objTextBoxSupervisorEmail.Text -customattribute4 "UT-OD-0"
                }
            #>
            
            Write-Host "Moving " $Global:SourceUser.distinguishedname " to " $Global:TargetOU
            $Global:SourceUser | Move-ADObject -TargetPath $Global:TargetOU  -Credential $Global:UserCredential -Server $dc



}

Function SetCA4 ($Global:RMBX)
    {
    #Determine whether the user has OD drive enabled or not. If OD is enabled set CA4
    $ODADuser = get-aduser -Identity $Global:RMBX.samaccountname -Properties * -Server $dc -Credential $Global:UserCredential
    Write-Host "Process OD is " $GlobalProcessOD " The HomeDirectory is " $ODADuser.HomeDirectory

        If($ODADuser.HomeDirectory -eq $null)
            {
            $GlobalProcessOD = $True
            set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute4 "UT-OD-0"
            }
            else
            {
            $GlobalProcessOD = $False           
            set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute4  ""
            }
        
       Write-Host "Process OD is " $GlobalProcessOD " The HomeDirectory is " $ODADuser.HomeDirectory
    
    }

#End Functions

Import-Module ActiveDirectory
$dc = "yellowdc2.yellow.local"
$userforestUPNSMTPAddress  = "@jorenby.us"
$rootOU = "DC=yellow,DC=local"
$Global:UserCredential = Get-Credential -username "Yellow\administrator" -Message "Enter Exchange onpremise Credentials:"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exyellow.yellow.local/PowerShell/ -Authentication Kerberos -Credential $Global:UserCredential
$Global:SharedMBXLicGroup = "CN=Lic-425Show,OU=Groups,OU=425 Show,DC=yellow,DC=local"
$orgName ="M365x81231015"
$Global:O365Creds = Get-Credential -username "admin@M365x81231015.onmicrosoft.com" -Message "Enter O365 Credentials:"



Import-PSSession $Session -DisableNameChecking -AllowClobber -CommandName enable-remotemailbox,set-remotemailbox, new-remotemailbox,get-remotemailbox,get-organizationalunit
$RenamePrefix = "DeComm."
$TransferPrefix = "Xfer"
$InactiveOU = "OU=InActive Accounts,OU=425 Show," + $rootOU



$CoexistanceRoutingDomain = "@M365x81231015.mail.onmicrosoft.com"

#Group DN's added in this array will not be removed with the remove group operations take place
#$Global:arrLicGroups = "CN=EXO_Only,OU=O365 License,DC=dom,DC=gov","CN=All_O365,OU=O365 License,DC=dom,DC=gov","CN=ExchOL_and_PP,OU=O365 License,DC=dom,DC=gov"

Write-Host "Being On premise login"



$tlogpath = "C:\Users\djorenby\OneDrive - Microsoft\425 Show\Scripts\Logs\"
    $filedate = (get-date).ToString().Replace("/","-") | ForEach {$_ -replace ":","."}
    $tloguser = $filedate + ".log"
    $tlog = $tlogpath + $tloguser
    Start-Transcript -Path $tlog

$CredCheck = $null
$CredCheck = Get-ADUser -ResultSetSize 3 -Credential $Global:UserCredential -Filter * -Server $dc
If ($CredCheck.Count -ne 3)
    {
    Write-Host "Failed Credential Check - Exiting script" -ForegroundColor red -BackgroundColor Yellow
    #Exit
    }




#Default new user password
$password = "Apple*" + (Get-Random -Minimum 1000 -Maximum 9999)
 ###+ "!"
$cleartextpassword = $password
$Spassword = $password | ConvertTo-SecureString -AsPlainText -Force
$Global:currentdate = Get-Date -UFormat "%m/%d/%Y"
$Global:currentdate = $Global:currentdate.Replace(" ","")
$password = "Temp!6^" + (Get-Random -Minimum 1000 -Maximum 9999)
$SecPaswd = ConvertTo-SecureString –String $password –AsPlainText –Force

$Global:GroupRemoveObject = $null
$Global:GroupAddObject = $null
$Global:MBXtoBeCleaned = $Null
$GlobalAttribCalObject = $null
$Global:ConvertToShareObject = $null

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "User Operations"
$objForm.Size = New-Object System.Drawing.Size(380,800) 
$objForm.StartPosition = "CenterScreen"
$objform.AutoScale = $true
$objform.AutoScaleMode = 2
#$objForm.VerticalScroll.Visible = $true
#$objForm.VerticalScroll.Minimum = 100
#$objForm.VerticalScroll.Maximum = 900


$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Size(75,720)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Visible = $true
$objForm.Controls.Add($OKButton)
$OKButton.Add_Click({$x=$objTextBox.Text;$objForm.Close()
$Global:Bailout = $false
})


$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Size(150,720)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"


$CancelButton.Add_Click({$objForm.Close()
$Global:Bailout = $true
})
$objForm.Controls.Add($CancelButton)

#User to transfer label
$objLabelUsertotransfer = New-Object System.Windows.Forms.Label
$objLabelUsertotransfer.Location = New-Object System.Drawing.Size(10,20) 
$objLabelUsertotransfer.Size = New-Object System.Drawing.Size(280,20) 
$objLabelUsertotransfer.Text = "Enter users Email address"
$objForm.Controls.Add($objLabelUsertotransfer) 


#User to transfer textbox
$objTextUsertotransfer = New-Object System.Windows.Forms.TextBox 
$objTextUsertotransfer.Location = New-Object System.Drawing.Size(10,40) 
$objTextUsertotransfer.Size = New-Object System.Drawing.Size(300,500)
$objTextUsertotransfer.text = $UPNDomain
$objTextUsertotransfer.TabIndex = 1 
$objForm.Controls.Add($objTextUsertotransfer)

#User checklabel
$objLabelUserChecker = New-Object System.Windows.Forms.Label
$objLabelUserChecker.Location = New-Object System.Drawing.Size(10,60) 
$objLabelUserChecker.Size = New-Object System.Drawing.Size(300,20) 
$objLabelUserChecker.Text = "No User Entered"
$objForm.Controls.Add($objLabelUserChecker) 

$objtextUsertotransfer.Add_TextChanged({
	$Global:SourceUser = $null
    $Global:SourceUser = Get-ADUser -Filter {userPrincipalName -eq $objTextUsertotransfer.Text} -server $dc -Properties * -credential $Global:UserCredential
    Write-Host " Enter User Checker With " $objTextBoxAgencyPrimarySMTP.Text " " $Global:SourceUser
    
	if ($Global:SourceUser -ne $null)
		{
		Write-Host $Global:SourceUser.distinguishedname " found"
        $objLabelUserChecker.Text = $objTextUsertotransfer.text + " found"
        $objLabelUserChecker.ForeColor = "Black"
        $Global:blnPrimarySMTP = $True
            if($Global:TransferMode -eq "Decommission")
                {
                    $Global:blnOKTemplate  = $True
                }
		}
		else
		{
        Write-Host $Global:SourceUser.distinguishedname "Not Found"
        $objLabelUserChecker.Text = $objTextUsertotransfer.text + " Not found!"
        $objLabelUserChecker.ForeColor = "Red"
        $OKButton.Visible = $false
        $Global:blnPrimarySMTP = $False
		}

        If($Global:blnOKTemplate -eq $True -and $Global:blnPrimarySMTP -eq $True -and $objMoveOptionListBox.text -ne "")
            {
                $OKButton.Visible = $true
            }
            else {
                $OKButton.Visible = $False
            }

        Write-host "The ok template is  " $Global:blnOKTemplate " and the smtp address is " $Global:blnPrimarySMTP
        Write-host "The source user is " $Global:SourceUser.distinguishedname


    }
)
    

#Department Label
$objLabelEmailToClone = New-Object System.Windows.Forms.Label
$objLabelEmailToClone.Location = New-Object System.Drawing.Size(10,200) 
$objLabelEmailToClone.Size = New-Object System.Drawing.Size(280,20) 
$objLabelEmailToClone.Text = 'Enter email address of the user to be cloned'
$objForm.Controls.Add($objLabelEmailToClone)


#InputUPN
$objInputUPNTextBox = New-Object System.Windows.Forms.TextBox
$objInputUPNTextBox.Location = New-Object System.Drawing.Size(10,220) 
$objInputUPNTextBox.Size = New-Object System.Drawing.Size(260,20) 
$objInputUPNTextBox.Height = 80
$objInputUPNTextBox.text = $userforestUPNSMTPAddress 

$objInputUPNTextBox.TabIndex = 3
$objForm.Controls.Add($objInputUPNTextBox)
$objInputUPNTextBox.Add_TextChanged({

    Write-host "Enter Template check with " $objInputUPNTextBox.text
    $Global:FinalTemplate  = $null
    $Global:FinalTemplate  = Get-ADUser -Filter {userPrincipalName -eq $objInputUPNTextBox.Text} -server $dc -Credential $Global:UserCredential -Properties City,co,company,department,description,fax,homedirectory,homedrive,memberOf,PostalCode,st,state,streetaddress,title,office,scriptPath,country
    
    If($Global:FinalTemplate -ne $null)
        {
            $objLabelUPNInputResults.Text = "Template at " + $Global:FinalTemplate.distinguishedname
            $Global:blnOKTemplate = $True
            $Global:TargetOU = $Global:FinalTemplate.distinguishedname.substring($Global:FinalTemplate.DistinguishedName.IndexOf("OU="))
            Write-Host "Found " $Global:TargetOU " for the new user " $objTextBoxAgencyPrimarySMTP.text  
        } 
        else {
            $objLabelUPNInputResults.Text = "Template Not Found"
            $Global:blnOKTemplate = $false
        } 
        
  

    If($Global:blnOKTemplate -eq $True -and $Global:blnPrimarySMTP -eq $True -and $objMoveOptionListBox.text -ne "")
        {
            $OKButton.Visible = $true
        }
        else {
            $OKButton.Visible = $False
        }

    Write-host "The ok template is  " $Global:blnOKTemplate " and the smtp address is " $Global:blnPrimarySMTP
    Write-host "The Global Template is " $Global:FinalTemplate.distinguishedname


})


#Input Disable label
$objLabelUPNInputResults = New-Object System.Windows.Forms.Label
$objLabelUPNInputResults.Location = New-Object System.Drawing.Size(10,260) 
$objLabelUPNInputResults.Size = New-Object System.Drawing.Size(280,40) 
$objLabelUPNInputResults.Text = "Template Not Found"
$objLabelUPNInputResults.TabStop = $False
$objForm.Controls.Add($objLabelUPNInputResults) 



#Move Option Label
$objLabelMoveOption = New-Object System.Windows.Forms.Label
$objLabelMoveOption.Location = New-Object System.Drawing.Size(10,100) #
$objLabelMoveOption.Size = New-Object System.Drawing.Size(140,20) 
$objLabelMoveOption.Text = 'Select User Operation'
$objForm.Controls.Add($objLabelMoveOption)

##DepartKeepSupAccessCheckbox
#$objchkDepartKeepSupAccess = New-Object System.Windows.Forms.CheckBox 
#$objchkDepartKeepSupAccess.Location = New-Object System.Drawing.Size(200,200) 
#$objchkDepartKeepSupAccess.Size = New-Object System.Drawing.Size(200,20)
#$objchkDepartKeepSupAccess.checked = $False
#$objchkDepartKeepSupAccess.text = "Supervisor Retain Access"
#$objForm.Controls.Add($objchkDepartKeepSupAccess)


##Move Option Dropdown
$objMoveOptionListBox = New-Object System.Windows.Forms.ListBox 
$objMoveOptionListBox.Location = New-Object System.Drawing.Size(10,120) 
$objMoveOptionListBox.Size = New-Object System.Drawing.Size(300,20) 
$objMoveOptionListBox.Height = 60
$objMoveOptionListBox.sorted = $true
$objMoveOptionListBox.TabIndex = 2
$objForm.Controls.Add($objMoveOptionListBox)

[void] $objMoveOptionListBox.Items.Add('Transfer Employee')
[void] $objMoveOptionListBox.Items.Add('Decommission')

$objMoveOptionListBox.add_Click({
Switch ($objMoveOptionListBox.Text)
    {
    
    'Transfer Employee'
        {
        #$objLabelChosenTemplate.visible = $True
        Write-host "Transfer selected"
        $Global:TransferMode = 'Transfer Employee'
        #$objDepartmentListBox.visible = $True
        #$objTitleListBox.visible = $True
        #$objchkDepartKeepSupAccess.visible = $False
        #$objLabelTemplateFound.Visible = $True
        #$objLabelTargetOU.Visible = $True
        #$objLabelNewUserOU.Visible = $True
        #$objLabelTemplateChoice.Visible = $True
        #$objTemplateChoiceListBox.Visible = $True
        $objLabelSupervisorEmail.visible = $False
        $objTextBoxSupervisorEmail.visible = $False
        $objLabelEmailToClone.Visible = $True
        $objInputUPNTextBox.Visible = $True
        $objLabelUPNInputResults.Visible = $True
        $OKButton.Visible = $False
        $Global:blnOKTemplate = $False
        $Global:blnPrimarySMTP = $False
        Write-host "Selected Mode is " $Global:TransferMode
        $InputName = $null
        $InputName = $objtextUsertotransfer.Text
        $objtextUsertotransfer.Text = ""
        $objtextUsertotransfer.Text = $InputName
        $InputName = $null
        $InputName = $objInputUPNTextBox.Text
        $objInputUPNTextBox.Text = ""
        $objInputUPNTextBox.Text = $InputName
        If($Global:blnOKTemplate -eq $True -and $Global:blnPrimarySMTP -eq $True -and $objMoveOptionListBox.text -ne "")
        {
            $OKButton.Visible = $true
        }
        else {
            $OKButton.Visible = $False
        }
    

        }

    'Decommission'
        {
        #$objLabelChosenTemplate.visible = $False
        Write-host "Decommission selected"
        $Global:TransferMode = 'Decommission'
        #$objDepartmentListBox.visible = $False
        #$objTitleListBox.visible = $False
        #$objchkDepartKeepSupAccess.visible = $True
        #$objLabelTemplateFound.Visible = $False
        #$objLabelTargetOU.Visible = $False
        #$objLabelTemplateChoice.Visible = $True
        ##$objTemplateChoiceListBox.Visible = $False
        $objLabelSupervisorEmail.visible = $True
        $objTextBoxSupervisorEmail.visible = $True
        $OKButton.Visible = $False
        $objLabelEmailToClone.Visible = $False
        $objInputUPNTextBox.Visible = $False
        $objLabelUPNInputResults.Visible = $False
        Write-host "Selected Mode is " $Global:TransferMode
        $Global:blnOKTemplate = $True
        $InputName = $null
        $InputName = $objtextUsertotransfer.Text
        $objtextUsertotransfer.Text = ""
        $objtextUsertotransfer.Text = $InputName
        $InputName = $null
        $InputName = $objInputUPNTextBox.Text
        $objInputUPNTextBox.Text = ""
        $objInputUPNTextBox.Text = $InputName
        $Global:blnOKTemplate = $True
        If($Global:blnOKTemplate -eq $True -and $Global:blnPrimarySMTP -eq $True)
        {
            $OKButton.Visible = $true
        }
        else {
            $OKButton.Visible = $False
        }
        }    
    
    }
}
)


#ChosenTemplate
#$objLabelChosenTemplate = New-Object System.Windows.Forms.Label
#$objLabelChosenTemplate.Location = New-Object System.Drawing.Size(10,420)
#$objLabelChosenTemplate.Size = New-Object System.Drawing.Size(400,20) 
#$objLabelChosenTemplate.Text = "<No Template>"
#$objLabelChosenTemplate.visible = $False
#$objForm.Controls.Add($objLabelChosenTemplate)


#Supervisor label
$objLabelSupervisorEmail = New-Object System.Windows.Forms.Label
$objLabelSupervisorEmail.Location = New-Object System.Drawing.Size(10,500) 
$objLabelSupervisorEmail.Size = New-Object System.Drawing.Size(340,20) 
$objLabelSupervisorEmail.Text = "Email address of Supervisor retaining access (if applicable)"
$objForm.Controls.Add($objLabelSupervisorEmail) 


#Supervisortextbox
$objTextBoxSupervisorEmail = New-Object System.Windows.Forms.TextBox 
$objTextBoxSupervisorEmail.Location = New-Object System.Drawing.Size(10,520) 
$objTextBoxSupervisorEmail.Size = New-Object System.Drawing.Size(300,500)
$objTextBoxSupervisorEmail.tabindex = 4
$objForm.Controls.Add($objTextBoxSupervisorEmail)



#TargetTemplateLabel
#$objLabelTemplateFound = New-Object System.Windows.Forms.Label
#$objLabelTemplateFound.Location = New-Object System.Drawing.Size(10,560) 
#$objLabelTemplateFound.Size = New-Object System.Drawing.Size(340,40) 
#$objLabelTemplateFound.Text = "Template Found - False"
#$objForm.Controls.Add($objLabelTemplateFound) 

#TargetOULabel
#$objLabelTargetOU = New-Object System.Windows.Forms.Label
#$objLabelTargetOU.Location = New-Object System.Drawing.Size(10,600) 
#$objLabelTargetOU.Size = New-Object System.Drawing.Size(340,20) 
#$objLabelTargetOU.Text = "Target OU: "
#$objForm.Controls.Add($objLabelTargetOU) 



    
    #Test to see if the OU exists
    <#
    $OUTest = $null
    $OUTest = Get-OrganizationalUnit $Global:TargetOU -DomainController $dc
    if($OUTest -ne $null)
        {
        $objLabelTargetOU.text = "Target OU: " + $Global:TargetOU
        $objLabelTargetOU.ForeColor = "Black"
        }
        else
        {
        $objLabelTargetOU.text = $Global:TargetOU + " not found"
        $objLabelTargetOU.ForeColor = "Red"
        }
       
        #If department changes need to reselect template
        $Global:FinalTemplate = $null
        #$objLabelTemplateFound.Text = "Template Found - False"
        $OKButton.Visible = $False
        #$objTemplateChoiceListBox.Items.clear()
        #$objTitleListBox.SelectedIndex = -1 
#>
#Title Dropdown label
#$objLabelTitleDropDown = New-Object System.Windows.Forms.Label
#$objLabelTitleDropDown.Location = New-Object System.Drawing.Size(10,300) 
#$objLabelTitleDropDown.Size = New-Object System.Drawing.Size(360,20) 
#$objLabelTitleDropDown.Text = 'Title Filter'
#$objForm.Controls.Add($objLabelTitleDropDown)

#Title Dropdown
#$objTitleListBox = New-Object System.Windows.Forms.ListBox 
#$objTitleListBox.Location = New-Object System.Drawing.Size(10,320) 
#$objTitleListBox.Size = New-Object System.Drawing.Size(260,20) 
#$objTitleListBox.Height = 80
#$objTitleListBox.sorted = $true


#Template Choice label
#$objLabelTemplateChoice = New-Object System.Windows.Forms.Label
#$objLabelTemplateChoice.Location = New-Object System.Drawing.Size(10,400) 
#$objLabelTemplateChoice.Size = New-Object System.Drawing.Size(360,20) 
#$objLabelTemplateChoice.Text = "Choose Template"
#$objForm.Controls.Add(#$objLabelTemplateChoice)





$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
$Global:blnPrimarySMTP = $False
$Global:blnOKTemplate = $False
[void] $objForm.ShowDialog()


if($Global:Bailout -eq $true)
    {
    Write-host "Cancel detected exit script with no changes" -ForegroundColor Red -BackgroundColor Yellow
    $objForm.Close()
    Stop-transcript
    exit
    }

If($Global:SourceUser.distinguishedname -eq $null)
    {
        Write-host "Could not find AD user " $objTextUsertotransfer.text 
        Write-Host "Exiting script" -ForegroundColor Red -BackgroundColor Yellow
        exit    
    }
    else
    {
        Write-host "Begin Remote Mailbox check for " $Global:SourceUser.distinguishedname
        $Global:RMBX = Get-remotemailbox -Identity $Global:SourceUser.distinguishedname
        Write-host "Entering SourceMailboxCheck with " $Global:RMBX                
        SourceMailboxCheck $Global:RMBX
        Write-Host "Post sourcemailbox check"    
    }


Switch ($Global:TransferMode)
    {
    
    
    "Transfer Employee"
        {
         Write-Host "Enter Transfer Employee"
         

        If($Global:TargetOU -eq $null)
            {
            Write-Host "Could not find the target OU for " $Global:SourceUser.distinguishedname " exiting with no changes" -ForegroundColor Red -BackgroundColor Yellow
            Exit    
            }
            else
            {
            Write-Host "Running Transfer Employee Logic on " $Global:SourceUser.distinguishedname

            }
        TranferUser
        }

   "Decommission"
            {            
               

                #Are we adding a number on the first user?
                Connect-PnPOnline -Url https://$orgname.sharepoint.com/sites/UserArchive -Credential $Global:O365Creds
                #Confirm permissions needed to query UserArchive site Folders
                $ArchiveFolders = $null
                #Could not use Get-pnpfolder because it recursed
                #$ArchiveFolders = Get-PnPFolder -List "Shared Documents"
                #https://www.sharepointdiary.com/2018/09/sharepoint-online-get-all-folders-from-list-using-powershell.html
                $ctx=Get-PnPContext
                $FolderRelativeURL = "/sites/userarchive/Shared Documents"
                $ArchiveFolders = $Ctx.web.GetFolderByServerRelativeUrl($FolderRelativeURL).Folders
                $Ctx.Load($ArchiveFolders)
                $Ctx.ExecuteQuery()
                #Create a search filter of First.last*
                $SearchFolderName =  [scriptblock]::create($Global:SourceUser.GivenName + "." + $Global:SourceUser.Surname + "*")
                $ConflictFolders = $null
                #Sort that results and pick the last one which should have the highest suffix
                [array]$ConflictFolders = $ArchiveFolders.name -like $SearchFolderName | sort | select -Last 1
                
                If($ConflictFolders)
                    {
                     Write-Host "The folder with the highest number is " $ConflictFolders
                     $RND = $ConflictFolders.substring($ConflictFolders.lastindexof(".")+ 1)
                      If($RND -match '\d+')
                        {
                        Write-Host "The last character is a number of " $RND " proceeding to increment by 1"
                        #Convert to integer to increment by 1
                        $RND = [int]$RND
                        #Add 1
                        $RND ++
                        #Convert back to string so we can use it in names
                        $RND = [string]$RND
                        }
                        else
                        {
                        #We have no numbers set the value to 1
                        Write-host "The user archive folder " $ConflictFolders " has no trailing number. The rename suffix will be 1"
                        $RND = "1"    
                        }
                    }
                    else
                    {
                        Write-host "Could not find the folder " $SearchFolderName " rename suffix will be 1"
                        $RND = "1"    
                    }

                
                    $newPrimarySMTPUPN = $RenamePrefix + $Global:RMBX.PrimarySMTPAddress.split("@")[0] +  "." + $RND + "@" + $Global:RMBX.PrimarySMTPAddress.split("@")[1] 
                    Write-Host "The decommission new UPN / SMTP is " $newPrimarySMTPUPN
                    $NewSam = "DC." + $Global:RMBX.samaccountname + "." + $RND 
                    Write-Host "The decommission newSamAccountname is " $NewSam
                    $newDisplayName = $RenamePrefix  + $Global:RMBX.PrimarySMTPAddress.split("@")[0] +  "." + $RND
                    Write-Host "The decommission new DisplayName " $newDisplayName
                    
                    
                    Write-Host "Decomissioning " $Global:SourceUser.distinguishedname
                    #Null CA2 and CA4 to clear out any previous transfers
                    set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute2 "" -customattribute4 ""
                    set-remotemailbox -identity $Global:SourceUser.distinguishedname -DisplayName $newDisplayName -hiddenfromaddresslistsenabled:$true -customattribute1 $Global:currentdate -customattribute2 $objTextBoxSupervisorEmail.Text 
    
                #Set CA4 
                if($objTextBoxSupervisorEmail.Text -ne "")
                    {
                        set-remotemailbox -identity $Global:SourceUser.distinguishedname -customattribute4 "DC-EXO-0,DC-OD-0"
                    }
                
                                    
                $Global:ConvertToShareObject = $Global:SourceUser.distinguishedname 
                ConvertToShare $Global:ConvertToShareObject
                
                $DeComPS = $newPrimarySMTPUPN.trim()                
                $DecomAlias = $Global:RMBX.alias + $RND
                Write-Host "GS is " $Global:SourceUser.distinguishedname
                
                Write-Host "The decom SMTP address is " $DeComPS.tostring() 
                Write-Host "The decom alias is " $DecomAlias
                Set-RemoteMailbox -Identity $Global:SourceUser.distinguishedname -EmailAddressPolicyEnabled:$false
                Write-Host "Post disabling email address policy"
                Sleep -Seconds 10
                Write-host "Post sleep"
                Set-remotemailbox -Identity $Global:SourceUser.distinguishedname -PrimarySmtpAddress $DeComPS -Alias $DecomAlias -Verbose -Debug
                Write-host "Post changing primary smtp address for " $Global:SourceUser.distinguishedname " to " $DeComPS
                Sleep -Seconds 20
                $PostPSChangeMBX = $null
                #Get the mailbox after primarySMTP address change to have most current proxys
                $PostPSChangeMBX = get-remotemailbox -identity $DeComPS

                if($PostPSChangeMBX -eq $null)
                    {
                    Write-host "Could not find remote mailbox " $DeComPS " after primary smtp address change to " $DeComPS " Exting Script" -ForegroundColor Red -BackgroundColor Yellow
                    exit 
                    }

                 
                if($PostPSChangeMBX -is [Array])
                    {
                    Write-Host "Multiple remotemailboxs returned for " $DeComPS " Exiting script" -ForegroundColor Red -BackgroundColor Yellow
                    Exit
                    }
                        
                 
                $AddrCounter = 0
                $Global:MBXtoBeCleaned = $PostPSChangeMBX
                RemoveSecondaryAddresses $Global:MBXtoBeCleaned
                #Add back in RRA after address cleaning
                $newRRA = $RenamePrefix +  $RND + $Global:RMBX.remoteroutingaddress.replace("SMTP:","")
                Write-host "The decommission RRA is " $newRRA 
                
                Set-RemoteMailbox -Identity $Global:SourceUser.distinguishedname -RemoteRoutingAddress $newRRA
                #Need to add $RRA to proxys
                $ProxyRRA = $newRRA.tolower()
                Set-RemoteMailbox -Identity $Global:SourceUser.distinguishedname -EmailAddresses @{add=$ProxyRRA}
            
                $Global:GroupRemoveObject = $Global:SourceUser.distinguishedname
                RemoveGroupMembers $Global:GroupRemoveObject

                #Add all Shared mailbox to a lic group
                #Write-Host "Adding "  $Global:SourceUser.distinguishedname " to Shared Mailbox license Group " $Global:SharedMBXLicGroup
                #Add-ADGroupMember -Identity $Global:SharedMBXLicGroup -Members  $Global:SourceUser.distinguishedname -Server $dc -Credential $Global:UserCredential -Confirm:$False 

                

                #$DecomUser = get-aduser -Identity $Global:DecomRMBX.samaccountname -server $DC -credential $Global:UserCredential
                $newCN = $RenamePrefix + $Global:SourceUser.Name + $RND
                Set-Aduser -Identity $Global:SourceUser.distinguishedname -SamAccountName $NewSam -UserPrincipalName $newPrimarySMTPUPN -server $DC -credential $Global:UserCredential
                Write-Host "The new CN is " $NewCN
                Rename-ADObject $Global:SourceUser.distinguishedname -NewName $newCN -server $DC -credential $Global:UserCredential
                Sleep -Seconds 5 
                
                
                $RenamedCNUser = $null
                $RenamedCNUser = Get-ADUser -Identity $NewSam -server $DC -credential $Global:UserCredential
                If($RenamedCNUser -ne $null)
                    {
                    Write-Host "Moving " $RenamedCNUser.distinguishedname " to " $InactiveOU
                    $RenamedCNUser.distinguishedname | Move-ADObject -TargetPath $InactiveOU -server $DC -credential $Global:UserCredential
                    Write-Host "Post moving " $RenamedCNUser.distinguishedname " to " $InactiveOU
                    }
                    else
                    {
                    Write-host "Could not find AD user after CN rename " $NewSam " to move to " $InactiveOU                    
                    }

                Write-Host "Decommission complete for " $RenamedCNUser.distinguishedname
            

        }
       Default
        {
        Write-Host "Transfer mode not found"        
        }
    
    }
    #Invoke-Command -ComputerName  yellowdc2.yellow.local -ScriptBlock {Import-Module 'C:\Program Files\Microsoft Azure AD Sync\Extensions\AADConnector.psm1'} 
    #Invoke-Command -ComputerName  yellowdc2.yellow.local -ScriptBlock {Start-ADSyncSyncCycle -PolicyType delta}     
Write-Host "Script complete"