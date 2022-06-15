
#5/4/22 
#Access Package names should avoid special characters such as ()' spaces are ok
#Install-PackageProvider NuGet 
#Install-Module PowerShellGet 
#Install-Module AzureAD
#Install-Module MicrosoftTeams
#Install-Module Az.Account
#Install-Module -Name MSAL.PS -Scope CurrentUser
#Install-Module -Name PnP.PowerShell -Scope CurrentUser

Function GetCatalog{($ChoosenCatalogID)
$Global:AccessPackageCatalog = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackagecatalogs/$AccessPackageCatalogID"
}


function GetAccessPackage {($Global:TargetAccessPackageName) 
    $AccessPackage = $null
    $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$Global:TargetAccessPackageName'"
    $AccessPackagesPolicies = $null
    $AccessPackageID = $AccessPackage.value.id
    $Global:ChoosenAccessPolicyID =$AccessPackage.value.id

}



Function ConvertEmailToB2BUPN{($InviteUserEmailaddress)    
#Write-Host "Enter B2B UPN Conversion"
   $Global:B2BUPN = $InviteUserEmailaddress.replace("@","_") + "#EXT#@" + $TenantName

}

Function ListAccessPackages{($FinalCatalogID)
    #Write-Host "Listing All Access Packages"
    $AllAccessPackages = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?"
    $AllAccessPackagesDisplayNamelist = $AllAccessPackages.value | select displayname,catalogId 
    #Write-host " The Access Packages are " $AllAccessPackagesDisplayNamelist

     $Global:AccessPackagesTable= @()
        $no = 0
        $APCounter = 1
        #Write-host "The Catalog ID to match is " $FinalCatalogID
         foreach ($PD in $AllAccessPackagesDisplayNamelist)
         {
            
            #Write-Host "Looking at " $PD
            If($PD.catalogid -eq $FinalCatalogID)
                {
                    $no++
                    $o = [PSCustomObject]@{
                        Number = $no
                        'Access Package DisplayName' = $PD.DisplayName
                        }
                #Write-host "Adding " $o
                $Global:AccessPackagesTable += $o
                $APCounter ++
                }
        }
     $Global:AccessPackagesTable | sort number
    }

    Function ListCatalogs{
        #Write-Host "Listing All Catalogs"
        $AllCatalogs = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs"
        $AllCatalogsDisplayNamelist = $AllCatalogs.value | select displayname,id
         $Global:CatalogsTable= @()
            $no = 0
            $AllCatalogsDisplayNamelist | foreach {
                        $no++
                        $o = [PSCustomObject]@{
                            Number = $no
                            'Catalog DisplayName' = $_.DisplayName
                            'Catalog ID' = $_.id
                            }
                    #Write-Host "Adding catalog " $o        
                    $Global:CatalogsTable += $o
        }
            $Global:CatalogsTable |sort number
    
        }

Function AddB2BUserToAccessPackage{($Global:TargetAccessPackageName,$Global:B2BUPN)
    Write-Host "Searching for user " $Global:B2BUPN " to add for the Access Package"
    $SupTest = $null
    $SupTest = get-AzureADUser -ObjectId $Global:B2BUPN            
    If(!$SupTest)
    {
        Write-host "Could not find B2B User " $Global:B2BUPN " to add to Access Package " $Global:TargetAccessPackageName " exiting script"
        exit
    }
    else
    {
        $AccessPackage = $null
        $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$Global:TargetAccessPackageName'"
        #Enumerate Access Package Policies
        #Write-Host "Listing All Access Package Policies"
        $AccessPackagesPolicies = $null
        $AccessPackageID = $AccessPackage.value.id
        $Global:ChoosenAccessPolicyID =$AccessPackage.value.id
        $AccessPackagesPolicies = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentPolicies?`$filter=accessPackageId eq '$AccessPackageID'" -Method Get
        $AccessPackagesPoliciesDisplayNamelist = $null
        [array]$AccessPackagesPoliciesDisplayNamelist = $AccessPackagesPolicies.value | select displayname,ID

        $Global:AccessPackagesPolicyTable = $null
        $Global:AccessPackagesPolicyTable= @()
            $no = 0
            $AccessPackagesPoliciesDisplayNamelist | foreach {
                $no++
                $o = [PSCustomObject]@{
                    Number = $no
                    'Access Package Policy DisplayName' = $_.DisplayName
                    ID = $_.ID
                }
                $Global:AccessPackagesPolicyTable+= $o
            }
            if($AccessPackagesPoliciesDisplayNamelist.Count -gt 1)
                {
                    $Global:AccessPackagesPolicyTable | sort number | FT
                    $AccessPackagePolicyNumberToAssign = Read-host "Enter the number of the Access Package Policy for user " $InviteUserEmailaddress
                }
                else
                {
                    #Write-Host "Choosing the only policy available"
                    #Since there's only on policy just hard code to index 1 also don't display the table that gives a Access Package policy choice
                    $AccessPackagePolicyNumberToAssign = 1
                }
            $Global:TargetAccessPolicyID = $null
            $Global:TargetAccessPolicyID = $Global:AccessPackagesPolicyTable | where {$_.number -eq $AccessPackagePolicyNumberToAssign} |select -ExpandProperty ID
            $AccessPackagePolicy = $null
            
            $AccessPackagePolicy = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentPolicies?`$filter=ID eq '$Global:TargetAccessPolicyID'" -Method Get 
        
        $accessPackageAssignment = $null
        
        $accessPackageAssignment = @{
                'targetId' = $SupTest.ObjectId;
                'assignmentPolicyId' = $AccessPackagePolicy.value.id;
                'accessPackageId' = $AccessPackage.value.id
            }
        
        #Forces Approval even on a direct assignment
        # 1/27 does not work
        #Sent same JSON request as the portal confirmed with Fidder
        #Bug or not implemented yet in Graph
        #Sent as an array of One
        #The @( @{ } ) Creates an array of one object in the JSON object
        #https://stackoverflow.com/questions/18662967/convertto-json-an-array-with-a-single-item 
        $accessPackageAssignmentParameters = $null
        $accessPackageAssignmentParameters = @(
            @{
            'Name' = 'IsApprovalRequired'
            'Value' = 'true'
            }
        )
       #End bug

        $AccessPackageAssignmentAdd = $null 
        $AccessPackageAssignmentAdd = @{            
            'requestType' = 'AdminAdd';
            'accessPackageAssignment' = $accessPackageAssignment
            'parameters' = $accessPackageAssignmentParameters
        }

        Write-host "Assignment JSON"
        $AccessPackageAssignmentAdd
        $AssignPackageJSON = $null
        $AssignPackageJSON = ConvertTo-Json -InputObject $AccessPackageAssignmentAdd
        
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentRequests" -Method Post -ContentType 'application/json' -Body $AssignPackageJSON
    }
}   

Function Get-AuthToken ($AzureTenantID,$SubscriptionID,$secret,$ApplicationID) 
{
    
    try 
        {   
            $connectionDetails = @{
                'TenantId'    = $AzureTenantID
                'ClientId'    = $AzureEnterpriseAppforGraphAccess_ClientID
                'Interactive' = $true
            }
            $connectionDetails
            $TokenResult = Get-MsalToken @connectionDetails -Scopes "EntitlementManagement.ReadWrite.All" 
            Write-Output $TokenResult    
        }       
    
    catch
        {
        Throw
        Write-Host "An error occurred when try to get an Access token for Graph. Exiting script"
        Exit   
        }

}

function SendB2BInvite {($InviteUserFirstName,$InviteUserEmailaddress,$InviteUserLastName,$B2BInviteURL,$InvitedUserDisplayName)

    $SendInviteBody = @{        
        'invitedUserDisplayName' = $InvitedUserDisplayName;
        'invitedUserEmailAddress' = $InviteUserEmailaddress;        
        'sendInvitationMessage' = "True";        
        'inviteRedirectUrl' = $B2BInviteURl
    }

    #
    $SendInviteBodyJSON = $null
    $SendInviteBodyJSON = ConvertTo-Json -InputObject $SendInviteBody -Depth 6
    $B2BInvitePost = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -method Post -uri "https://graph.microsoft.com/v1.0/invitations" -ContentType 'application/json' -Body $SendInviteBodyJSON
    Write-Output $B2BInvitePost
}

function SendEXTeMail {($InviteUserEmailaddress,$Global:ChoosenAccessPolicyID)
        Write-Host "Checking SharePoint site for email attachment docx"
        IF((Get-PnPfile -url "Shared Documents\internalsupportfiles\EXTemailbody.txt") -and (Get-PnPfile -url "Shared Documents\internalsupportfiles\EXT Drive for External Users.docx"))
            {
                Get-PnPFile -Url "Shared Documents\internalsupportfiles\EXTemailbody.txt" -Path $env:temp -FileName EXTemailbody.txt -AsFile -Force
                Get-PnPFile -Url "Shared Documents\internalsupportfiles\EXT Drive for External Users.docx" -Path $env:temp -FileName "EXT Drive for External Users.docx" -AsFile -Force
                $ReadyToSend = $True
            }
            else
            {
                write-host "Could not find EXTemailbody.txt and / or EXT Drive for External Users.docx in the SharePoint site. Will NOT send EXT email"
                $ReadyToSend = $false
            }

        
        if($ReadyToSend -eq $True)
        {
        $AttachmentPath = $Env:Temp + "\EXT Drive for External Users.docx"
        $Body = $null
        #If using Word, save as html filtered
        $Body = get-content ($Env:Temp + "\EXTEmailBody.txt") -raw
        #Convert to string to get rid of other properties
        $Body = $body.ToString()
        $Body = $body.Replace("#AccessPackageID#",$Global:ChoosenAccessPolicyID)
        Write-Host "Sending EXT Email to " $InviteUserEmailaddress
        #$EMailCreds = Get-Credential -Message "Enter the email address and password for the user that will be on the FROM line in the EXT Email"
        $EmailFromAddress = read-host "Enter the email address for the user that will be on the FROM line in the EXT Email"
        $Subject = "EXT|Drive SharePoint access"
        
        
        $EXTMailconnectionDetails = @{
            'TenantId'    = $AzureTenantID
            'ClientId'    = $AzureADEnterpriseApplicationForSendingEmail_ClientID
            'Interactive' = $true
        }
        Write-host "Logging in for email with account " $EmailFromAddress
        #Need to limit scopes for mail user
        $MailToken = Get-MsalToken @EXTMailconnectionDetails -Scopes "mail.send" -LoginHint $EmailFromAddress
        $ApiUrl = "https://graph.microsoft.com/v1.0/me/sendMail"
        # Create JSON Body object
        $MessageBody = $null
        $MessageBody = @{
                    'contentType' = 'Text';
                    'content' = $Body
                    }
        $AttachmentFile - $null
        Write-host "Importing attachement " $AttachmentPath " this may take a few seconds"        
        $AttachmentFile = [convert]::ToBase64String((Get-Content $AttachmentPath -Encoding Byte))
        $MessageAttachment = @()
        $MessageAttachment = @( @{
                '@odata.type' = "#microsoft.graph.fileAttachment"
                'Name'= 'EXT Drive for External Users.docx'
                'contentType' = 'application/docx'
                'contentBytes' = $Attachmentfile                
                }
        )  
        
        #The recipient element must be an array even if there is just one entry
        #The @( @{ } ) Creates an array of one object in the JSON object
        #https://stackoverflow.com/questions/18662967/convertto-json-an-array-with-a-single-item 
        $EXTMailtoRecipients = @()
        $EXTMailtoRecipients = @( 
            @{
            'emailaddress' = @{'Address' = $InviteUserEmailaddress}
            }           
        )
        
        $EXTMailRequest = $null       
        $EXTMailRequest =  @{
               'Message' = @{
                   'Subject' = $Subject;
                   'Body' = $MessageBody;
                   'toRecipients' = $EXTMailtoRecipients;
                  'attachments' = $MessageAttachment
               } 
        }

    $EXTMailJSON = ConvertTo-Json -InputObject $EXTMailRequest -Depth 4        
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($MailToken.accesstoken)"} -Uri $ApiUrl -Method Post -Body $EXTMailJSON -ContentType "application/json"
    }
}
Import-Module AzureAD


#Lab
$AzureTenantID = "c54df794-107d-4aab-84b0-b5108bebf9fa"

#This appregistration needs
#Delegated - Mail.send
#Delegated - Mail.Read
#Delegated - EntitlementManagement.All
#Delegated - UserInvite.All 
$AzureEnterpriseAppforGraphAccess_ClientID = "15d298d0-1a8b-4118-a0e5-32537868ec1f"
#May want to create seperate App registration for sending mail
$AzureADEnterpriseApplicationForSendingEmail_ClientID = "15d298d0-1a8b-4118-a0e5-32537868ec1f"
#GET https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs
$TenantName = 'M365x81231015.onmicrosoft.com'
$AccessPackageCatalogID = "0bcc1662-dc48-48dd-899f-6ebf0c53f2ab"
$SubscriptionID = "fa581805-25b6-4ac0-87d9-7230ee153b36"
$orgName ="M365x81231015"
#Path to site where the support files are held
$SPOURL ="https://M365x81231015.sharepoint.com/sites/EXTDrive"

#Folder path to support files
$SupportFilePath = "internalsupportfiles"




Connect-AzAccount -Tenant $AzureTenantID -Subscription $SubscriptionID 
$context = Get-AzContext
Connect-AzureAD -TenantId $context.Tenant.TenantId -AccountId $context.Account.Id
Connect-PnPOnline -Url $SPOURL -Interactive

$apiEndpointUri = "https://graph.microsoft.com"
$Authcode = $null
$Authcode = Get-AuthToken -ApplicationID $AzureEnterpriseAppforGraphAcces1s_ClientID -AzureTenantID $AzureTenantID -SubscriptionID $SubscriptionID


#List all Catalogs
$Global:CatalogTable = $null
$Global:CatalogTable = ListCatalogs
$Global:CatalogTable | ft
$CatalogNumber = $null
#
$CatalogNumber = Read-host "Enter the number of the catalog that contains the Access Packages you want to assign or have noted in the email"

Write-Host "The choosen number is " $CatalogNumber
$ChoosenCatalogID = $Global:CatalogTable | where {$_.number -eq $CatalogNumber} |select -ExpandProperty 'Catalog ID'
$FinalCatalogID = $null
$FinalCatalogID = GetCatalog($ChoosenCatalogID)
#Write-Host "The chosen Catalog ID is " $FinalCatalogID
$Global:AccessPackagesTable = $null
$Global:AccessPackagesTable = ListAccessPackages ($FinalCatalogID)

#List out packages in the choosen catalog
$Global:AccessPackagesTable | ft
$AccessPackageToAddB2BUserToNumber = Read-host "Enter the number of the Access Package that you wish to assign or have noted in the B2B invite email to"
$Global:TargetAccessPackageName = $Global:AccessPackagesTable | where {$_.number -eq $AccessPackageToAddB2BUserToNumber} |select -ExpandProperty 'Access Package DisplayName'
GetAccessPackage ($Global:TargetAccessPackageName) 
Write-host "The Access Package Policy ID is " $Global:ChoosenAccessPolicyID

$InviteUserEmailaddress = Read-Host "Enter the email address of the guest account to be assigned to the Access Package " $Global:TargetAccessPackageName  " The guest account will be created automatically if needed"


$AADB2BUser = Get-AzureADUser -Filter "mail eq '$InviteUserEmailaddress'" 

If(!$AADB2BUser)
    {
        Write-host "The Guest User (B2B Account) " $InviteUserEmailaddress " does not exist"
        $SendB2BInvite = Read-host "Do you want to create a Guest account and send a Guest Account invite to " $InviteUserEmailaddress"? Press y to create and send or any other key to exit script"

        switch($SendB2BInvite)
        {
            "y"
                {
                    Write-host "y was pressed"
                    $InviteUserFirstName = Read-Host "Enter the First Name of the guest account to be assigned to an Access Package"
                    $InviteUserLastName = Read-Host "Enter the Last Name of the guest account to be assigned to an Access Package"
                    $InvitedUserDisplayName = $InviteUserFirstName + " " + $InviteUserLastName
                    $CompanyName = Read-Host "Enter the Company name of the guest account to be assigned to an Access Package. Press Enter to leave blank"
                    $JobTitle = Read-Host "Enter the job title of the guest account to be assigned to an Access Package. Press Enter to leave blank"
                    #Write-host "Choose URL to be displayed in B2B invite email the user recieves" -ForegroundColor Yellow 
                    #$B2BInviteURLCoice = Read-Host "T = Teams, D = EXTDrive" 
                    $B2BInviteURL = "https://myaccess.microsoft.com/@M365x81231015.onmicrosoft.com#/access-packages/" + $Global:ChoosenAccessPolicyID
                    <#
                    Switch ($B2BInviteURLCoice) 
                        { 
                        D {Write-host "EXTDrive" ; $B2BInviteURL = "https://msM365x81231015.sharepoint.com/sites/EXTDrive"} 
                        T {Write-Host "Teams" ; $B2BInviteURL = "https://teams.microsoft.com"} 
                        Default {Write-Host "Teams" ; $B2BInviteURL = "https://teams.microsoft.com"} 
                        }
                    #>    
                    #Switched to Graph because of MFA
                    #$invitation = New-AzureADMSInvitation -InvitedUserEmailAddress $InviteUserEmailaddress -InvitedUserDisplayName $InvitedUserDisplayName -SendInvitationMessage $true -InviteRedirectUrl $B2BInviteURL -OutBuffer
                    
                    
                    SendB2BInvite ($InviteUserFirstName,$InviteUserEmailaddress,$InviteUserLastName,$B2BInviteUR,$InvitedUserDisplayName) 
                 
                    $Global:B2BUPN = $null
                    ConvertEmailToB2BUPN ($InviteUserEmailaddress)
                    $B2BUserCheckLoop = 1
                    $Global:ErrorActionPreference = "Stop"
                    Do{
                        $AADB2BUser = $null    
                        $error.Clear()
                        Write-Host "Looking for B2B account " $InviteUserEmailaddress " in loop "  $B2BUserCheckLoop 
                        try 
                        {
                            #Commerical account such as hotmail populates "OtherMails" Azure AD account populates both mail and OtherMails
                            $AADB2BUser = Get-AzureADUser -Filter "OtherMails eq '$InviteUserEmailaddress'"
                        }
                        catch 
                        {
                            Write-Host "Could not find " $InviteUserEmailaddress " in loop " $B2BUserCheckLoop " sleeping 10 seconds"
                            
                        }
            
                        If($B2BUserCheckLoop -eq 120)
                        {
                            Write-Host "Could not find " $InviteUserEmailaddress " in alotted time. Exiting script"
                            exit
                        }
                    Write-Host "Sleeping 10 Seconds"
                    start-sleep -Seconds 10
                    $B2BUserCheckLoop ++
                    }
                    Until ($AADB2BUser -ne $null) 
                    $Global:ErrorActionPreference = "Continue"     
                    
                    
                    #Set-AzureADUser -ObjectId $Global:B2BUPN -GivenName $InviteUserFirstName -Surname $InviteUserLastName
                    Set-AzureADUser -ObjectId $AADB2BUser.ObjectId  -GivenName $InviteUserFirstName -Surname $InviteUserLastName

                    If($CompanyName -ne $null)
                    {
                        Set-AzureADUser -ObjectId $AADB2BUser.ObjectId  -CompanyName $CompanyName
                    }

                    If ($JobTitle -ne $null)
                    {
                        Set-AzureADUser -ObjectId $AADB2BUser.ObjectId  -JobTitle $JobTitle
                    }

                Write-Host "Post B2B account creation"
                }
            Default
                {
                    Write-host "Did not detect a y. Exiting script"
                    exit
                }
        }
    }
    else
    {
        if($AADB2BUser.usertype -eq "Guest")
        {
            Write-host "The Guest User (B2B Account) " $InviteUserEmailaddress " already exists"
        }
        else
        {
            Write-host "The Azure AD User account " $InviteUserEmailaddress " already exists"    
        }    
    }



 

$blnAssignToAccessPackage = Read-host "Do you want to assign " $InviteUserEmailaddress " to an Access Package and bypass approvals? Press y to assign to an Access Package or any other key to continue script"
switch($blnAssignToAccessPackage)
        {
            "y"
                {
                    ConvertEmailToB2BUPN ($InviteUserEmailaddress) 
                }
            Default
                {
                    
                    Write-host "Did not detect a y. not assigning " $InviteUserEmailaddress " to any access packages"
                    
                }
        }

iF($blnAssignToAccessPackage -eq "y")
        {
         AddB2BUserToAccessPackage ($Global:TargetAccessPackageName,$Global:B2BUPN)
        }
$SendEXTEmailChoice = Read-host "Would you like to send the EXT email to " $InviteUserEmailaddress"? Press Y to send"

Switch($SendEXTEmailChoice)
    {
        "y"
        {

            SendEXTeMail ($InviteUserEmailaddress,$Global:ChoosenAccessPolicyID)
        }
        Default
        {
        Write-Host "Not sending EXT Email"                                
        }
    }
Disconnect-AzAccount
Disconnect-AzureAD
Disconnect-PnPOnline
$Authcode = $null
$MailToken = $null
Write-host "Script Complete at " (get-date)
