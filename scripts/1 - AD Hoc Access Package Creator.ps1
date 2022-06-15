#5/4/22
#Install-Module PowerShellGet 
#Install-PackageProvider NuGet 
#Install-Module AzureAD
#Install-Module MicrosoftTeams
#Install-Module Az.Accounts
#Install-Module -Name MSAL.PS -Scope CurrentUser
#Install-Module -Name PnP.PowerShell -Scope CurrentUser


function CloudWait {
    Write-Host "Sleeping 2 Seconds to allow cloud objects to be created"
    Start-Sleep -Seconds 2
    
}
Function GetCatalog{($AccessPackageCatalogID)
$Global:AccessPackageCatalog = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackagecatalogs/$AccessPackageCatalogID"
}


Function ConvertEmailToB2BUPN{($InviteUserEmailaddress)
#Write-Host "Enter B2B UPN Conversion"
$Global:B2BUPN = $InviteUserEmailaddress.replace("@","_") + "#EXT#@" + $TenantName
}

Function ListAccessPackages{
    #Write-Host "Listing All Access Packages"
    $AllAccessPackages = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?"
    $AllAccessPackagesDisplayNamelist = $AllAccessPackages.value | select displayname,catalogid
     $Global:AccessPackagesTable= @()
        $no = 0
        $AllAccessPackagesDisplayNamelist | foreach {
            #Only add Access Packages from the catalog defined in the variables
            #Remove if to list multiple catalogs
            If($_.catalogid -eq $Global:AccessPackageCatalogID)
                {
                    $no++
                    $o = [PSCustomObject]@{
                        Number = $no
                        'Access Package DisplayName' = $_.DisplayName
                        }
                $Global:AccessPackagesTable+= $o
                }
    }
        $Global:AccessPackagesTable|sort number

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
                    $Global:CatalogsTable+= $o
        }
            $Global:CatalogsTable|sort number
    
    }

function GetSingleCatalog {($AccessPackageCatalogID)
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID"
}

Function AddB2BUserToAccessPackage{($Global:TargetAccessPackageName,$Global:B2BUPN)
    $SupTest = $null
    $SupTest = get-AzureADUser -ObjectId $Global:B2BUPN            
    If(!$SupTest)
    {
        Write-host "Could not find B2B User " $Global:B2BUPN " to add to Access Package " $Global:TargetAccessPackageName
    }
    else
    {
        $AccessPackage = $null
        $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$Global:TargetAccessPackageName'"
        #Enumerate Access Package Policies
        #Write-Host "Listing All Access Package Policies"
        $AccessPackagesPolicies = $null
        $AccessPackageID = $AccessPackage.value.id
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
                'targetId' = $Suptest.ObjectId;
                'assignmentPolicyId' = $AccessPackagePolicy.value.id;
                'accessPackageId' = $AccessPackage.value.id
            }
           
        $AccessPackageAssignmentAdd = $null        
        $AccessPackageAssignmentAdd = @{
            'requestType' = 'AdminAdd';
            'accessPackageAssignment' = $accessPackageAssignment
        }

        $AssignPackageJSON = $null
        $AssignPackageJSON = ConvertTo-Json -InputObject $AccessPackageAssignmentAdd
        
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentRequests" -Method Post -ContentType 'application/json' -Body $AssignPackageJSON
    }
    }   

Function CreateSecurityGroup{($BusinessPartnerCompanyName)
$strAzureADGroupDescription = "Security Group that controls access to the SharePoint Folder " + $BusinessPartnerCompanyName
$CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
$DepartingUserAccessGroup = Get-AzureADGroup -Filter "mailnickname eq '$CleanedBusinessPartnerCompanyName'" 
If(!$DepartingUserAccessGroup)
    {
    Write-Host "Creating Azure AD security group to secure access to " $CleanedBusinessPartnerCompanyName " SharePoint folder"
    New-AzureADGroup -SecurityEnabled $true -MailEnabled $false -DisplayName $CleanedBusinessPartnerCompanyName -Description $strAzureADGroupDescription -MailNickName $CleanedBusinessPartnerCompanyName
    $DepartingUserAccessGroup = $null
        try
        {
           
        }
    catch 
        { 
            throw
            Write-Host "An error occurred during the creation of the a security group named " $CleanedBusinessPartnerCompanyName
            Write-Host "Exiting Script"
            exit
        }
    
    }
    else
    {
    #To do retry on conflict group and Access package
    Get-AzureADGroup  -Filter "mailnickname eq '$CleanedBusinessPartnerCompanyName'" 
    Write-Host $DepartingUserAccessGroup.DisplayName " group already exists"    
    }


    }

function CreateAccessPackage {($BusinessPartnerCompanyName,$AccessPackage,$FirstApprover,$SecondApprover,$AccessDuration,$RequestorJustification,$CreateMode)
    $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
    $AccessPackageCreateParams = $null
    $strAccessPackageName =  $CleanedBusinessPartnerCompanyName
    Switch ($CreateMode) 
    { 
        "Team"
        {
            $strAccessPackageDescription = "Access Package to control access to Team " + $CleanedBusinessPartnerCompanyName
        }

        "ADHoc" 
        {
            $strAccessPackageDescription = "Access Package to control access to SharePoint Folder " + $CleanedBusinessPartnerCompanyName
        }
   
    } 
    

    #Create JSON to make the Access packages
    $AccessPackageCreateParams = @{
    catalogId = $AccessPackageCatalogID;
    displayName = $strAccessPackageName;
    description= $strAccessPackageDescription
    }

    #Convert to JSON parameters
    $CreateAPJSON = $null
    $CreateAPJSON = ConvertTo-Json -InputObject $AccessPackageCreateParams

    #POST https://graph.microsoft.com/beta/identityGovernance/entitlementManagementidentityGovernance/entitlementManagement/accessPackages
    $AccessPackage = $null
    $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$strAccessPackageName'"
    If(!$AccessPackage.value.displayname)
    {    
    
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages" -Method Post  -ContentType  'application/json' -Body $CreateAPJSON
    #https://graph.microsoft.com/beta/identityGovernance/entitlementManagementidentityGovernance/entitlementManagementaccessPackageAssignmentPolicies
    #Get the newly created package
    $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$strAccessPackageName'"
    Start-Sleep -Seconds 10
    Write-host "Creating " $AccessDuration " day policy for Access Package " $AccessPackage.value.displayname
    $AccessPackageID = $AccessPackage.value.ID
    #To properly create the JSON certain data types cannot be strings
    #Non strings data types remove the quotes from around the values in the JSON file
    #Create the data types we need
    $T = [Boolean]('true')
    $F = [Boolean]('true')
    $F = $false
    $NullArray = @()

    #Create the JSON object to create the Access Package Policy
    #https://docs.microsoft.com/en-us/graph/api/resources/requestorsettings?view=graph-rest-beta
    $requestorSettings = $null

    #AllExistingDirectorySubjects - B2B accounts must already exist
    #AllExternalSubjects - B2B Invites will be sent as needed

    $requestorSettings = @{
        'scopeType' = 'AllExistingDirectorySubjects';
        #'scopeType' = 'AllExternalSubjects';
        'acceptRequests' = $T;
        'allowedRequestors' = $NullArray
    }
      
   
    $Approver1 = @()
   #The approvers element must be an array even if there is just one entry
   #The @( @{ } ) Creates an array of one object in the JSON object
   #https://stackoverflow.com/questions/18662967/convertto-json-an-array-with-a-single-item 

   If($FirstApprover -ne $null)
    {
        $Approver1 = @(
            @{
            '@odata.type' = '#microsoft.graph.singleUser';
            'isBackup'= $T;
            'id' = $FirstApprover.objectID;
            'description' = $FirstApprover.DisplayName
            }

        )
    }
    else
    {
        $Approver1 = $NullArray    
    }


   
    $ApprovalStage1 =  @{
            'approvalStageTimeOutInDays' = '14';
            'isApproverJustificationRequired' = $T;
            'isEscalationEnabled'= $F;
            'escalationTimeInMinutes' = '11520';
            'escalationApprovers' = $NullArray;
            'primaryApprovers' = $Approver1
                   
    } 

    $Approver2= @()
    If($SecondApprover -ne $null)
        {
        $Approver2 = @(
            @{
            '@odata.type' = '#microsoft.graph.singleUser';
            'isBackup'= $T;
            'id' = $SecondApprover.objectID;
            'description' = $SecondApprover.DisplayName
            }
        )
        }
        else
        {
            $Approver2 =$NullArray
        }
    $ApprovalStage2 = $null
    $ApprovalStage2 =  @{
        'approvalStageTimeOutInDays' = '14';
        'isApproverJustificationRequired' = $T;
        'isEscalationEnabled'= $F;
        'escalationTimeInMinutes' = '11520';
        'escalationApprovers' = $NullArray;
        'primaryApprovers' = $Approver2
        }
     

    $requestApprovalSettings = $null
        #No Approver / No Justification
        If(($FirstApprover -eq $null) -and ($SecondApprover -eq $null) -and ($RequestorJustification -eq $false))
        {
        $requestApprovalSettings = @{
            'isApprovalRequired' = $F;
            'isApprovalRequiredForExtension' = $F;
            'isRequestorJustificationRequired' = $F;
            'approvalMode' = 'NoApproval'
            'approvalStages' = $NullArray
            }
        }

<#
    #No Approver / Justification - 9/17/20 Not a valid combination. Requestor justifcation requires approval
    If(($FirstApprover -eq $null) -and ($SecondApprover -eq $null) -and ($RequestorJustification -eq $True))
    {
    $requestApprovalSettings = @{
        'isApprovalRequired' = $F;
        'isApprovalRequiredForExtension' = $F;
        'isRequestorJustificationRequired' = $T;
        'approvalMode' = 'NoApproval'
        'approvalStages' = $NullArray
        }
    } 
#>

#First Approver Specified
    If(($FirstApprover -ne $null) -and ($SecondApprover -eq $null) -and ($RequestorJustification -eq $false))
    {
        #Put the approval stage into an array of 1
        $ApprovalStage1 =  @(
            @{
            'approvalStageTimeOutInDays' = '14';
            'isApproverJustificationRequired' = $T;
            'isEscalationEnabled'= $F;
            'escalationTimeInMinutes' = '11520';
            'escalationApprovers' = $NullArray;
            'primaryApprovers' = $Approver1                       
            }
        )  
        $requestApprovalSettings = @{
            'isApprovalRequired' = $T;
            'isApprovalRequiredForExtension' = $F;
            'isRequestorJustificationRequired' = $F;
            'approvalMode' = 'SingleStage'
            'approvalStages' = $ApprovalStage1
        }    
    }

    If(($FirstApprover -ne $null) -and ($SecondApprover -eq $null) -and ($RequestorJustification -eq $True))
    {
        #Put the approval stage into an array of 1
        $ApprovalStage1 =  @(
            @{
            'approvalStageTimeOutInDays' = '14';
            'isApproverJustificationRequired' = $T;
            'isEscalationEnabled'= $F;
            'escalationTimeInMinutes' = '11520';
            'escalationApprovers' = $NullArray;
            'primaryApprovers' = $Approver1                       
            }
        )  
        $requestApprovalSettings = @{
            'isApprovalRequired' = $T;
            'isApprovalRequiredForExtension' = $F;
            'isRequestorJustificationRequired' = $T;
            'approvalMode' = 'SingleStage'
            'approvalStages' = $ApprovalStage1
        }    
    }

#First Approver and Second Approver specified
    If(($FirstApprover -ne $null)  -and ($SecondApprover -ne $null) -and ($RequestorJustification -eq $false))
        {
        $requestApprovalSettings = @{
            'isApprovalRequired' = $T;
            'isApprovalRequiredForExtension' = $F;
            'isRequestorJustificationRequired' = $F;
            'approvalMode' = 'Serial'
            'approvalStages' = $ApprovalStage1,$ApprovalStage2
            }
        }

        If(($FirstApprover -ne $null) -and ($SecondApprover -ne $null) -and ($RequestorJustification -eq $True))
        {
        $requestApprovalSettings = @{
            'isApprovalRequired' = $T;
            'isApprovalRequiredForExtension' = $F;
            'isRequestorJustificationRequired' = $T;
            'approvalMode' = 'Serial'
            'approvalStages' = $ApprovalStage1,$ApprovalStage2
            }
        }


    $AccessPolicyDescription = 'Access will be automatically removed after ' + $AccessDuration + ' days'
    $AccessPolicyDisplayName = $AccessDuration + ' Day Access'
    $IntialPolicyParams = @{
        'accessPackageId' = $AccessPackage.value.ID;
        'displayName' = $AccessPolicyDisplayName;
        'description' = $AccessPolicyDescription;
        'canExtend' = $T;
        'durationInDays' = $AccessDuration;
        'expirationDateTime' = $null;
        'requestorSettings' = $requestorSettings;
        'requestApprovalSettings' = $requestApprovalSettings;
        'accessReviewSettings' = $null
        }

    $accessPackageAssignmentPoliciesJSON = $null
    $accessPackageAssignmentPoliciesJSON = ConvertTo-Json -InputObject $IntialPolicyParams -Depth 6
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentPolicies" -Method Post  -ContentType  'application/json' -Body $accessPackageAssignmentPoliciesJSON    
    # 12/10
    #Add Permanent Policy 
    $requestorSettings = $null
    $requestorSettings = @{
        #'scopeType' = 'AllExistingDirectorySubjects';
        'scopeType' = 'AllExistingDirectorySubjects';
        'acceptRequests' = $T;
        'allowedRequestors' = $NullArray
    }
    $requestApprovalSettings = $null
    $requestApprovalSettings = @{
        'isApprovalRequired' = $F;
        'isApprovalRequiredForExtension' = $F;
        'isRequestorJustificationRequired' = $T;
        'approvalMode' = 'NoApproval'
        'approvalStages' = $NullArray
        }
    $AccessPolicyDescription = 'Permanent Access'
    $AccessPolicyDisplayName = 'Permanent Access'
    $IntialPolicyParams = $null
    $IntialPolicyParams = @{
        'accessPackageId' = $AccessPackage.value.ID;
        'displayName' = $AccessPolicyDisplayName;
        'description' = $AccessPolicyDescription;
        'canExtend' = $F;
        'durationInDays' = '0';
        'expirationDateTime' = $null;
        'requestorSettings' = $requestorSettings;
        'requestApprovalSettings' = $requestApprovalSettings;
        'accessReviewSettings' = $null
        }

    $accessPackageAssignmentPoliciesJSON = $null
    $accessPackageAssignmentPoliciesJSON = ConvertTo-Json -InputObject $IntialPolicyParams -Depth 6
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentPolicies" -Method Post  -ContentType  'application/json' -Body $accessPackageAssignmentPoliciesJSON    



    }
    else
    {
    Write-host "The Access Package " $AccessPackage.value.displayname " already exists. No changes will be made to the existing Access Package"
    return $AccessPackageID = $AccessPackage.value.id
    }

    
}

function AddTeamToAccessPackageCatalog{($CurrentAccessPackageCatalog,$BusinessPartnerCompanyName, $Team,$AzureTenantID)
    $TeamURL = "https://account.activedirectory.windowsazure.com/r?tenantId=" + $AzureTenantID  + "#/manageMembership?objectType=Group&objectId=" + $team.GroupId
    $AccessPackageTeam = @{
        'displayName' = $Team.displayname;
        'description' = $Team.displayname;
        'url'=$TeamURL;
        'resourceType' = 'O365 Teams Group';
        'originId' = $team.GroupId;
        'originSystem' = 'AadGroup'
    }

    $AccessPackageTeamRequest = @{
        'catalogId' = $CurrentAccessPackageCatalog.ID;
        'requestType' = 'AdminAdd';
        'justification' = '';
        'accessPackageResource' = $AccessPackageTeam
    }


    Write-Host "Checking to see if the Team " $AccessPackageTeam['displayname'] " exists in Access Package catalog " $CurrentAccessPackageCatalog.displayname
    $TeamRolename =  $AccessPackageTeam['displayname']
    $TeamRole = $null
    $TeamRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$TeamRolename'"
    If($TeamRole.value.displayname -eq $null)
        {
        Write-Host "Adding Team " $Team.DisplayName " to Access Package catalog " $CurrentAccessPackageCatalog.displayname
        $AccessPackageTeamJSON = $null
        $AccessPackageTeamJSON = ConvertTo-Json -InputObject $AccessPackageTeamRequest 
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageResourceRequests" -Method Post -ContentType 'application/json' -Body $AccessPackageTeamJSON  
        $TeamRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$TeamRolename'"
        }
        else
        {
        Write-Host $Team.DisplayName " Team already exists in catalog " $CurrentAccessPackageCatalog.displayname     
        }
} 

function AddTeamtoAccessPackage {($DepartingUserAccessGroup,$Team)
    $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
    $AccessPackage = $null
    $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$CleanedBusinessPartnerCompanyName'"
    $accesspackageid = $AccessPackage.value.id
    $ResourcesInAccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid`?`$expand=accessPackageResourceRoleScopes" -Method Get
    
    $TeamDisplayname = $Team.displayname
    $TeamRole = $null
    $TeamRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$TeamDisplayname'"
    
    #https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/fde14ace-0a28-4b51-b705-5ddffbd21eb2?$expand=accessPackageResourceRoleScopes
    Write-Host "The Access Package " $AccessPackage.value.displayname " currently has " $ResourcesInAccessPackage.accessPackageResourceRoleScopes.count " resources"
    $ResourcedisplayCounter = 1
    $ResourceCounter = 0
    #Get the resource from the Catalog
    $GroupCatalogResource = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$TeamDisplayname'"
    #Build the JSON File object
    $accessPackageResourceScope = @{
        'originId' = $GroupCatalogResource.value.originid 
        'originSystem' = 'AadGroup'
    }

    $accessPackageResource = @{
        'id' = $TeamRole.value.id;
        'resourceType' = $TeamRole.value.resourcetype;
        'originId' = $TeamRole.value.originid;
        'originSystem' = $TeamRole.value.originSystem
    }

    #To find the format of the role, add the resource to a test package then
    #List the Resources in an Access Package
    #$accessPackageAssignmentResourceRoles = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentResourceRoles" -Method Get -Headers @{'Content-type' = 'application/json'}

    #Roles are assigned in the Access Package
    #The Roles are driven by the type of resources being added
    #For example Group has member and owner
    #$AccessPackageRolesAndScopes = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid`?`$expand=accessPackageResourceRoleScopes(`$expand=accessPackageResourceRole,accessPackageResourceScope)" -Method Get -Headers @{'Content-type' = 'application/json'}
    #The AzureAD group and SPO site
    $accessPackageResourceRoleoriginId = 'Member_' + $GroupCatalogResource.value.originid

    $accessPackageResourceRole = @{
            'originId' = $accessPackageResourceRoleoriginId
            'displayName' = 'Member';
            'originSystem' = 'AadGroup';
            'accessPackageResource' = $accessPackageResource;
    }   

    $AccessPackageRoleAdd  =@{
        'accessPackageResourceRole' = $accessPackageResourceRole; 
        'accessPackageResourceScope' = $accessPackageResourceScope

    }

    $AddSecurityGroupResourceJSON = $null
    $AddSecurityGroupResourceJSON = ConvertTo-Json -InputObject $AccessPackageRoleAdd

    #Add the Group to the Access Package
    Write-Host "Adding Team " $team.DisplayName " to Access Package " $AccessPackage.value.displayname
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid/accessPackageResourceRoleScopes" -Method Post -Body $AddSecurityGroupResourceJSON -ContentType 'application/json'
    
}


function AddResourcesToAccessPackageCatalog {($CurrentAccessPackageCatalog,$ADHocSharingSPOSite,$DepartingUserAccessGroup,$BusinessPartnerCompanyName,$ADHocSharingFolder)
    $Global:ErrorActionPreference = "Continue"
    $AccessPackageSPOSite = @{
        'displayName' = $ADHocSharingFolder;
        'description' = $ADHocSharingSPOSite;
        'url'=$ADHocSharingSPOSite;
        'resourceType' = 'SharePoint Online Site';
        'originId' = $ADHocSharingSPOSite;
        'originSystem' = 'SharePointOnline'
    }

    $AccessPackageSPOSiteRequest = @{
        'catalogId' = $CurrentAccessPackageCatalog.ID;
        'requestType' = 'AdminAdd';
        'justification' = '';
        'accessPackageResource' = $AccessPackageSPOSite
    }

    #Get the SPO and Group Resource to add to the Access package
    Write-Host "Checking to see if the SharePoint Access Package " $AccessPackageSPOSite['displayname'] " exists in Access Package catalog " $CurrentAccessPackageCatalog.displayname
    $SPORolename =  $AccessPackageSPOSite['displayname']
    $SPORole = $null
    $SPORole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$SPORolename'"
    If($SPORole.value.displayname -eq $null)
        {
        Write-Host "Adding SharePoint Archive User site resource to Access Package catalog " $AccessPackageCatalog.displayname
        $AccessPackageSPOSiteJSON = $null
        $AccessPackageSPOSiteJSON = ConvertTo-Json -InputObject $AccessPackageSPOSiteRequest 
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageResourceRequests" -Method Post -ContentType 'application/json' -Body $AccessPackageSPOSiteJSON  
        $SPORole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$SPORolename'"
        }
        else
        {
        Write-Host $ADHocSharingSPOSite " SPO site already exists in catalog " $CurrentAccessPackageCatalog.displayname     
        }

    Write-host "Checking to see if Azure AD security group " $DepartingUserAccessGroup.displayname " exists in Access Package catatlog " $CurrentAccessPackageCatalog.displayname
    $AccessPackageGroup = @{
        'originId' = $DepartingUserAccessGroup.ObjectId;
        'originSystem'='AadGroup'
    }

    $AccessPackageGroupRequest = @{
        'catalogId' = $CurrentAccessPackageCatalog.id;
        'requestType' = 'AdminAdd';
        'accessPackageResource' = $AccessPackageGroup
    }
    $GroupDisplayname = $DepartingUserAccessGroup.displayname
    $GroupRole = $null
    $GroupRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$GroupDisplayname'"

    IF($GroupRole.value.displayname -eq $null)
        {
        $AccessGroupAddCounter = 1
        Do{
            $AccessPackageGroupJSON = $null
            $AccessPackageGroupJSON = ConvertTo-Json -InputObject $AccessPackageGroupRequest
            #To do 12/8 
            Write-Host "Adding Azure AD Security Group " $DepartingUserAccessGroup.displayname  " to Access Package catalog " $CurrentAccessPackageCatalog.displayname " in loop " $AccessGroupAddCounter
            $Global:ErrorActionPreference = "Stop"
            $error.Clear()
            try 
            {
                Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageResourceRequests" -Method Post -Body $AccessPackageGroupJSON -ContentType 'application/json'
            }
            catch 
            {
                Write-Host $DepartingUserAccessGroup.displayname " was not a member of catalog " $CurrentAccessPackageCatalog.displayname " in " $AccessGroupAddCounter " loops"
            }

            If($AccessGroupAddCounter -eq 120)
            {
            Write-Host "Could not add the security group " $DepartingUserAccessGroup.displayname " to Access Package Catalog " $CurrentAccessPackageCatalog.displayname  " in " $AccessGroupAddCounter " loops" -BackgroundColor Red -ForegroundColor Yellow
            Write-Host "Action needed - " $DepartingUserAccessGroup.DisplayName " must manually be added to to Access Package Catalog " $CurrentAccessPackageCatalog.displayname
            Break
            }

           #To do  12/8
           $AccessGroupAddCounter ++
    Write-host "Sleeping 10 seconds in group add to Access Catalog"
    start-Sleep -Seconds 10 
    } 
    Until ($Error.count -eq 0) 
    $Global:ErrorActionPreference = "Continue"
        $GroupRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$Groupdisplayname'"    
        }
        else
        {
        Write-Host "The security group " $DepartingUserAccessGroup.displayname " already exists in catalog " $CurrentAccessPackageCatalog.displayname
        }

}


function AddResourcestoAccessPackage {($BusinessPartnerCompanyName,$AccessPackageCatalogID,$DepartingUserAccessGroup,$ADHocSharingFolder)
    $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
    $AccessPackage = $null
    $AccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accesspackages?`$filter=displayname eq '$CleanedBusinessPartnerCompanyName'"
    $accesspackageid = $AccessPackage.value.id
    $ResourcesInAccessPackage = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid`?`$expand=accessPackageResourceRoleScopes" -Method Get
    
    $GroupDisplayname = $DepartingUserAccessGroup.displayname
    $GroupRole = $null
    $GroupRole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$GroupDisplayname'"
    
    $SPORole = $null
    $SPORole = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$ADHocSharingFolder'"
    
    
    
    #https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/fde14ace-0a28-4b51-b705-5ddffbd21eb2?$expand=accessPackageResourceRoleScopes
    Write-Host "The Access Package " $AccessPackage.value.displayname " currently has " $ResourcesInAccessPackage.accessPackageResourceRoleScopes.count " resources"
    $ResourcedisplayCounter = 1
    $ResourceCounter = 0
    #Get the resource from the Catalog
    $GroupCatalogResource = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$GroupDisplayname'"
    #Build the JSON File object
    $accessPackageResourceScope = @{
        'originId' = $GroupCatalogResource.value.originid 
        'originSystem' = 'AadGroup'
    }

    $accessPackageResource = @{
        'id' = $GroupRole.value.id;
        'resourceType' = $GroupRole.value.resourcetype;
        'originId' = $GroupRole.value.originid;
        'originSystem' = $GroupRole.value.originSystem
    }

    #To find the format of the role, add the resource to a test package then
    #List the Resources in an Access Package
    #$accessPackageAssignmentResourceRoles = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentResourceRoles" -Method Get -Headers @{'Content-type' = 'application/json'}

    #Roles are assigned in the Access Package
    #The Roles are driven by the type of resources being added
    #For example Group has member and owner
    #$AccessPackageRolesAndScopes = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid`?`$expand=accessPackageResourceRoleScopes(`$expand=accessPackageResourceRole,accessPackageResourceScope)" -Method Get -Headers @{'Content-type' = 'application/json'}
    #The AzureAD group and SPO site
    $accessPackageResourceRoleoriginId = 'Member_' + $GroupCatalogResource.value.originid

    $accessPackageResourceRole = @{
            'originId' = $accessPackageResourceRoleoriginId
            'displayName' = 'Member';
            'originSystem' = 'AadGroup';
            'accessPackageResource' = $accessPackageResource;
    }   

    $AccessPackageRoleAdd  =@{
        'accessPackageResourceRole' = $accessPackageResourceRole; 
        'accessPackageResourceScope' = $accessPackageResourceScope

    }

    $AddSecurityGroupResourceJSON = $null
    $AddSecurityGroupResourceJSON = ConvertTo-Json -InputObject $AccessPackageRoleAdd

    #Add the Group to the Access Package
    Write-Host "Assigning Azure AD Security Group Catalog Resource" $DepartingUserAccessGroup.displayname  " to Access Package " $AccessPackage.value.displayname
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid/accessPackageResourceRoleScopes" -Method Post -Body $AddSecurityGroupResourceJSON -ContentType 'application/json'

    #Repeat for the SPO Site
    #Add SPO Site to Access Package
    $SPOCatalogResource = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -method Get -uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs/$AccessPackageCatalogID/accessPackageResources?`$filter=displayname eq '$ADHocSharingFolder'"
    $accessPackageResourceScope = @{
        'originId' = $SPOCatalogResource.value.originid
        'originSystem' = 'SharePointOnline'
        #5/4/22
        "isRootScope" = $T
    }

    $accessPackageResource = @{
        'id' = $SPORole.value.id;
        'resourceType' = $SPORole.value.resourcetype;
        'originId' = $SPORole.value.originid;
        'originSystem' = $SPORole.value.originSystem
    }

    $accessPackageResourceRoleoriginId = $SPOCatalogResource.value.originid
    $accessPackageResourceRoleDisplayname = $SPOCatalogResource.value.displayName + " Visitors"

    #OriginID is an index to the different roles. Owner,Member,visitor 
    #4 is Visitors
    $accessPackageResourceRole = @{        
            'displayName' = $accessPackageResourceRoleDisplayname;
            'originSystem' = 'SharePointOnline';
            'accessPackageResource' = $accessPackageResource;
            'originId' = '4'
    }   

    $AccessPackageRoleAdd = @{
        'accessPackageResourceRole' = $accessPackageResourceRole; 
        'accessPackageResourceScope' = $accessPackageResourceScope

    }

    $AddSPOResourceJSON = $null
    $AddSPOResourceJSON = ConvertTo-Json -InputObject $AccessPackageRoleAdd

    Write-Host "Assigning SharePoint Group Catalog Resource" $SPORole.value.displayname  " to Access Package " $AccessPackage.value.displayname
    Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid/accessPackageResourceRoleScopes" -Method Post -ContentType 'application/json' -Body $AddSPOResourceJSON
    
    
    #This will not work if there's more than one group in the package$AccessPackageRolesAndScopes = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Authcode.accesstoken)"} -Credential $AzureCreds -Uri "https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackages/$accesspackageid`?`$expand=accessPackageResourceRoleScopes(`$expand=accessPackageResourceRole,accessPackageResourceScope)" -Method Get -Headers @{'Content-type' = 'application/json'}
    $GroupAccessPackageRole = $AccessPackageRolesAndScopes.accessPackageResourceRoleScopes.accessPackageResourceRole | Where {$_.displayname -like '*member*'}
    
}

Function CreateSharePointFolder{($BusinessPartnerCompanyName,$ADHocSharingSPOSite)
    $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
    $NewSPOFolder = Resolve-PnPFolder -SiteRelativePath ("Shared Documents/" + $CleanedBusinessPartnerCompanyName) 
}

function FixSharePointFolderPermissions {($BusinessPartnerCompanyName,$DepartingUserAccessGroup)
write-host "Business partner name is " $BusinessPartnerCompanyName
Write-Host "The departing access group is " $DepartingUserAccessGroup

    $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.replace(" ","")
	Write-Host "The cleaned business partner name is "  $CleanedBusinessPartnerCompanyName
	#6/17/2021
    $SPOFolderFullPath = "Shared Documents/" + $CleanedBusinessPartnerCompanyName  
	Write-Host "The SPO folder path is " $SPOFolderFullPath
    #https://docs.microsoft.com/en-us/powershell/module/sharepoint-pnp/set-pnpfolderpermission?view=sharepoint-ps
    
    write-host "The group to find is " ("c:0t.c|Tenant|" + $DepartingUserAccessGroup.objectid) 
    $Looptester = 1 
    Do {
        $Global:ErrorActionPreference = "Stop"
        $Error.clear()
            try {                
                Write-host 'Enter try in Set-PnPFolderPermission loop ' $Looptester "The folder path is " $SPOFolderFullPath
                Set-PnPFolderPermission -List 'Shared Documents' -Identity  $SPOFolderFullPath -user ("c:0t.c|Tenant|" + $DepartingUserAccessGroup.objectid) -AddRole 'Edit'
                Write-Host "Post Sucessfull Set at "  (get-date) " in loop " $Looptester ' for group ' $DepartingUserAccessGroup.DisplayName
                }
            catch
                {
                Write-host 'Enter Catch in Set-PnPFolderPermission loop ' $Looptester ' for group ' $DepartingUserAccessGroup.DisplayName
                $Error[0].Exception.Message
                Write-host "Retrying SPO permission set for group " $DepartingUserAccessGroup.DisplayName
                }
			#6/17	
        If($Looptester -eq 80)
            {
            Write-Host "Could not set PNP Folder permissions in " $Looptester " seconds" -BackgroundColor Red -ForegroundColor Yellow
            Write-Host "Action needed - " $DepartingUserAccessGroup.DisplayName " must manually be added to " $SPOFolderFullPath
            Break
            }
           $looptester ++
          
        start-Sleep -Seconds 10 
        }
    Until ($Error.count -eq 0)
    $Global:ErrorActionPreference = "Continue"
  
    Write-Host "Removing ExtDrive Vistors Read Role"
    #Set-PnPFolderPermission -List 'Shared Documents' -Identity  $SPOFolderFullPath -group "ExtDrive Visitors" -RemoveRole 'Read'
    Set-PnPFolderPermission -List 'Shared Documents' -Identity  $SPOFolderFullPath -group $ExtDriveVisitorGroupName  -RemoveRole 'Read'
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

Write-Host "Start Script"
Import-Module AzureAD


$AzureTenantID = "c54df794-107d-4aab-84b0-b5108bebf9fa"
$AzureEnterpriseAppforGraphAccess_ClientID = "15d298d0-1a8b-4118-a0e5-32537868ec1f"
#GET https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageCatalogs
$TenantName = 'M365x81231015.onmicrosoft.com'
#URL to the SharePoint site for AD HOC sharing
$ADHocSharingSPOSite = "https://M365x81231015.sharepoint.com/sites/ExtDrive"
$ADHocSharingFolder = "ExtDrive"
$AccessPackageCatalogID = "0bcc1662-dc48-48dd-899f-6ebf0c53f2ab"
$SubscriptionID = "fa581805-25b6-4ac0-87d9-7230ee153b36"
$ExtDriveVisitorGroupName = "ExtDrive Visitors" 

#06/17/2021
$error.clear()


#We use the Connect-AzAccount to get access to Key vault and other services
Connect-AzAccount -Tenant $AzureTenantID -Subscription $SubscriptionID
$context=Get-AzContext
Connect-AzureAD -TenantId $context.Tenant.TenantId -AccountId $context.Account.Id
Connect-MicrosoftTeams
# 6/17
Connect-PnPOnline  -Url $ADHocSharingSPOSite -Interactive
#Connect-ExchangeOnline -UserPrincipalName $context.Account.Id
$aztoken = Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com"


$Authcode = $null
$Authcode = Get-AuthToken -ApplicationID $AzureEnterpriseAppforGraphAcces1s_ClientID  -AzureTenantID $AzureTenantID -SubscriptionID $SubscriptionID
$T = [Boolean]('true')
$F = [Boolean]('true')
$F = $false

#Write-Host $Authcode

Write-host "Do you want to create a Team or AD Hoc Share? The default is AD Hoc Share" -ForegroundColor Yellow 
$InputCreateMode = Read-Host "T = Team, A = AD Hoc Share" 
Switch ($InputCreateMode) 
    { 
    T {Write-host "Create Team"; $CreateMode="Team"} 
    A {Write-Host "Create AD Hoc Share"; $CreateMode="ADHoc"}
    Default {Write-Host "Create AD Hoc Share"; $CreateMode="ADHoc"} 
    } 

$BusinessPartnerCompanyName = $null
$BusinessPartnerCompanyName =  Read-Host "Enter the name of the busiess partner or Team you wish to AD Hoc share with. This will be the name of the SharePoint Folder that is created and shared with the partner"



If($BusinessPartnerCompanyName -eq "")
    {
        Write-host "No Tean or Business partner name was entered. Exiting script"
        exit
    }
$AccessDuration = Read-host "Enter the number of days the users should have access for"
$CurrentAccessPackageCatalog = GetSingleCatalog ($AccessPackageCatalogID)

$InputFirstApprover =  $null
$FirstApprover = $null
$InputFirstApprover = Read-host "Enter the email address of the first approver, leave blank for no approval required"

$RequestorJustification = $False
If($InputFirstApprover -ne "")
    {
    $FirstApprover = Get-AzureADUser -Filter "mail eq '$InputFirstApprover'"

        
        $RequestorJustificationChoice = $null
        Write-host "Require Requestors to provide a justficiation (Default is No)" -ForegroundColor Yellow 
        $RequestorJustificationChoice = Read-Host " ( y / n ) " 
        Switch ($RequestorJustificationChoice) 
        { 
        Y {Write-host "Require Requestors to provide a justficiation message"; $RequestorJustification=$true} 
        N {Write-Host "No justficiation required"; $RequestorJustification=$false} 
        Default {Write-Host "No justficiation required"; $RequestorJustification=$false} 
        } 

    #Since we have a first approver ask for a second
        $InputSecondApprover = $null
        $SecondApprover = $Null
        $InputSecondApprover = Read-host "Enter the email address of the Second approver, leave blank to not assign second approver"
        if($InputSecondApprover -ne "")
            {
            $SecondApprover = Get-AzureADUser -Filter "mail eq '$InputSecondApprover'"
            }

    }




Write-Host "Creating Access Package " $BusinessPartnerCompanyName " and Access Package policy"
$fnOBJAccessPackage = CreateAccessPackage($BusinessPartnerCompanyName,$AccessDuration,$RequestorJustification,$CreateMode)
CloudWait


Switch ($CreateMode) 
    { 
    "Team" 
        {
            $CleanedBusinessPartnerCompanyName = $BusinessPartnerCompanyName.Replace(" ","")
            $TeamOwner = $null
            $TeamOwner = Read-Host "Enter the email address of the Team Owner"
            $TeamDescription = Read-Host "Enter a description of the Team"
            If($TeamOwner -ne "")
                {
                New-Team -MailNickname $CleanedBusinessPartnerCompanyName -displayname $BusinessPartnerCompanyName -Visibility "Private" -Owner $TeamOwner -Description $TeamDescription
                }
                else
                {
                New-Team -MailNickname $CleanedBusinessPartnerCompanyName -displayname $BusinessPartnerCompanyName -Visibility "Private"  -Description $TeamDescription 
                }
            CloudWait
            $Team = $null
            $TeamLoopCheck = 1
            do {
                Write-Host "Waiting for Team " $BusinessPartnerCompanyName " to become available. Wait Loop = " $TeamLoopCheck 
                $Team = get-Team -MailNickName $Cleaned$BusinessPartnerCompanyName
                Start-Sleep 2
                $TeamLoopCheck ++

                if($TeamLoopCheck -eq 30)
                    {
                    Write-host "Could not retrieve the Team " $BusinessPartnerCompanyName " in 60 seconds. Exiting script."
                    Exit
                    }
                
            } Until ($Team -ne $null)
           
            Write-Host "Adding the Team " $BusinessPartnerCompanyName " to Access Package catalog " $CurrentAccessPackageCatalog.displayname
            $fnOBJAddTeamToCatalog = AddTeamToAccessPackageCatalog($CurrentAccessPackageCatalog,$BusinessPartnerCompanyName, $Team,$AzureTenantID)
            CloudWait
            AddTeamtoAccessPackage {($DepartingUserAccessGroup,$Team)}

        }


    
    "ADhoc"
        {
            $DepartingUserAccessGroup = CreateSecurityGroup($BusinessPartnerCompanyName)
            Write-Host "Post security group creation of " $BusinessPartnerCompanyName
            Write-Host "Adding SharePoint site " $ADHocSharingSPOSite " and security group " $DepartingUserAccessGroup.displayname " to Access Package catalog " $CurrentAccessPackageCatalog.displayname
            $fnOBJAddResourceToCatalog = AddResourcesToAccessPackageCatalog($ADHocSharingSPOSite,$DepartingUserAccessGroup,$CurrentAccessPackageCatalog,$BusinessPartnerCompanyName)
            CloudWait

            Write-Host "Adding resources to Access Package " $BusinessPartnerCompanyName
            $fnOBJAddResourceToAccessPackage =  AddResourcestoAccessPackage($BusinessPartnerCompanyName,$AccessPackageCatalogID,$DepartingUserAccessGroup,$ADHocSharingFolder)
            CloudWait

            Write-Host "Creating and modifying permissions for folder" $BusinessPartnerCompanyName  " in SharePoint site " $ADHocSharingSPOSite
            CreateSharePointFolder($BusinessPartnerCompanyName,$ADHocSharingSPOSite)
            CloudWait

            Write-Host "Modifying permissions for folder" $BusinessPartnerCompanyName  " in SharePoint site " $ADHocSharingSPOSite
            FixSharePointFolderPermissions ($BusinessPartnerCompanyName,$DepartingUserAccessGroup)
            Disconnect-PnPOnline
            CloudWait

        }

    } 
Disconnect-AzAccount
$Authcode = $null
Write-Host "Script complete at " (get-date)
