<#       
Remove a specified file from OneDrive for Business Sites.

    Adapted from https://gallery.technet.microsoft.com/Remove-a-OneDrive-for-c6fd3c30
    This script depends on PnP Powershell. You can download it here: https://www.powershellgallery.com/packages/SharePointPnPPowerShellOnline/3.29.2101.0

    .DESCRIPTION
    Removes file from OneDrive for Business sites.

    .PARAMETER Credential
    Specify SharePoint or Global Administrator credential.

    .PARAMETER Tenant
    Specify Office 365 tenant name.  For example, if the Sharepoint
    Online site is 'contoso.onmicrosoft.com', 'contoso' is the tenant
    name.

    .PARAMETER InputFile
    Supply a CSV with the header 'userprincipalname' to process a subset
    of users. If a CSV is not supplied, all OneDrive sites are enumerated
    and processed.

    .PARAMETER GrantPermission
    Grants Site Collection Administrator permission to the administrator
    account specified in the -Username parameter. Recommended to set
    this switch if it is unknown if the admin user already has Site
    Collection Administrator permissions. If granted, it must be removed
    separately at this time.

    .PARAMETER FileToDelete
    Specify name of File in OneDrive for Business site to delete.

    .PARAMETER Confirm
    Confirm deletion of File specified in -FileToDelete parameter.

    .EXAMPLE
    $SpoAdmin = Get-Credential
    .\Remove-OneDriveForBusinessFile.ps1 -Credential $SpoAdmin -Tenant contoso -FileToDelete 'test.exe' -GrantPermission -LogFile contoso_od4b.log

    Enumerate all OneDrive for Business Files for tenant Contoso and save output to contoso_od4b.log.

    #>
[Cmdletbinding()]
    Param (
        [Parameter(Mandatory=$true)]
            [System.Management.Automation.PSCredential]$Credential,

        [Parameter(mandatory=$true)]
            [String]$Tenant,

        [Parameter(mandatory=$false)]
        [ValidateScript({Test-Path $_;if (!((gc $_ | Select-Object -First 1) -like "*UserPrincipalName*")) 
            { Write-Host -Fore Red "Please make sure input CSV only contains header UserPrincipalName.";Break }})]
            [String]$InputFile,

        [Parameter(mandatory=$false)]
        [Switch]$GrantPermissions,

        [Parameter(mandatory=$false)]
        [String]$LogFile,

        [Parameter(Mandatory=$true,HelpMessage='File to delete')]
        [string]$FileToDelete,

        [Parameter(Mandatory=$false,HelpMessage='Confirm removal of the Files')]
        [Bool]$Confirm
        )

    begin 
        {        
        }
    process 
        {
        # Locating Sharepoint Server Client Components and loading
        #Install-Module -Name SharePointPnPPowerShellOnline
        Write-Host -Fore Yellow "Locating SharePoint Server Client Components installation..."
        If (Test-Path 'c:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll')
            {
                Write-Host -ForegroundColor Green "Found SharePoint Server Client Components installation."
                
                Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll" 
                Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll" 
                Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"
                Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll"
                
            }
        ElseIf
            ( $filename = Get-ChildItem 'C:\Program Files' -Recurse -ea silentlycontinue | where { $_.name -eq 'Microsoft.SharePoint.Client.UserProfiles.dll' })
            {
                $Directory = $filename.DirectoryName
                Write-Host -ForegroundColor Green "Found SharePoint Server Client Components at $Directory."
                Add-Type -Path "$Directory\Microsoft.SharePoint.Client.dll" 
                Add-Type -Path "$Directory\Microsoft.SharePoint.Client.Runtime.dll" 
                Add-Type -Path "$Directory\Microsoft.SharePoint.Client.Taxonomy.dll"
                Add-Type -Path "$Directory\Microsoft.SharePoint.Client.UserProfiles.dll"
            }

        # Create Log File
        If (!(Test-Path $LogFile))
            {
            Write-Host -ForegroundColor Yellow "Log file not found. Creating."
            $LogFileHeader = "File" + "," + "User" + "," + "Path" + "," + "Action Taken" 
            $LogFileHeader | Out-File $LogFile
            }
        Else
            {
            Write-Host -ForegroundColor Yellow "Existing log file found. Appending."
            }

        # Define URLs
        $MySiteURL = "https://$tenant-my.sharepoint.com"
        $AdminURL = "https://$tenant-admin.sharepoint.com"

        # Define Contexts
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($MySiteURL)

        # Define Credentials
        $Creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username,$Credential.Password)
        $Context.Credentials = $Creds

        # Connect to SPO Service for granting permissions if necessary
        $SPOServiceCreds = $Credential
        Connect-SPOService -url $AdminURL -credential $SPOServiceCreds

        # Get OD4B WebSite Users
        $Users = $Context.Web.SiteUsers
        $Context.Load($Users)
        $Context.ExecuteQuery()
        $peopleManager = New-Object Microsoft.SharePoint.Client.UserProfiles.PeopleManager($Context)

        # Check to see if there is an input file specified
        If ($InputFile)
            {
            $UserList = Import-Csv $InputFile -Header UserPrincipalName
            $i = 0
            foreach ($User in $Users)
                    {
                        $userProfile = $peopleManager.GetPropertiesFor($user.LoginName)
                        $Context.Load($userProfile)
                        $Context.ExecuteQuery()
                        if ($userList.UserPrincipalname -contains $userProfile.Email -and $userProfile.UserProfileProperties.PersonalSpace -ne "" )
                            {
                            $i++
                            $OD4BPath = $MySiteURL+$userProfile.UserProfileProperties.PersonalSpace
                            Write-Host $OD4BPath # $userProfile.UserProfileProperties.PersonalSpace
                            If ($GrantPermissions) 
                                {
                                Write-Host -ForegroundColor Green "Granting permissions on $OD4BPath"
                                Set-SPOUser -Site $OD4BPath -LoginName $Credential.Username -IsSiteCollectionAdmin $true | Out-Null
                                }

                            # If OD4BPath is present, enumerate Files
                            if ($OD4BPath) 
                                {
                                #Write-Host -ForegroundColor Cyan "User $($userprofile.Email) has a OneDrive for Business Site."
                                $ClientContextSource = New-Object Microsoft.SharePoint.Client.ClientContext($OD4BPath);
                                Write-Host "     URL is $OD4BPath";

                                $ClientContextSource.Credentials = $Creds
                                $ClientContextSource.ExecuteQuery()

                                $personalWeb = $ClientContextSource.Web
                                $ClientContextSource.Load($personalWeb)
                                $ClientContextSource.ExecuteQuery()

                                $docList = $personalWeb.Lists.GetByTitle("Documents")
                                $ClientContextSource.Load($docList)
                              
                                $ClientContextSource.ExecuteQuery()

                                $allFiles = $docList.RootFolder.Files
                                $ClientContextSource.Load($allFiles)
                                $ClientContextSource.ExecuteQuery()

                                # Delete Specified File 
                                foreach ($toBeDeleted in $allFiles)
                                            {
                                                #Write-Host "Examining File" $toBeDeleted.Name
                                                if ($toBeDeleted.Name -eq $FileToDelete)
                                                {
                                                    Write-Host -Fore Green "     $($FileToDelete) present in $OD4BPath"
                                                    If ($Confirm)
                                                        { 
                                                        Write-Host -ForegroundColor Cyan "     ** Confirm enabled. Deleting File " $toBeDeleted.Name
                                                        $toBeDeleted.Recycle()
                                                        $personalWeb.Update()
                                                        $ClientContextSource.ExecuteQuery()
                                                        If ($LogFile)
                                                            {
                                                            $logEntry = $FileToDelete + "," + $userprofile.email + "," + $OD4BPath + "," + "File Deleted."
                                                            $logEntry | Out-File $LogFile -Append -Force
                                                            $logEntry = $null
                                                            }
                                                        }
                                                    Else
                                                        { 
                                                        Write-Host -ForegroundColor Yellow "No action taken."
                                                        $logEntry = $FileToDelete + "," + $userprofile.email + "," + $OD4BPath + "," + "No action taken."
                                                        $logEntry | Out-File $LogFile -Append -Force
                                                        $logEntry = $null
                                                        }
                                                }
                                            }
                                }
                            }
                    }
        Write-Host "Matching Onedrive for Business sites:"$i
            }
        Else
            {
            $i = 0

            foreach ($User in $Users)
                    {
                        $userProfile = $peopleManager.GetPropertiesFor($user.LoginName)
                        $Context.Load($userProfile)
                        $Context.ExecuteQuery()
                        if ($userProfile.Email -ne $null -and $userProfile.UserProfileProperties.PersonalSpace -ne "" )
                            {
                            $i++
                            $OD4BPath = $MySiteURL+$userProfile.UserProfileProperties.PersonalSpace
                            Write-Host $OD4BPath # $userProfile.UserProfileProperties.PersonalSpace
                            If ($GrantPermissions) 
                                {
                                Write-Host -ForegroundColor Green "Granting permissions on $OD4BPath"
                                Set-SPOUser -Site $OD4BPath -LoginName $Credential.Username -IsSiteCollectionAdmin $true | Out-Null
                                }

                            # If OD4BPath is present, enumerate Files
                            if ($OD4BPath) 
                                {
                                #Write-Host -ForegroundColor Cyan "User $($userprofile.Email) has a OneDrive for Business Site."
                                $ClientContextSource = New-Object Microsoft.SharePoint.Client.ClientContext($OD4BPath);
                                Write-Host "     URL is $OD4BPath";

                                $ClientContextSource.Credentials = $Creds
                                $ClientContextSource.ExecuteQuery()

                                $personalWeb = $ClientContextSource.Web
                                $ClientContextSource.Load($personalWeb)
                                $ClientContextSource.ExecuteQuery()

                                $docList = $personalWeb.Lists.GetByTitle("Documents")
                                $ClientContextSource.Load($docList)
                                $ClientContextSource.ExecuteQuery()


                                $allFiles = $docList.RootFolder.Files
                                $ClientContextSource.Load($allFiles)
                                $ClientContextSource.ExecuteQuery()
                               
                                
                                # Delete Specified File
                                foreach ($toBeDeleted in $allFiles)
                                            {
                                                #Write-Host "Examining File" $toBeDeleted.Name
                                                if ($toBeDeleted.Name -eq $FileToDelete)
                                                {
                                                    Write-Host -Fore Green "     $($FileToDelete) present in $OD4BPath"
                                                    If ($Confirm)
                                                        { 
                                                        Write-Host -ForegroundColor Cyan "     ** Confirm enabled. Deleting file  " $toBeDeleted.Name
                                                        $toBeDeleted.Recycle()
                                                        #$personalWeb.Update()
                                                        $ClientContextSource.ExecuteQuery()
                                                        If ($LogFile)
                                                            {
                                                            $logEntry = $FileToDelete +  "," + $userprofile.email + "," +  $OD4BPath + "," + "File Deleted." 
                                                            $logEntry | Out-File $LogFile -Append -Force
                                                            $logEntry = $null
                                                            }
                                                        }
                                                    Else
                                                        { 
                                                        Write-Host -ForegroundColor Yellow "No action taken."
                                                        $logEntry = $FileToDelete + "," + $userprofile.email + "," + $OD4BPath + "," +  "No action taken." 
                                                        $logEntry | Out-File $LogFile -Append -Force
                                                        $logEntry = $null
                                                        }
                                                }
                                            }
                                }
        #####
        #Write-Host "Matching Onedrive for Business sites:"$i
                }
            }

        }
    }

    