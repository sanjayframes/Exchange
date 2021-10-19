#############################################################################################################################################################################
#############################################################################################################################################################################
# 
# Author:   Sanjay Kumar Pasupuleti (sanjayframes@outlook.com)                  
# Version:  1.0  
# Last Modified Date:    06/11/2019 
#
#############################################################################################################################################################################
#############################################################################################################################################################################


param
(
    [Parameter(Mandatory=$false)]
    [switch]$FullAccess,

    [parameter(Mandatory=$false)]
    [switch]$SendAs,

     [parameter(Mandatory=$FALSE)]
    [switch]$SendOnBehalfTo,

      [parameter(Mandatory=$FALSE)]
    [switch]$Calendar

);

Write-Host "
    To run the script use the following switches

    -FullAccess = To Export Full Mailbox Access
    -SendAs = To Export SendAs access on Mailbox 
    -SendOnBehalfTo = To Export SendOnBehalfTo permissions on mailbox level
    -Calendar = To export calendar folder permissions (including Delegates)

    Example:
    " -NoNewline -ForegroundColor Yellow; Write-Host "
    .\Mailbox-Permissions.ps1 -FullAccess -SendAs -SendOnBehalfTo -Calendar


" -ForegroundColor Green


#Creating Directory
        Write-Host "`n`tCreating Directories for reports: " -NoNewline
        $null = mkdir ".\Permissions" -ErrorAction SilentlyContinue
        
        Write-Host "COMPLETED" -ForegroundColor Green
#Functions

$ErrorActionPreference = “silentlycontinue”

##### get time function
Function get-time {Get-Date -Format "hh:mm:ss tt"}


Write-Host "`nGetting Mailbox Count...@ " -NoNewline; Write-Host "$(get-time)" -ForegroundColor Green
$mailboxes = Get-Mailbox -ResultSize unlimited |?{$_.RecipientTypeDetails -ne "DiscoveryMailbox"} | Sort-Object -Property UserPrincipalName

       

if($FullAccess.IsPresent)
{


        Write-Host "`n`tExporting Mailbox FullAccess Info....@ " -NoNewline; Write-Host  "$(get-time)" -ForegroundColor Green
        if(Test-Path $FullPermissionsFileName){Write-Host "File already exist"}else{
        $fullpermissionsHeader = "Mailbox,User,AccessRights"
        $FullPermissionsFileName = ".\Permissions\FullAccess-Permissions.txt"
        Out-File -FilePath $FullPermissionsFileName -InputObject $fullpermissionsHeader -Append}


        Write-Host "`nTotal Mailboxes Found: " -NoNewline;Write-Host "$($mailboxes.count)" -ForegroundColor Magenta
        $Count = $mailboxes.count
     

            Write-Host "`nYou need the " -NoNewline; Write-Host " 'Startnumber' "-ForegroundColor Cyan -NoNewline;Write-Host "and " -NoNewline;Write-Host " 'Endnumber' "-ForegroundColor Cyan -NoNewline; Write-Host " to split the accounts"
            Write-Host "################## FOR EXAMPLE ##############################"
            Write-Host "`nIf you want to run for first 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '999'`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for second 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '100' and 'endnumber' as '199' and so on...`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for ALL USERS" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '$($count)' ...`n" -ForegroundColor Yellow
            Write-Host "#############################################################`n"
            

            $startnumber = Read-Host "Enter Startnumber"
            $endnumber = Read-Host "Enter Lastnumber"

            Write-Host "`n"

            $newcount = $mailboxes[$startnumber..$endnumber]
             
            $counter = 1
            $counting = $newcount.count

            ForEach ($mailbox in $newcount)
            {
                
               $pct = [int](($counter/$counting)*100)
               Write-Progress -Activity "Exporting Full Access Perms for " -Status " $($mailbox.PrimarySmtpAddress) ($counter of $counting)" -PercentComplete $pct

                
                $FullAccessPermissions = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress |?{($_.IsInherited -eq $false) -and ($_.AccessRights -match "FullAccess") -and ($_.user -notlike "NT AUTHORITY\*") -and ($_.User -notlike "*\Domain Admins") -and ($_.user -notlike "*\Exchange Servers") -and ($_.User -notlike "*\Exchange Trusted Subsystem") -and ($_.user -notlike "*\Enterprise Admins") -and ($_.user -notlike "*\Administrator") -and ($_.user -notlike "*\Organization management") -and ($_.user -ne "Discovery Management") -and ($_.User -notlike "*\Exchange Organization Administrators") -and ($_.User -notlike "S-1-5*")}
                Start-Sleep -Seconds 1
                
               
               
               
                            ForEach ($full in $FullAccessPermissions)
                                {
                                    
                                    Add-Content -Value ($mailbox.PrimarySmtpAddress.ToString() + "," + $full.User.ToString() + "," + $full.AccessRights) -Path $FullPermissionsFileName

                                }
                        
             
                        
                        $counter++
            }
            
Write-Host "`t`tExported to $($FullPermissionsFileName) ...@ " -NoNewline; Write-Host "$(get-time)" -ForegroundColor Green
}

if($SendAs.IsPresent)
{
        Write-Host "`n`tExporting Mailbox SendAs Info....@ " -NoNewline; Write-Host  "$(get-time)" -ForegroundColor Green

        $SendpermissionsHeader = "Identity,User,AccessRights,IsInherited"
        $SendPermissionsFileName = ".\Permissions\SendAs-Permissions.txt"
        Out-File -FilePath $SendPermissionsFileName -InputObject $SendpermissionsHeader -Append


        Write-Host "`nTotal Mailboxes Found: " -NoNewline;Write-Host "$($mailboxes.count)" -ForegroundColor Magenta
        $Count = $mailboxes.count
       
            Write-Host "`nYou need the " -NoNewline; Write-Host " 'Startnumber' "-ForegroundColor Cyan -NoNewline;Write-Host "and " -NoNewline;Write-Host " 'Endnumber' "-ForegroundColor Cyan -NoNewline; Write-Host " to split the accounts"
            Write-Host "################## FOR EXAMPLE ##############################"
            Write-Host "`nIf you want to run for first 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '99'`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for second 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '100' and 'endnumber' as '199' and so on...`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for ALL USERS" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '$($count)' ...`n" -ForegroundColor Yellow
            Write-Host "#############################################################`n"
            

            $startnumber = Read-Host "Enter Startnumber"
            $endnumber = Read-Host "Enter Lastnumber"

            Write-Host "`n"

            $newcount = $mailboxes[$startnumber..$endnumber]
             
            $counter = 1
            $counting = $newcount.count

            ForEach ($mailbox in $newcount)
            {

               $pct = [int](($counter/$counting)*100)
                Write-Progress -Activity "Exporting  SendAs Access for " -Status " $($mailbox.PrimarySmtpAddress) ($counter of $counting)" -PercentComplete $pct

                $SendAccessPerms = Get-RecipientPermission -Identity $mailbox.PrimarySmtpAddress  |?{$_.Trustee -ne "NT AUTHORITY\SELF"}
                Start-Sleep -Seconds 1
                
                if($SendAccessPerms)
                        {
                            ForEach ($sendAccess in $SendAccessPerms)
                                {
                                    Add-Content -Value ($mailbox.PrimarySmtpAddress.ToString() + "," + $sendAccess.Trustee + "," + $sendAccess.AccessRights + "," + $sendAccess.IsInherited) -Path $SendPermissionsFileName

                                }
                        
             
                        }
                        $counter++
            }

     Write-Host "`t`tExported to $($SendPermissionsFileName) ...@ " -NoNewline; Write-Host "$(get-time)" -ForegroundColor Green
}

if($SendOnBehalfTo.IsPresent)
{
    Write-Host "`n`tExporting 'SendOnBehalfTo' Information....@ " -NoNewline; Write-Host  "$(get-time)" -ForegroundColor Green

    $SendOnBehalfToReport = @()
    $SendOnBehalfUsers = $mailboxes |?{$_.GrantSendOnBehalfTo -ne $NULL}

    ForEach ($SendOnBehalf in $SendOnBehalfUsers){
        $mailbox = $null
        $mailbox = $SendOnBehalf.PrimarySmtpAddress

        ForEach($GrantUser in $SendOnBehalf.GrantSendOnBehalfTo){

        $SendOnBehalfID = $null
        $SendOnBehalfID = Get-Recipient $GrantUser -ErrorAction $ErrorActionPreference

        $SendOnBehalfProperties = @{

        Mailbox = $mailbox
        User = $SendOnBehalfID.PrimarySmtpAddress
        RecipientTypeDetails = $SendOnBehalfID.RecipientTypeDetails
        AccessRights = "GrantSendOnBehalfTo"
        }
        
        
        }
        $SendOnBehalfToReport += New-Object psobject -Property $SendOnBehalfProperties
    }

    

    $SendOnBehalfToReport | Select Mailbox, User, RecipientTypeDetails, AccessRights| Export-Csv ".\Permissions\SendOnBehalfTo-Permissions.csv" -NoTypeInformation
    Write-Host "`t`tExported to 'SendOnBehalfTo Permissions.csv' ...@ " -NoNewline; Write-Host "$(get-time)" -ForegroundColor Green

}

if($Calendar.IsPresent)
{
    Write-Host "`nTotal Mailboxes Found: " -NoNewline;Write-Host "$($mailboxes.count)" -ForegroundColor Magenta

    $Count = $mailboxes.count
    
    $CalpermissionsHeader = "Mailbox,UserEmail,UserRecipientTypeDetails,AccessRights,SharingPermissionFlags"
    $CalendarPermissionsFileName = ".\Permissions\CalendarPermissions.txt"
    Out-File -FilePath $calendarPermissionsFileName -InputObject $CalpermissionsHeader -Append

            Write-Host "`nYou need the " -NoNewline; Write-Host " 'Startnumber' "-ForegroundColor Cyan -NoNewline;Write-Host "and " -NoNewline;Write-Host " 'Endnumber' "-ForegroundColor Cyan -NoNewline; Write-Host " to split the accounts"
            Write-Host "################## FOR EXAMPLE ##############################"
            Write-Host "`nIf you want to run for first 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '99'`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for second 100 users" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '100' and 'endnumber' as '199' and so on...`n" -ForegroundColor Yellow
            Write-Host "`nIf you want to run for ALL USERS" -ForegroundColor Yellow
            Write-Host "Enter 'startnumber' as '0' and 'endnumber' as '$($count)' ...`n" -ForegroundColor Yellow
            Write-Host "#############################################################`n"
            

            $startnumber = Read-Host "Enter Startnumber"
            $endnumber = Read-Host "Enter Lastnumber"

            Write-Host "`n"

            $newcount = $mailboxes[$startnumber..$endnumber]
             
            $counter = 1
            $counting = $newcount.count

            ForEach ($mailbox in $newcount)
            {

               $pct = [int](($counter/$counting)*100)
                Write-Progress -Activity "Exporting Calendar Perms for " -Status " $($mailbox.PrimarySmtpAddress) ($counter of $counting)" -PercentComplete $pct

                $delegates = Get-MailboxFolderPermission -Identity ($mailbox.alias + ':\calendar') |?{$_.User -notmatch "Default" -and $_.user -notmatch "Anonymous"}
                Start-Sleep -Seconds 1
                if($delegates)
                        {
                            ForEach ($delegate in $delegates)
                                {
                                    $delegateuser = $null
                                    $delegateuser = Get-Recipient $delegate.User.ToString() -errorAction $ErrorActionPreference
                                    
                                    Add-Content -Value ($mailbox.PrimarySmtpAddress.ToString() + "," + $delegateuser.PrimarySmtpAddress + "," + $delegateuser.RecipientTypeDetails + "," + $delegate.AccessRights + "," + $delegate.SharingPermissionFlags) -Path $CalendarPermissionsFileName

                                }
                        
             
                        }
                        $counter++
            }

Write-Host "`t`tExported to $($CalendarPermissionsFileName) ...@ " -NoNewline; Write-Host "$(get-time)" -ForegroundColor Green

}



