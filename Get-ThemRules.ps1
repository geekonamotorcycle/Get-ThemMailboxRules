<#
.SYNOPSIS
    This script will return a CSV with all of the rules found for all mailboxes in exchange online
.DESCRIPTION
    Initializes an empty $Array
    Retrieves a list of mailbox UserPrincipalName
    Iterates through the list of UserPrincipalName with the Get-InboxRule cmdlet
    Loads reults into a PSCustomObject and adds to $Array
    Outputs to a CSV in the same folder the Script was run from (will over write)
.NOTES
    Last reveiwed: 2022-12-010
    Version: 1
    Written By: Joshua Porrata
    Email: joshua.porrata@gmail.com
    License: Not free for commercial use without permission
    There is a Verbose value, if you like spam.
.LINK
    https://github.com/geekonamotorcycle/Get-ThemMailboxRules
.EXAMPLE
    Get-ThemMailboxRules.ps1
#>
    
# Set error action
$StartTime = (Get-Date);
$ErrorActionPreference = "prompt";

#Check Exchange online connection
if (!(
        Get-PSSession | Where-Object { 
            $_.Name -match 'ExchangeOnline' -and $_.Availability -eq 'Available'
        }
    )
) { 
    Connect-ExchangeOnline;
};

# Get Mailbox UPN list
$mailboxes = get-mailbox | Select-Object UserPrincipalName;
$Results = foreach ($box in $mailboxes) { Get-InboxRule -mailbox $box.UserPrincipalName.ToString() };
Write-Host "there are " $Results.count "Rules Found"
# Develop file and path names based on date and time
$Date = Get-Date;
$DateStr = $Date.ToString("yyyyMMdd-HHmm");
$filename = $DateStr + "_MailboxRulesReport.csv";
$BasePath = Get-Location;
$fullpath = Join-Path -Path $BasePath -ChildPath $filename;

# set to True to spam your screen with Results (best for debugging)
$verbose = $False;

# Initialize $Array
$Array = @();

# Set error action
$ErrorActionPreference = "silentlycontinue";

# begin rule lookup loop
foreach ($MailBox in $Results) {
    
    if ($verbose) {
        Write-Host "Identity: " $MailBox.Identity
        Write-Host "IsValid: "$MailBox.IsValid
        Write-Host "Enabled: "$MailBox.Enabled
        Write-Host "Legacy: "$MailBox.Legacy
        Write-Host "RuleIdentity: "$MailBox.RuleIdentity
        Write-Host "MailboxOwnerId: "$MailBox.MailboxOwnerId
        Write-Host "Name: "$MailBox.Name
        Write-Host "From: "$MailBox.From
        Write-Host "SentTo: "$MailBox.SentTo
        Write-Host "HeaderContainsWords: "$MailBox.HeaderContainsWords
        Write-Host "SubjectContainsWords: "$MailBox.SubjectContainsWords
        Write-Host "BodyContainsWords: "$MailBox.BodyContainsWords
        Write-Host "SoftDeleteMessage: "$MailBox.SoftDeleteMessage
        Write-Host "DeleteMessage: "$MailBox.DeleteMessage
        Write-Host "ForwardTo: "$MailBox.ForwardTo
        Write-Host "ForwardAsAttachmentTo: "$MailBox.ForwardAsAttachmentTo
        Write-Host "MoveToFolder: "$MailBox.MoveToFolder
        Write-Host "RedirectTo: "$MailBox.RedirectTo
        Write-Host "Description: "$MailBox.Description
    };
    
    #some logic that doesnt really work

    if ($null -ne $MailBox.Identity) {}else { $MailBox.Identity = " " };
    if ($null -ne $MailBox.IsValid) {}else { $MailBox.IsValid = " " };
    if ($null -ne $MailBox.Enabled) {}else { $MailBox.Enabled = " " };
    if ($null -ne $MailBox.Legacy) {}else { $MailBox.Legacy = " " };
    if ($null -ne $MailBox.RuleIdentity) {}else { $MailBox.RuleIdentity = " " };
    if ($null -ne $MailBox.MailboxOwnerId) {}else { $MailBox.MailboxOwnerId = " " };
    if ($null -ne $MailBox.From) {}else { $MailBox.From = " " };
    if ($null -ne $MailBox.SentTo) {}else { $MailBox.SentTo = " " };
    if ($null -ne $MailBox.HeaderContainsWords) {}else { $MailBox.HeaderContainsWords = " " };
    if ($null -ne $MailBox.SubjectContainsWords) {}else { $MailBox.SubjectContainsWords = " " };
    if ($null -ne $MailBox.BodyContainsWords) {}else { $MailBox.BodyContainsWords = " " };
    if ($null -ne $MailBox.SoftDeleteMessage) {}else { $MailBox.SoftDeleteMessage = " " };
    if ($null -ne $MailBox.DeleteMessage) {}else { $MailBox.DeleteMessage = " " };
    if ($null -ne $MailBox.ForwardTo) {}else { $MailBox.ForwardTo = " " };
    if ($null -ne $MailBox.ForwardAsAttachmentTo) {}else { $MailBox.ForwardAsAttachmentTo = " " };
    if ($null -ne $MailBox.MoveToFolder) {}else { $MailBox.MoveToFolder = " " };
    if ($null -ne $MailBox.RedirectTo) {}else { $MailBox.RedirectTo = " " };
    if ($null -ne $MailBox.Description) {}else { $MailBox.Description = " " };

    $Array += [PSCustomObject]@{
        Identity              = $MailBox.Identity
        IsValid               = $MailBox.IsValid
        Enabled               = $MailBox.Enabled
        Legacy                = $MailBox.Legacy
        RuleIdentity          = $MailBox.RuleIdentity
        MailboxOwnerId        = $MailBox.MailboxOwnerId
        Name                  = $MailBox.Name
        From                  = $MailBox.From
        SentTo                = $MailBox.SentTo
        HeaderContainsWords   = $MailBox.HeaderContainsWords
        SubjectContainsWords  = $MailBox.SubjectContainsWords
        BodyContainsWords     = $MailBox.BodyContainsWords
        SoftDeleteMessage     = $MailBox.SoftDeleteMessage
        DeleteMessage         = $MailBox.DeleteMessage
        RedirectTo            = $MailBox.RedirectTo
        ForwardTo             = $MailBox.ForwardTo
        ForwardAsAttachmentTo = $MailBox.ForwardAsAttachmentTo
        MoveToFolder          = $MailBox.MoveToFolder
        Description           = $MailBox.Description
    }
};

#output to CSV
$array | Export-Csv -NoTypeInformation -Delimiter "," -Force -Path $fullpath;
$EndTime = (Get-Date);
$TimeBetween = $EndTime - $StartTime;

# present run-time
Write-Host "Run time was"`n`t "Hours: " $TimeBetween.TotalHours `n`t "Minutes: " $TimeBetween.TotalMinutes `n`t "Seconds: " $TimeBetween.TotalSeconds;
