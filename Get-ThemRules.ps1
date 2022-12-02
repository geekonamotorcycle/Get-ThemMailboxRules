<#
.SYNOPSIS
    This script will return a CSV with all of the rules found for all mailboxes in exchange online
    It will also output the results of a search for all mailboxes with forwarding enabled
.DESCRIPTION
    1.  checks for Exchangeonline session and starts a session if none is detected
    2.  intializes an empty $Array
    3.  Retrieves a list of mailbox UserPrincipalName
    4.  Iterates through the list of UserPrincipalName with the Get-InboxRule cmdlet
    5.  Loads reults into a PSCustomObject and adds to $Array
    6.  Outputs to a CSV in the same folder the Script was run from (will overwrite)
    Also
    1.  searches for SMTP forwarding
    2.  Outputs to a CSV in the same folder the Script was run from (will overwrite)

.NOTES
    Last reveiwed: 2022-12-02
    Version: 1.5
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
$MailboxList = get-mailbox | Select-Object UserPrincipalName;
Write-Host "there are " $MailboxList.count "mailbox to check" -ForegroundColor Cyan;

# Develop file and path names based on date and time
$Date = (Get-Date);
$DateStr = $Date.ToString("yyyyMMdd-HHmm");
$filename = $DateStr + "_MailboxRulesReport.csv";
$BasePath = Get-Location;
$FullPath = Join-Path -Path $BasePath -ChildPath $filename;

# set to True to spam your screen with Results (best for debugging)
$Verbose = $False;

# Initialize $Array
$Array = @();

# Set error action
$ErrorActionPreference = "silentlycontinue";

# begin rule lookup loop
for ($i = 0; $i -lt $MailboxList.Count; $i++) {
    
    $Row = Get-InboxRule -Mailbox $MailBoxlist[$i].UserPrincipalName.ToString() -IncludeHidden;
    Write-Host "User: " $MailBoxlist[$i].UserPrincipalName.ToString() "Number: " $i "/" $MailboxList.Count -ForegroundColor Green;
    for ($inner = 0; $inner -lt $Row.Count; $inner++) {
        Write-Host `t $Row[$inner].Identity " Rule #: " $inner "/" $Row.Count -ForegroundColor Yellow;
        if ($verbose) {
            Write-Host "Identity: " $Row[$inner].Identity
            Write-Host "IsValid: "$Row[$inner].IsValid
            Write-Host "Enabled: "$Row[$inner].Enabled
            Write-Host "Legacy: "$Row[$inner].Legacy
            Write-Host "RuleIdentity: "$Row[$inner].RuleIdentity
            Write-Host "MailboxOwnerId: "$Row[$inner].MailboxOwnerId
            Write-Host "Name: "$Row[$inner].Name
            Write-Host "From: "$Row[$inner].From
            Write-Host "SentTo: "$Row[$inner].SentTo
            Write-Host "HeaderContainsWords: "$Row[$inner].HeaderContainsWords
            Write-Host "SubjectContainsWords: "$Row[$inner].SubjectContainsWords
            Write-Host "BodyContainsWords: "$Row[$inner].BodyContainsWords
            Write-Host "SoftDeleteMessage: "$Row[$inner].SoftDeleteMessage
            Write-Host "DeleteMessage: "$Row[$inner].DeleteMessage
            Write-Host "ForwardTo: "$Row[$inner].ForwardTo
            Write-Host "ForwardAsAttachmentTo: "$Row[$inner].ForwardAsAttachmentTo
            Write-Host "MoveToFolder: "$Row[$inner].MoveToFolder
            Write-Host "RedirectTo: "$Row[$inner].RedirectTo
            Write-Host "Description: "$Row[$inner].Description`n
        };
        # Enter Blank spaces for Null Values
        
        if ($null -ne $Row[$inner].Identity) {}else { $Row[$inner].Identity = " " };
        if ($null -ne $Row[$inner].IsValid) {}else { $Row[$inner].IsValid = " " };
        if ($null -ne $Row[$inner].Enabled) {}else { $Row[$inner].Enabled = " " };
        if ($null -ne $Row[$inner].Legacy) {}else { $Row[$inner].Legacy = " " };
        if ($null -ne $Row[$inner].RuleIdentity) {}else { $Row[$inner].RuleIdentity = " " };
        if ($null -ne $Row[$inner].MailboxOwnerId) {}else { $Row[$inner].MailboxOwnerId = " " };
        if ($null -ne $Row[$inner].From) {}else { $Row[$inner].From = " " };
        if ($null -ne $Row[$inner].SentTo) {}else { $Row[$inner].SentTo = " " };
        if ($null -ne $Row[$inner].HeaderContainsWords) {}else { $Row[$inner].HeaderContainsWords = " " };
        if ($null -ne $Row[$inner].SubjectContainsWords) {}else { $Row[$inner].SubjectContainsWords = " " };
        if ($null -ne $Row[$inner].BodyContainsWords) {}else { $Row[$inner].BodyContainsWords = " " };
        if ($null -ne $Row[$inner].SoftDeleteMessage) {}else { $Row[$inner].SoftDeleteMessage = " " };
        if ($null -ne $Row[$inner].DeleteMessage) {}else { $Row[$inner].DeleteMessage = " " };
        if ($null -ne $Row[$inner].ForwardTo) {}else { $Row[$inner].ForwardTo = " " };
        if ($null -ne $Row[$inner].ForwardAsAttachmentTo) {}else { $Row[$inner].ForwardAsAttachmentTo = " " };
        if ($null -ne $Row[$inner].MoveToFolder) {}else { $Row[$inner].MoveToFolder = " " };
        if ($null -ne $Row[$inner].RedirectTo) {}else { $Row[$inner].RedirectTo = " " };
        if ($null -ne $Row[$inner].Description) {}else { $Row[$inner].Description = " " };
        
        $Array += [PSCustomObject]@{
            Identity              = $Row[$inner].Identity
            IsValid               = $Row[$inner].IsValid
            Enabled               = $Row[$inner].Enabled
            Legacy                = $Row[$inner].Legacy
            RuleIdentity          = $Row[$inner].RuleIdentity
            RowOwnerId            = $Row[$inner].MailboxOwnerId
            Name                  = $Row[$inner].Name
            From                  = $Row[$inner].From
            SentTo                = $Row[$inner].SentTo
            HeaderContainsWords   = $Row[$inner].HeaderContainsWords
            SubjectContainsWords  = $Row[$inner].SubjectContainsWords
            BodyContainsWords     = $Row[$inner].BodyContainsWords
            SoftDeleteMessage     = $Row[$inner].SoftDeleteMessage
            DeleteMessage         = $Row[$inner].DeleteMessage
            RedirectTo            = $Row[$inner].RedirectTo
            ForwardTo             = $Row[$inner].ForwardTo
            ForwardAsAttachmentTo = $Row[$inner].ForwardAsAttachmentTo
            MoveToFolder          = $Row[$inner].MoveToFolder
            Description           = $Row[$inner].Description
        }
    }
        
};

# Output to CSV
$Array | Export-Csv -NoTypeInformation -Delimiter "," -Force -Path $fullpath;

# Forwarding Report
$ForwardingReport = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | Select-Object UserPrincipalName, ForwardingSmtpAddress, DeliverToMailboxAndForward;

# Develop file and path names based on date and time
$FwdRTDateStr = $Date.ToString("yyyyMMdd-HHmm");
$FwdRTfilename = $FwdRTDateStr + "_ForwardingReport.csv";
$FwdRTfullpath = Join-Path -Path $BasePath -ChildPath $FwdRTfilename;
$ForwardingReport | Export-Csv -NoTypeInformation -Force -Path $FwdRTfullpath;
 
# present run-time
$EndTime = (Get-Date);
$TimeBetween = $EndTime - $StartTime;
Write-Host "Run time was"`n`t "Hours: " $TimeBetween.TotalHours `n`t "Minutes: " $TimeBetween.TotalMinutes `n`t "Seconds: " $TimeBetween.TotalSeconds;
