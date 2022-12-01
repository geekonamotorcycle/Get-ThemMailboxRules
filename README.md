#SYNOPSIS

    This script will return a CSV with all of the rules found for all mailboxes in exchange online

#DESCRIPTION

    1.  Initializes a connection to exchange online if not already connected
    2.  Initializes an empty $Array
    3.  Retrieves a list of mailbox UserPrincipalName
    4.  Iterates through the list of UserPrincipalName with the Get-InboxRule cmdlet
    5.  Loads reults into a PSCustomObject and adds to $Array
    6.  Outputs to a CSV in the same folder the Script was run from (will over write)
    
#NOTES

    * Last reveiwed: 2022-12-010
    * Version: 1
    * Written By: Joshua Porrata
    * Email: joshua.porrata@gmail.com
    * License: Not free for commercial use without permission
    * There is a Verbose value, if you like spam.
#LINK

    https://github.com/geekonamotorcycle/Get-ThemMailboxRules
    
#EXAMPLE

    Get-ThemMailboxRules.ps1
