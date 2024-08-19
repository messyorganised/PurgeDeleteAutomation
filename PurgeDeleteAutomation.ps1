##### Check if the ExchangeOnlineManagement module is installed #####

if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "The required ExchangeOnlineManagement module is not installed. Installing now..."
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
    Write-Host "ExchangeOnlineManagement module installed successfully."
} else {
    Write-Host "The required ExchangeOnlineManagement module is already installed."
}

##### End #####



function NewEmailSearch {
    #Search Input
    do {
        $EmailToDelete = Read-Host -Prompt "Enter the email title or sender to delete"
    } while ([string]::IsNullOrWhiteSpace($EmailToDelete))

    $SearchName="PSL"
    New-ComplianceSearch -Name $SearchName -ExchangeLocation All -ContentMatchQuery $EmailToDelete
    Start-ComplianceSearch -Identity $SearchName
    Write-Host "Starting Search..." "`n" 

    do {
        Start-Sleep -Seconds 5
        $searchStatus = Get-ComplianceSearch -Identity $SearchName
    } while ($searchStatus.Status -ne "Completed")

    Write-Host "Search Completed" "`n" 
    Start-Sleep -Seconds 1

}

function SoftPurging{
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete
    Write-Host "Starting Soft Delete..." "`n" 

    # Wait for the SoftDelete action to complete
    do {
        Start-Sleep -Seconds 5
        $actionStatus = Get-ComplianceSearchAction -Identity "$($SearchName)_Purge"
    } while ($actionStatus.Status -ne "Completed")
    Write-Host "Soft Delete completed. Queried emails are now located in the delete folder." "`n" 
}

function HardPurging{
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete
    Write-Host "Starting Hard Delete..." "`n" 

    # Wait for the HardDelete action to complete
    do {
        Start-Sleep -Seconds 5
        $actionStatus = Get-ComplianceSearchAction -Identity "$($SearchName)_Purge"
    } while ($actionStatus.Status -ne "Completed")

    Write-Host "Hard Delet completed. Please beware that emails removed is not retrievable on client end and will be fully removed within 30 days." "`n" 

}

# Connect to Exchange Online
Connect-ExchangeOnline
Connect-IPPSSession

## END ## 

#Prompt the user for the email title or sender to delete

NewEmailSearch

Write-Host "
What would you like to do with the queried emails.

1= Soft Delete
2= Hard Delete

Press any other keys to end the set up

"
$codelink = Read-Host
switch ($codelink) {
    '1' { SoftPurging }
    '2' { HardPurging }
    default { Write-Host "No valid options selected. Closing App."   "`n" }
}

Start-Sleep -Seconds 1
Write-Host "Removing search query..." "`n" 
Remove-ComplianceSearch -Identity $SearchName 


Write-Host "Done! Have a nice Day!" "`n" 
# Disconnect from Exchange Online
Disconnect-ExchangeOnline




