#
# SearchAndRemoveItemsMailboxes.PS1
# https://github.com/12Knocksinna/Office365itpros/blob/master/SearchAndRemoveItemsMailboxes.PS1
# A script to use the Search-Mailbox cmdlet to remove calendar items from user mailboxes.
# Requires the Exchange Online management module

param (
  [Parameter(Mandatory,HelpMessage='Please define the meeting subject.')]
  [string]$Subject,
  [string]$StartDate = '1/1/2019',
  [string]$EndDate = '12/31/2022',
  [string]$CsvFileName = 'MailboxesToSearch.csv'
)

$scriptVersion = '1.0'

Clear-Host

$ModulesLoaded = Get-Module | Select-Object Name
#If (!($ModulesLoaded -match "ExchangeOnlineManagement")) {Write-Host "Please connect to the Exchange Online Management module and then restart the script" ; break}
# Add Check for Exchange PS

Write-Host ('Reading users from {0}' -f $CsvFileName)

# Set up the search query

#$StartDate = "1/1/2019"
# $EndDate = "12/31/2022"

$Subject = 'Die Serie'

$Query = "Received:$($StartDate)..$($EndDate) kind:meetings" + " AND (subject:""$($Subject)"")"
# $Query = "kind:meetings" + " AND (subject:""$($Subject)"")"

$CsvRemovedItemsFileName = ('Removed-Items-{0}' -f (Get-Date -Format 'yyyy-mm-HH') )

# Initalize some parameters
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name

if(Test-Path -Path (Join-Path -Path $ScriptDir -ChildPath $CsvFileName)) {

  # Find the mailboxes to process - this example uses a check against the custom attribute 12. You could also read in user details from a CSV file
  #$Users = Get-ExoMailbox -Filter {CustomAttribute12 -eq "Search"} -Properties CustomAttribute12 -RecipientTypeDetails UserMailbox -ResultSize Unlimited
    
  # Import Users from CSV file
  $Users = Import-Csv -Path (Join-Path -Path $ScriptDir -ChildPath $CsvFileName) -Encoding UTF8 

  If (!$Users) {
    Write-Host 'No matching users found - exiting' 
    break
  }
  else {
    Write-Host ('{0} users to check' -f ($Users | Measure-Object).Count )
  }

  $UserReport = [System.Collections.Generic.List[Object]]::new() # Create output file 

  ForEach ($User in $Users)  {
        
    $UserDetails = Get-User $User.UserPrincipalName

    $Status = (Search-Mailbox -Identity $User.UserPrincipalName -SearchQuery $Query -EstimateResultOnly -DoNotIncludeArchive -SearchDumpster:$False)
        
    If ($Status) {
      $ReportLine = [PSCustomObject] @{
        UserPrincipalName  = $User.UserPrincipalName
        DisplayName        = $UserDetails.DisplayName
        ItemsFound         = $Status.ResultItemsCount
        ItemsSize          = $Status.ResultItemsSize
        SearchType         = 'Estimate'
      SearchTime         = Get-Date }                
    $UserReport.Add($ReportLine) } #End if
  } # End For

  # Filter the users where we have found some items
  $ProcessUsers = $UserReport | Where-Object {$_.ItemsFound -ne '0'}
  Clear-Host

  if(($ProcessUsers | Measure-Object).Count -ne 0) {
    
    Clear-Host
    $ProcessUsers | Format-Table -Property DisplayName, UserPrincipalName, ItemsFound, ItemsSize -AutoSize

    # preset some variables
    $PromptTitle = 'Remove items from mailboxes'
    $PromptMessage = 'Please confirm whether to proceed to remove found items from mailboxes'

    $yes = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&yes', 'yes?'
    $no = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&no', 'no?'
    $cancel = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription -ArgumentList '&cancel', 'Exit'
    $PromptOptions = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no, $cancel)
    $PromptDecision = $host.ui.PromptForChoice($PromptTitle, $PromptMessage, $PromptOptions, 0) 

    $i = 0

    Switch ($PromptDecision) {
      '0' { 
        ForEach ($User in $ProcessUsers) {

          Write-Host ('Removing items from the mailbox of {0}' -f $User.DisplayName)
          $Status = (Search-Mailbox -Identity $User.UserPrincipalName -SearchQuery $Query -DeleteContent -DoNotIncludeArchive -SearchDumpster:$False -Confirm:$False -Force)

          If ($Status) {  # Add record to capture what we did
            Write-Host ('Mailbox for {0} processed to remove {1} items' -f  $User.DisplayName, $Status.ResultItemsCount)
            $ReportLine = [PSCustomObject] @{
              UserPrincipalName  = $User.UserPrincipalName
              DisplayName        = $User.DisplayName
              ItemsFound         = $Status.ResultItemsCount
              ItemsSize          = $Status.ResultItemsSize
              SearchType         = 'Removal'
            SearchTime         = Get-Date }                 
                 
            $ProcessUsers += $ReportLine
            $i++
          }
        } #End ForEach

        Write-Host ('All done.  {0} mailboxes processed and cleaned up. Details stored in {1}' -f $i, (Join-Path -Path $ScriptDir -ChildPath $CsvRemovedItemsFileName ))
              
        # Export users to CSV file
        $ProcessUsers | Export-Csv -Path (Join-Path -Path $ScriptDir -ChildPath $CsvRemovedItemsFileName) -NoTypeInformation -Encoding UTF8 -Force

      }
      '1' {
        Write-Host 'OK. Maybe later? Messages not removed from mailboxes.'
      }
      '2' {
        Write-Host 'Cancelled. Messages not removed from mailboxes.'
      }
    } #End Switch
  }
  else {
    Write-Host 'No users found'
  }

}