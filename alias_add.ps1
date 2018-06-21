# Google to AD Email Alias Import
# Leon Topliss - ITV Plc


# Is this a dry run - If true we won't make any changes just log as if we were.
# Set this to "no" if you want to apply or "yes" to log the changes that would be made in the debug log
$dryRun = "yes"

# The variables below are:
# $GoogleExportCsvFile = CSV file containing the GAM G-Suite extract
# $DebugLogFile = The debug log output.
# $ErrorLogFile = A seperare log of the errors, they are also included in the debug output 
# All files should be in the same directory as the script

$GoogleExportCsvFile = "allusers.csv" 
$DebugLogFile = "debug-output.txt" 
$ErrorLogFile = "error-output.txt" 

# How long in milliseconds to sleep between adding users
# The aim is to avoid hammering AD
$sleepBetweenUsers = 500


$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Setup the logging function
# If we see an error log it to the console and a seperate error file
Function Log
{
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [string]$level,
         [Parameter(Mandatory=$true, Position=1)]
         [string]$logstring
    )
    
    $date = Get-Date -Format g
    $full_logstring = $date + " - " + $level + " - " + $logstring
    $DebugLogFileFullPath = $scriptDir + "\" + $DebugLogFile
    $ErrorLogFileFullPath = $scriptDir + "\" + $ErrorLogFile
    Add-content $DebugLogFileFullPath -value $full_logstring
    if($level -match "ERROR") {
        Add-content $ErrorLogFileFullPath -value $full_logstring 
        Write-Output  $full_logstring 
    }
}

$csvFile = $scriptDir + "\" + $GoogleExportCsvFile


# Read the CSV File
# Powershell is pretty good at CSV parsing so you can use the headers at the top
# as keys for the data
$csv = Import-Csv $csvFile

$index = 0
$total = $csv.count

foreach ($line in $csv) {

    $index++

    # The Google extract has "PrimaryEmail" as he header for Primary Email
    $primaryEmail = $line."PrimaryEmail"

    Write-Progress `
        -Activity "Email Alias Adds" `
        -Status "$index of $total [$('{0:N2}' -f (($index/$total)*100))%]" `
        -CurrentOperation "USER: $primaryEmail" `
        -PercentComplete (($index/$total)*100)
    
    # Check the primary email is in the right format
    if (-Not [bool]($primary_email -as [Net.Mail.MailAddress])) {
        Log "ERROR" "primary email: $primary_email is not a valid email format"
        continue
    }

    $aliasArray = @()
    # Look at each header if it contains aliases.*
    # Take the value it's an alias we will want to impliment
    foreach ($property in $line.PSObject.Properties) {
        if($property.Name -match "aliases.*") {
            $alias = $property.Value
            # If the value isn't empty use it
            if ($alias) {
                # Check it's a valid email
                if (-Not [bool]($property.Value -as [Net.Mail.MailAddress])) {
                    Log "ERROR" "alias email: $alias recorded against $primaryEmail is not a valid email format"
                    continue
                } else {
                    # Its in a column with a header aliases.*, the cell isn't empty and it's a valid email
                    # So add it to the aliases array for this row.
                    $aliasArray +=  $alias
                }                
            }
        }
    }

    # At this point we have a 
    # primary email $primaryEmail 
    # and 
    # an array of aliases $aliasArray
    # for the individual row of the CSV


    # Set a sleep interval between users to minimise load
    Start-Sleep -m $sleepBetweenUsers

    # Retrieve the object we would like to add the alias to using primary email as the key
    $user = Get-ADObject -Properties mail -Filter {mail -eq $primaryEmail}

    # If the user primary email doesn't exist in then skip.
    if (-Not $user) {
        Log "ERROR" "Could not find the user $primaryEmail"
        continue
    }


    # Now loop through each alias and add it to the AD account
    Foreach ($alias in $aliasArray) {
            
        # Prepend SMTP: to the alias
        $aliasAttribute = "SMTP:$($alias)"
            

        # Check the email alias isn't set elsewhere
        # If we find the alias is set elswhere just report an error.
        $aliasExists = Get-ADObject -Properties proxyAddresses -Filter {proxyAddresses -eq $aliasAttribute}
        if ($aliasExists) {
            if ($aliasExists –Match $user) {
                Log "INFO" "THe alias $alias is already set on the user $aliasExists"
            } else {
                Log "ERROR" "The email alias $alias is already set on object $aliasExists"
            }
            continue
        }

        # The user exists and the alias isn't set elsewhere (if we have made it this far)
        # So now add the alias
        if ($dryRun -Match "No") {
            $user | Set-ADUser -Add @{ProxyAddresses=$aliasAttribute}
            Log "INFO" "Added the alias $alias  to user $user"
        } else {
            Log "INFO" "DRY RUN - But would have added the alias $alias  to user $user"
        }
    }
}






