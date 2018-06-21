# Google to AD Email Alias Import
# Leon Topliss - ITV Plc

# -csv	  		- the input GAM CSV file
# -commit 		- by default false, but needs to be specified as an option to write changes
#				  without commit its a dry run and the changed that would be made are logged
# -aggressive  	- if specified skips the delay in between users. Will place AD under additional load

param (
	[string]$csv,
	[switch]$commit = $false,
	[switch]$aggressive = $false
)
 
if(-not($csv)) { Throw "The GAM user extract CSV must be specified -csv" }


# Log File Output
# $DebugLogFile = The debug log output.
# $ErrorLogFile = A seperare log of the errors, they are also included in the debug output 
# Log files will be written to the same directory
$fileDate = $((get-date).ToString("yyyyMMdd-hhmmss"))
$DebugLogFile = "debug-output-" +  $fileDate + ".txt"
$ErrorLogFile = "error-output-" +  $fileDate + ".txt" 

# How long in milliseconds to sleep between adding users
# The aim is to avoid hammering AD
# If aggressive this is not applicable
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

# Read the CSV File
# Powershell is pretty good at CSV parsing so you can use the headers at the top
# as keys for the data

$csvParse = Import-Csv $csv

$index = 0
$total = $csvParse.count

foreach ($line in $csvParse) {

    $index++

    # The Google extract has "PrimaryEmail" as he header for Primary Email
    $primaryEmail = $line."PrimaryEmail"

    Write-Progress `
        -Activity "Email Alias Adds" `
        -Status "$index of $total [$('{0:N2}' -f (($index/$total)*100))%]" `
        -CurrentOperation "USER: $primaryEmail" `
        -PercentComplete (($index/$total)*100)
    
    # Check the primary email is in the right format
    if (-Not [bool]($primaryEmail -as [Net.Mail.MailAddress])) {
        Log "ERROR" "primary email: $primary_email is not a valid email format"
        continue
    }

    $aliasArray = @()
    # Look at each header if it matches ^aliases.*
	# We are matching line begins with as we should not import nonEditableAliases.*
	# as nonEditableAliases is set globally in G-Suite and not on a per user basis
    # Take the value it's an alias we will want to impliment
    foreach ($property in $line.PSObject.Properties) {
        if($property.Name -match "^aliases.*") {
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
	
    # If there aren't any aliases skip the ad lookup and continue to the
    # next user
    if ($aliasArray.count -eq 0) {
    	continue
    }

    # Set a sleep interval between users to minimise load
    if (-not $aggressive) {
	Start-Sleep -m $sleepBetweenUsers
    }
	
    # Retrieve the object we would like to add the alias to using primary email as the key
    $user = Get-ADUser -Properties mail -Filter {mail -eq $primaryEmail}

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
        $aliasExists = Get-ADUser -Properties proxyAddresses -Filter {proxyAddresses -eq $aliasAttribute}
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
        if ($commit) {
            $user | Set-ADUser -Add @{ProxyAddresses=$aliasAttribute}
            Log "INFO" "Added the alias $alias  to user $user"
        } else {
            Log "INFO" "DRY RUN - But would have added the alias $alias  to user $user"
        }
    }
}
