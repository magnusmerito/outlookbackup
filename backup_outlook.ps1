################################################################################
# Outlook Mailbox Backup Script                                                #
#                                                                              #
# This PowerShell script helps you to backup mailboxes in your Outlook profile #
# to PST files. It lists all the mailboxes, and you can choose to backup       #
# a specific mailbox or all of them. The script creates a timestamped          #
# directory for the backups and exports each mailbox to a separate PST file.   #
# After the backup is completed, it compares the source and destination item   #
# counts to ensure that the backup was successful.                             #
#                                                                              #
# Functions:                                                                   #
# - Get-Mailboxes: Retrieves a list of mailboxes in the Outlook profile        #
# - Export-PST: Exports the specified mailbox to a PST file, one at a time     #
#                                                                              #
# Usage:                                                                       #
# 1. Run the script as yourself (no admin needed) in a PowerShell console      #
# 2. Enter the mailbox number or press 'Enter' to backup all mailboxes         #
# 3. Check the backup status and file information after the script completes   #
#                                                                              #
# Disclaimer: This script has been tested extensively on a single machine and  #
# has been deemed safe for use in that particular environment.                 #
# If you decide to run this on your host, be prepared for possible fireworks,  #
# implosions, or even an interdimensional rift.                                #
# In other words, it works for me, no guarantees, proceed at your own risk.    #
#                                                                              #
# Final note: If you plan to copy the PST files to an untrusted destination    #
#             (cloud, usb stick), you may want to encrypt them first.          #
#                                                                              #
################################################################################

# Global variables, can be changed
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$backupDir = "$env:USERPROFILE\Backup\Outlook Files\$timestamp"

function Get-Mailboxes {
    param ($Outlook)
    $Namespace = $Outlook.GetNamespace("MAPI")
    $mailboxes = $Namespace.Stores | ForEach-Object { [PSCustomObject]@{ DisplayName = $_.DisplayName; IsOpen = $_.IsOpen } }
    $mailboxes | Where-Object { $_.DisplayName -notlike "Outlook Data File" }
}

function Export-PST {
    param (
        [string]$Name,
        [string]$Path,
        $Outlook
    )

    # Get the MAPI namespace and find the store/mailbox
    $ns = $Outlook.GetNamespace("MAPI")
    $store = $ns.Stores | Where-Object { $_.DisplayName -eq $Name }
	
    $VerbosePreference = "Continue"
	
    # Check if the store is found
    if ($store) {
        # Add a new PST file to the Outlook profile,
        $ns.AddStore($Path)
        $backup = $ns.Stores | Where-Object { $_.FilePath -eq $Path }

        # Check if the PST file is accessible
        if ($backup) {
            $Error.Clear()
            # Copy each folder from the source store to the backup PST file
            $store.GetRootFolder().Folders | ForEach-Object {
				try { 
					$folderItemCount = $_.Items.Count
					Write-Verbose "Exporting folder: $($_.Name) ($folderItemCount items)"
					$_.CopyTo($backup.GetRootFolder()) | Out-Null
				} catch {
					throw "Fatal Error: Could not copy folder $($_.Name)"
				}
			}

            # Just in case the previous error check didn't catch all
            if ($Error) {
                # Remove the backup PST file and return an error message
                $ns.RemoveStore($backup.GetRootFolder())
				throw "Fatal Error: One or more folders failed to copy during the backup process"

            } else {
                # Count the total number of items in the source store
                $srcCount = 0
                $store.GetRootFolder().Folders | ForEach-Object { $srcCount += $_.Items.Count }

                # Verify: Close and immediately reopen the backup PST file
                $ns.RemoveStore($backup.GetRootFolder())
                try {
                    $ns.AddStore($Path)
                    $reopened = $ns.Stores | Where-Object { $_.FilePath -eq $Path }
                } catch { "Error reopening exported PST: $($_.Exception.Message)" }

                # Count the total number of items in the newly created backup file
                $reopenCount = 0
                $reopened.GetRootFolder().Folders | ForEach-Object { $reopenCount += $_.Items.Count }

                # Remove the reopened PST file from the Outlook profile
                $ns.RemoveStore($reopened.GetRootFolder())

                # Compare the source and backup item counts
                if ($srcCount -eq $reopenCount) { 
                    Write-Verbose "Backup status   : $Name completed successfully ($srcCount items)" 
                } else { 
                    Write-Verbose "Backup status   : $Name completed with discrepancies (Source: $srcCount, Destination: $reopenCount items)" 
                }
            }
        } else { 
            throw "Fatal Error: Could not create or access the PST file"
			
        }
    } else { 
        throw "Fatal Error: Mailbox not found"
    }
    Write-Verbose "Housekeeping    : Release Outlook objects to free up memory"
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($reopened)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($backup)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($store)
    $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ns)
}

# Main script starts here
try {
    $Outlook = New-Object -ComObject Outlook.Application
} catch {
    throw "Fatal Error: Creating Outlook application instance failed"
}

# Show the list, and get user input
$mailboxes = Get-Mailboxes -Outlook $Outlook
$mailboxes | Select-Object @{Name = 'Index'; Expression = { $mailboxes.IndexOf($_) + 1 }}, DisplayName, IsOpen | Format-Table -AutoSize
$input = Read-Host "Enter mailbox number, or press 'Enter' to backup all mailboxes"
[int]$parsedInput = 0
$parsed = [int]::TryParse($input, [ref]$parsedInput)

# Verify and further process the answer
if ($input -eq '' -or $input -eq 'all') {
    $mailboxesToBackup = $mailboxes
} elseif ($parsed -and $parsedInput -ge 1 -and $parsedInput -le $mailboxes.Count) {
    $mailboxesToBackup = $mailboxes[$parsedInput - 1]
} else {
	Write-Output "Invalid input"
	exit
}

# Create timestamped directory
try {
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
} catch {
    throw "An error occurred while creating the backup directory: $($_.Exception.Message)"
}

# Iterate over the mailboxes found, show some stats after each backup
foreach ($mailbox in $mailboxesToBackup) {
    $mailboxName = $mailbox.DisplayName -replace '[\\/:"*?<>|]', '_'
	Write-Output "`nBackup started  : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Backup mailbox  : ${mailboxName}"
    $pstPath = Join-Path $backupDir "export_$mailboxName.pst"
    $result = Export-PST -Name $mailboxName -Path $pstPath -Outlook $Outlook
    $fileSize = '{0:N2}' -f ((Get-Item $pstPath).Length / 1MB)
	Write-Output "Backup finished : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "File location   : $pstPath"
    Write-Output "File size       : ${fileSize}MB`n"
}

# Release the Com object
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook)
