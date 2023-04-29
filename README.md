```
################################################################################
#                                                                              #
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
```
