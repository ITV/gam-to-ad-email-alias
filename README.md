# GAM (G-Suite) to Active Directory
### Email Alias import

This script takes email aliases from GSuite and imports them to the 
proxyAddresses field in Active Directory in the format SMTP: addresss 

This will ensure GSuite and Active Directory aliases are aligned prior to enabling syncronisation (GCDS)

Step 1 - use the GAM tool (https://github.com/jay0lee/GAM) and export users
```
gam print users allfields > users.csv
```

Step 2 - Copy the script in this repo alias_add.ps1 and the export into a directory on a domain controller

Step 3 - Run the script as a dry run
```
.\alias_add.ps1 -csv users.csv
```
Step 4 - Review the debug output to ensure you are comftable with the changes that would be commited
         Debug and error log files will be written to the same directory

Step 5 - Commit the changes
```
.\alias_add.ps1 -csv allusers.csv -commit
```

Flags that can be used:
```
-csv (Mandatory) - the input GAM CSV file
-commit - to commit changes
-aggressive - Skips the delay in between users. Will place Active Directory under additional load
```

The GAM output will contain both a number of rows aliases.{integer} and nonEditableAliases.{integer}
The script will not import aliases under nonEditableAliases these are "domain aliases" that are appied to all users in the G-Suite domain
so do not need to be applied to individual users.

