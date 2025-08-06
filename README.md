# Gather-O365-User-Data
 gather data from o365 that we can use internally
This script goes to Azure in O365 and gathers data like the MFA status of users. 

1. Go to the folder and update the env if you need to. (if there is no .env it will create one and you can edit it)
2. run PSGatherUsers.ps1 it will prompt you go each account.
3. for each client it will create a csv file like acme_users_01012025.csv

<!-- Purpose: This script goes to Azure in O365 and gathers data like the MFA status of users. -->
<!-- INSTALL_COMMAND: curl -o PSGatherUsers.ps1 https://github.com/mrdatawolf/PSHTMLOFTOOLS/raw/main/PSGatherUsers.ps1; -->
<!-- RUN_COMMAND: PSGatherUsers.ps1 -->