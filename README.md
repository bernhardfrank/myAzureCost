# myAzureCost v2 is a rehaul of the previous version to make it work with the latest Azure updates.  

Wrapping up your daily Azure consumption and cost and sending it to you via email + attached csv and printable html.

The solution uses:  
- a storage account with a table to hold every days cost (will be filled day by day)
- a communication service that will send emails to you.
- an azure automation account with a PowerShell 7.2 runbook to process the data.
- a system managed identity with minimum rights required (Cost Management Reader on subscription, Contributor on communicationServices, Storage Table Data Contributor + Storage Blob Data Contributor on storageaccount)

[![Deploy to Azure](https://aka.ms/deploytoazurebutton)](https://portal.azure.com/#create/Microsoft.Template/uri/https://raw.githubusercontent.com/bernhardfrank/myAzureCost/masterV2/azuredeploy.json)  

# Result & Screenshots  
  
| ![email](./pics/email.PNG)  | ![attached csv](./pics/csv.PNG)  | ![attached printable html](./pics/printable.PNG) |![azure deployment](./pics/deployment.PNG) |
|--|--|--|--|
| **cost email contains graphs** | Attached a CSV with details e.g. **quantity and cost in your locale** | Attached a html for better **print** | How the deployment should look like. |
