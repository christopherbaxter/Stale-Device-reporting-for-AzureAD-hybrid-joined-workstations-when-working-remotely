# StaleComputerAccounts
Managing stale workstations when most users are working from home using AzureAD in a Hybrid configuration is a significant challenge. Here is some code that will help.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes, I would not advise copying my permissions. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what workd for now. 