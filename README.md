# StaleComputerAccounts
Managing stale workstations when most users are working from home using AzureAD in a Hybrid configuration is a significant challenge. Here is some code that will help.

## What is needed for this script to function?

You will need a Service Principal in AzureAD with sufficient rights. I have a Service Principal that I use for multiple processes, I would not advise copying my permissions. I suggest following the guide from <https://msendpointmgr.com/2021/01/18/get-intune-managed-devices-without-an-escrowed-bitlocker-recovery-key-using-powershell/>. My permissions are set as in the image below. Please do not copy my permissions, this Service Principal is used for numerous tasks. I really should correct this, unfortunately, time has not been on my side, so I just work with what workd for now. 

![](https://github.com/christopherbaxter/StaleComputerAccounts/blob/main/ServicePrincipal%20-%20API%20Permissions.jpg)

I also elevate my AzureAD account to 'Intune Administrator', 'Cloud Device Administrator' and 'Security Reader'. These permissions also feel more than needed. Understand that I work in a very large environment, that is very fast paced, so I elevate these as I need them for other tasks as well.

You will need to make sure that you have the following Powershell modules installed. There is a lot to consider with these modules as some cannot run with others. This was a bit of a learning curve. 

ActiveDirectory
AzureAD
ImportExcel
JoinModule
MSAL.PS
PSReadline (May not be needed, not tested without this)

Ultimately, I built a VM on-prem in one of our data centers to run this script, including others. My machine has 4 procs and 16Gb RAM, the reason for an on-prem VM is because most of our workforce is working from home (me included), and running this script is a little slow through the VPN. Our ExpressRoute also makes this data collection significantly more efficient. In a small environment, you will not need this.

### How does the script work?

