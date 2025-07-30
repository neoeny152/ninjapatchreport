# **NinjaOne Windows Patch Compliance Report**

A PowerShell script for NinjaRMM that generates a monthly Windows patch compliance report in HTML and CSV formats.

This script pulls all active Windows devices from the NinjaOne API, filters for the current month's critical "Patch Tuesday" updates (Cumulative Updates and .NET Framework), and generates an actionable HTML report with at-a-glance summaries and a detailed list of all non-compliant machines.

## **Prerequisites: The Secure Script Runner**

For security, it is critical to run this script from a dedicated and isolated machine. This machine will have access to API credentials with high-level permissions.

1. **Select a Dedicated Machine**: Use a dedicated machine for running scripts that will not be used for other tasks. A Windows Server Core VM is a great choice. This machine should only be accessible to System Administrator accounts.
2. **Create a "Scripting" Organization**: In NinjaOne, create a new organization (e.g., "Internal Scripting") to house the dedicated script runner machine. This isolates it from your client organizations.
3. **Create a Custom Role**: Navigate to **Configuration** > **Roles**. Create a new custom role (e.g., "Script Runner Role") for your script machine.
4. **Secure Technician Permissions**: Update your technician roles to ensure only System Administrators have access to the new "Scripting" organization and the "Script Runner Role".
5. **Assign the Machine**: Install the NinjaOne agent on your dedicated machine and move it into the new "Scripting" organization. Edit the device and assign it the "Script Runner Role".

## **How to Set Up in NinjaRMM**

### **1\. Create API Credentials in NinjaOne**

First, create a new API client for the script to use.

- In NinjaOne, navigate to **Configuration** > **Integrations** > **API**.
- Click **Add** and create a new **Client ID / Secret Credential**.
- Grant it the **Monitoring** and **Management** scopes.
- Securely save the **Client ID** and **Client Secret** for the next steps.

### **2\. Create Role Custom Fields for Credentials**

The script uses NinjaRMM's secure custom fields to access your API credentials. By creating them as **Role Custom Fields**, you ensure they are only available to the secure role you created earlier.

- Navigate to **Configuration** > **Roles**.
- Select the **Script Runner Role** you created.
- Go to the **Custom Fields** tab.
- Click **Add a Field** and create the following three fields. The **Name** is used by the script and must match exactly. The **Label** is the display name you see in the UI.

| **Label (Display Name)** | **Name (Internal)** | **Type** | **Permissions** | | Ninja API URL | ninjaoneInstance | Secret Text | Technician: Editable, Automations: Read Only | | Ninja API Client ID | ninjaoneClientId | Secret Text | Technician: Editable, Automations: Read Only | | Ninja API Client Secret | ninjaoneClientSecret | Secret Text | Technician: Editable, Automations: Read Only |

### **3\. Populate Credentials on the Script Runner Device**

- Navigate to the dedicated script runner device within NinjaOne.
- In the device details, find the **Custom Fields** section.
- Enter your API URL, Client ID, and Client Secret into the values for the fields you just created.

### **4\. Create the Script in NinjaRMM**

- Navigate to **Configuration** > **Scripting**.
- Click **Create New Script**.
- Configure the script settings:
  - **Name**: Generate Windows Patch Report
  - **Description**: Generates a monthly patch compliance report in HTML and CSV format.
  - **Language**: PowerShell
  - **Operating System**: Windows
  - **Architecture**: All
- Copy the entire content of the Generate-NinjaPatchReport.ps1 file and paste it into the script editor.

### **5\. Schedule the Script to Run**

- Create a new scheduled task or policy to run the script on your dedicated machine.
- **IMPORTANT**: Set the script to **Run As: System**.
  - **Why?** The script uses Ninja's built-in Ninja-Property-Get command to read the secure custom fields. This command requires SYSTEM privileges to function.

### **6\. (Optional) Configure File Share Copy**

If you want the script to automatically copy the reports to a network share:

- In the script's ★★★ SCRIPT CONFIGURATION ★★★ section, uncomment the $fileSharePath variable and replace the placeholder with your actual UNC path.
- When scheduling the script, use NinjaRMM's **"Network Credentials"** feature. Provide a user account (like a service account) in this section that has permission to write to the file share.
- This allows the script to run as SYSTEM locally (to get the API keys) while using your specified credentials only for accessing the network share.
