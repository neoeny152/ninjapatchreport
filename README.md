# **NinjaOne Windows Patch Compliance Report**

A PowerShell script for NinjaRMM that generates a monthly Windows patch compliance report in HTML and CSV formats.

This script pulls all active Windows devices from the NinjaOne API, filters for the current month's critical "Patch Tuesday" updates (Cumulative Updates and .NET Framework), and generates an actionable HTML report with at-a-glance summaries and a detailed list of all non-compliant machines.

## **Prerequisites: The Secure Script Runner**

For security, it is critical to run this script from a dedicated and isolated machine. This machine will have access to API credentials with high-level permissions.

1. **Select a Dedicated Machine**: Use a dedicated machine for running scripts that will not be used for other tasks. A Windows Server Core VM is a great choice. This machine should only be accessible to System Administrator accounts.
2. **Create a "Scripting" Organization**: In NinjaOne, create a new organization (e.g., "Internal Scripting") to house the dedicated script runner machine. This isolates it from your client organizations.
3. **Create a Custom Role**:
    - Navigate to **Home** > **Administration** > **Devices** > **Roles**.
    - Scroll to the **Windows Server** section, click the ellipsis (**...**), and select **Add** to create a new role named "Script Runner Role".
4. **Secure Technician Permissions**: Update your technician roles to ensure only System Administrators have access to the new "Scripting" organization and the "Script Runner Role".
5. **Assign the Machine**: Install the NinjaOne agent on your dedicated machine and move it into the new "Scripting" organization. Edit the device and assign it the "Script Runner Role".

## **How to Set Up in NinjaRMM**

### **1\. Create API Credentials in NinjaOne**

First, create a new API client for the script to use.

- In NinjaOne, navigate to **Home** > **Administration** > **Apps** > **API**.
- Click **\+ Add client app** and create a new credential.
- Grant it the **Monitoring** and **Management** scopes.
- Securely save the **Client ID** and **Client Secret** for the next steps.

### **2\. Create Role Custom Fields for Credentials**

The script requires NinjaOne API details stored in secure custom fields. These must be **Role Custom Fields** assigned only to your "Script Runner Role" to maintain security.

- Navigate to **Home** > **Administration** > **Devices** > **Roles**.
- Select the **Script Runner Role** you created.
- Go to the **Device Custom Fields** tab.
- Click **Add a Field** and create the following three fields.
  - **Important**: The Name must match exactly, as the script uses it directly. The Label is the friendly display name you will see in the UI.

| **Label (Display Name)** | **Name (Internal)** | **Type** | **Permissions** | **Description** |
| --- | --- | --- | --- | --- |
| NinjaOne Instance | ninjaoneInstance | Text | Technician: Editable&lt;br&gt;Automations: Read Only&lt;br&gt;API: None | Stores the name of your NinjaOne Instance (e.g., app.ninjarmm.com or eu.ninjarmm.com) |
| --- | --- | --- | --- | --- |
| NinjaOne Client ID | ninjaoneClientId | Secure | Technician: Editable&lt;br&gt;Automations: Read Only&lt;br&gt;API: None | Stores the API Client ID |
| --- | --- | --- | --- | --- |
| NinjaOne Client Secret | ninjaoneClientSecret | Secure | Technician: Editable&lt;br&gt;Automations: Read Only&lt;br&gt;API: None | Stores the API Client Secret |
| --- | --- | --- | --- | --- |

### **3\. Populate Credentials on the Script Runner Device**

- Navigate to the dedicated script runner device within NinjaOne.
- In the device details, find the **Custom Fields** section.
- Enter your API URL, Client ID, and Client Secret into the values for the fields you just created.

### **4\. Create the Script in NinjaRMM**

- Navigate to **Home** > **Administration** > **Library** > **Automation**.
- Click **\+ Add**.
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

You can configure the script to automatically save a copy of the reports to a network share.

**Disclaimer:** This feature is difficult to configure and should be attempted at your own risk. Because the script must run as the SYSTEM account to access the secure API credentials, it does not have inherent permissions to access network resources. Authenticating to a network share from the SYSTEM context is complex and may not work reliably in all environments.

If you wish to attempt this:

1. In the script's ★★★ SCRIPT CONFIGURATION ★★★ section, uncomment the $fileSharePath variable.
2. Replace the placeholder with your UNC path (e.g., \\\\YourServer\\YourShare\\PatchReports).
3. You will need to independently solve the network authentication challenge for the SYSTEM account on your script runner machine to access the specified path.
