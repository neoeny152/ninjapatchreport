# **NinjaOne Windows Patch Compliance Report**

A PowerShell script for NinjaRMM that generates a monthly Windows patch compliance report in HTML and CSV formats.

This script pulls all active Windows devices from the NinjaOne API, filters for CURRENT MONTHs critical "Patch Tuesday" updates (Cumulative Updates and .NET Framework), and generates an actionable HTML report with at-a-glance summaries and a detailed list of all non-compliant machines.

## **How to Set Up in NinjaRMM**

Follow these steps to get the script running in your environment.

### **1\. Create API Credentials in NinjaOne**

If you don't have them already, create a new API client:

1. In NinjaOne, navigate to **Configuration** > **Integrations** > **API**.
2. Click **Add** and create a new **Client ID / Secret Credential**.
3. Grant it the **Monitoring** and **Management** scopes.
4. Securely save the **Client ID** and **Client Secret**.

### **2\. Create Custom Fields for Credentials**

The script uses NinjaRMM's secure custom fields to access your API credentials.

1. Navigate to **Configuration** > **Global Custom Fields**.
2. Click **Add a Field** and create the following three fields. **The "Name" must match exactly.**

| Name | Type | Label |

| mescriptNinjaurl | Secret Text | Ninja API URL |

| mescriptNinjacid | Secret Text | Ninja API Client ID |

| mescriptNinjasec | Secret Text | Ninja API Client Secret |

1. Once created, find the device you will use to run the script (e.g., a domain controller or management server). Go to that device's details page, find the **Custom Fields** section, and enter your API credentials into the values for these new fields.

### **3\. Create the Script in NinjaRMM**

1. Navigate to **Configuration** > **Scripting**.
2. Click **Create New Script**.
3. Configure the script settings:
    - **Name:** Generate Windows Patch Report
    - **Description:** Generates a monthly patch compliance report in HTML and CSV format.
    - **Language:** PowerShell
    - **Operating System:** Windows
    - **Architecture:** All
4. Copy the entire content of the Generate-NinjaPatchReport.ps1 file and paste it into the script editor.

### **4\. Schedule the Script to Run**

1. Create a new scheduled task or policy to run the script.
2. **IMPORTANT:** Set the script to **Run As: System**. This is required for the Ninja-Property-Get commands to access the secure custom fields.

### **5\. (Optional) Configure File Share Copy**

If you want the script to automatically copy the reports to a network share:

1. In the script's ★★★ SCRIPT CONFIGURATION ★★★ section, uncomment the $fileSharePath variable and replace the placeholder with your actual UNC path.
2. When scheduling the script, use NinjaRMM's **"Network Credentials"** feature. Provide a user account (like your svcuem domain admin) in this section that has permission to write to the file share.

This allows the script to run as SYSTEM locally to get the API keys, while using your specified credentials only for accessing the network.
