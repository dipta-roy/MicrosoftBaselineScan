### README: Offline Microsoft Baseline Scanner with PowerShell

#### **Description**

This PowerShell script provides a simple and efficient way to scan for missing Windows updates on a system using an offline CAB file (`wsusscn2.cab`). The script leverages the Windows Update COM API to search for updates that are not installed and exports the results to a CSV file for further review. It’s ideal for environments where internet connectivity is limited or restricted, and updates need to be assessed offline.


#### **How It Works**

1. **Administrative Privileges**:  
   The script validates if it is running with administrative privileges. If not, it notifies the user to run with administrator privilege.

2. **CAB File Selection**:  
   Users are prompted to select the `wsusscn2.cab` file, which contains the metadata required for offline update scanning.

3. **Update Scanning**:  
   The script uses the `Microsoft.Update.Session` and `Microsoft.Update.ServiceManager` COM objects to search for missing updates based on the CAB file.

4. **Exporting Results**:  
   The scan results, including update titles and descriptions, are saved in a CSV file. The file is named using the format:  
   `COMPUTERNAME_YYYY-MM-DD_HH-MM-SS.csv`  
   and is saved in the same directory as the script for easy access.

---

#### **Requirements**

1. **PowerShell Execution Policy**:  
   The script must run with the execution policy set to allow scripts. 
   Example: `-ExecutionPolicy Bypass`

2. **Offline CAB File**:  
   The script requires the latest `wsusscn2.cab` file to perform the update scan.  
   You can download the file from the official Microsoft Update Catalog:  
   [Download wsusscn2.cab](https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab)

3. **Administrator Rights**:  
   The script must be executed with administrative privileges to access the Windows Update API.

---

#### **Features**

- **Graphical File Picker**:  
   Users are prompted with a graphical file picker to select the CAB file, making it user-friendly.
  
- **Customizable Output**:  
   Results are saved in a structured CSV format, including the update title, description, and index for easy analysis.
  
- **Error Handling**:  
   The script includes robust error handling to ensure issues are captured and logged, helping users debug any problems.

---

#### **Usage Instructions**

1. **Download the Script**:  
   Save the PowerShell script (e.g., `UpdateScanner.ps1`) to your preferred location.

2. **Download the CAB File**:  
   Visit [this link](https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab) to download the `wsusscn2.cab` file. Place it in an accessible directory on your system.

3. **Run the Script**:  
   Open a PowerShell terminal, navigate to the directory containing the script, and execute it:  
   ```powershell -ExecutionPolicy Bypass
   .\MicrosoftBaselineScan.ps1
   ```

4. **Follow the Prompts**:  
   - Select the downloaded `wsusscn2.cab` file when prompted.  
   - Wait while the script scans for missing updates. A progress message will be displayed.  
   - After completion, find the CSV output in the same directory as the script.

5. **Review the Results**:  
   Open the CSV file in a spreadsheet editor like Microsoft Excel or any text editor to view the list of missing updates.

---

#### **Output Example**

The output CSV file includes the following columns:  
- **Index**: Sequential number of the update.  
- **Title**: The name of the update.  
- **Description**: A brief description of the update.

Example:  

| Index | Title                                  | Description                              |
|-------|----------------------------------------|------------------------------------------|
| 1     | Security Update for Windows 10 (KB123)| Fixes a security vulnerability.          |
| 2     | Update for .NET Framework (KB456)     | Improves .NET Framework performance.     |

---

#### **FAQ**

1. **Where should I save the CAB file?**  
   Anywhere on your computer. The script will prompt you to locate and select the file during execution.

2. **Why does the script require admin rights?**  
   Scanning for updates using the Windows Update API requires elevated privileges to access system-level information.

3. **What happens if there are no missing updates?**  
   The script will notify you that there are no applicable updates and terminate without creating a CSV file.

4. **Can I use this script on multiple systems?**  
   Yes, you can copy the script and CAB file to other systems and run it as long as the requirements are met.

---

#### **Known Limitations**

1. The script only works with Windows operating systems that support the `Microsoft.Update.Session` COM object.
2. The accuracy of the scan depends on the freshness of the `wsusscn2.cab` file. Ensure you download the latest version regularly.

---

Feel free to modify and customize the script to fit your needs.
