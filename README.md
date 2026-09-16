# ARS Bulk Computer Prestage

### Prepare computer accounts in a few guided steps

**User Guide · Version 2.1**  
Created by **CJ Micklitsch**

Import a list of computer names, choose a destination, review the settings, and create or update the accounts through **One Identity Active Roles**.

> **What does “prestage” mean?**  
> It means creating a computer account in advance. This utility prepares the account; it does not join the physical computer to the domain.

---

## 🚀 Quick start

**From GitHub:** download and extract the repository, then copy `config.example.ini` to `config.ini` and enter your own connection settings and OU mappings. Local `config.ini` files, `.ars` files, installer media, and generated reports are excluded from Git. Obtain the Active Roles installers separately and arrange them as shown in the package layout below.

1. **Open the utility** using your provided launcher.
2. **Check prerequisites** under **Tools → Active Roles Prerequisites**. Use **Install All** if components are missing.
3. **Check AR Server and Domain DN** against the values supplied by your administrator.
4. **Browse to your CSV** containing the computer names.
5. **Choose Destination OU** from the friendly-name dropdown.
6. **Choose a Report Folder** where you can save files.
7. **Review the preview** showing the full destination and attributes to apply.
8. Click **Start Import**, review the confirmation, and choose **Yes** to proceed.
9. **Read the results report**, especially any rows marked `FAILED`.

> **Existing computer accounts are updated, not moved.**  
> The selected destination applies to **new** accounts. Configured attributes are applied to both new and existing accounts.

---

## ✅ Before you begin

Have these ready:

- A Windows computer with **64-bit Windows PowerShell 5.1** and the supplied launcher or script.
- Network access to your Active Roles server.
- A Windows account with permission to create or update the intended computer accounts through Active Roles.
- A computer-name CSV and a report folder you can write to.
- The supplied `config.ini`, with your server settings and destination choices.

Keep the script, `config.ini`, and `Prerequisites` folder together. If you received an existing configuration, keep it rather than replacing it with the examples below.

### First-time setup

Open **Tools → Active Roles Prerequisites** to see:

| Component | Action when missing |
|---|---|
| ADSI Provider | Click its **Install** button. |
| SDK | Click its **Install** button. |
| PowerShell Module | Click its **Install** button. |
| All missing components | Click **Install All**. |

**Install All** skips installed components and installs the missing ones in order: ADSI Provider → SDK → PowerShell Module. Windows may request administrator approval or credentials once for the batch.

If an installer fails, later components are not installed. Review the result and the log location shown in the installation window. If a restart is requested, restart Windows and reopen the utility.

The submenu also includes **Refresh Module / Prerequisite Status**. Installation credentials are handled by Windows; the utility continues using your original Windows account for Active Roles operations.

---

## 📋 Prepare your computer list

Your CSV must have a column named **`ComputerName`**, with one computer name per row:

```csv
ComputerName
PC-001
PC-002
PC-003
```

**Using Excel:** enter `ComputerName` in cell A1, place the names below it, and save as **CSV**. An Excel workbook (`.xlsx`) is not the import format.

Check the list before importing. Remove accidental duplicates and confirm that every computer belongs in the selected destination.

---

## 🖥️ Understand the main window

| Field or section | What to do |
|---|---|
| **AR Server** | Enter the Active Roles server supplied by your administrator. |
| **Domain DN** | Enter the domain’s full directory name, such as `DC=example,DC=com`. |
| **Computers CSV** | Choose your prepared CSV file. |
| **Destination OU** | Select the friendly name for the destination. |
| **Report Folder** | Choose an existing folder for the results CSV. |
| **Configured Attributes** | Check the number of attributes that will be applied. |
| **Preview below the count** | Review the selected friendly name, full OU designation, and every attribute/value. |

**OU** means *Organizational Unit*: a folder-like location in the directory. Its **DN**, or *Distinguished Name*, is the full address of that location.

The attribute count and preview refresh from `config.ini` while the window is open. The OU choices refresh too. If a selected OU mapping changes or disappears, select a destination again.

### Before confirming the import

- Check the **full OU designation**, not just its friendly name.
- Check the **attribute names and values** in the preview.
- Check the **computer count and names** in the confirmation. For longer lists, the confirmation shows the first ten names and the remaining count.

Clicking **Start Import** first validates the settings and opens the confirmation. Choose **Yes** there to begin creating or updating accounts.

---

## ⚙️ Configure destinations and attributes

Use **Config → Open Config File** to open `config.ini` in Notepad. Save your edits in Notepad so the utility can read them.

### Add destination choices

Each entry in `[OU]` uses this format:

```ini
FriendlyName=Full OU designation
```

For example:

```ini
[OU]
Example=OU=Computers,DC=example,DC=com
```

The dropdown displays **Example**. The utility uses the entire designation after the first `=` for validation and import. The other `=` signs in the designation are preserved.

Add one destination per line and give each a unique friendly name. Use the exact OU designation supplied by your administrator. The example above is a format example, not a confirmation that the OU exists or is accessible.

### Choose attributes to apply

The `[Attributes]` section lists the properties to set on each computer account:

```ini
[Attributes]


```

The example configuration has no attributes configured. Confirm the desired attributes with your administrator before adding them. `TRUE` and `FALSE` are interpreted as Boolean values, and whole-number values are interpreted as numbers. If `config.ini` is missing, the script retains its inherited defaults: `edsvaCHSServer=FALSE` and `edsaJoinComputerToDomain=US\Domain Users`. Create your configuration before first use to choose the settings appropriate for your environment.

- Adding a line adds an attribute to the preview and import.
- Removing a line removes it from the attributes to apply.
- An empty `[Attributes]` section means no configured attributes will be applied.
- If `config.ini` is missing, the utility uses its defaults and labels the count accordingly.

If the file cannot be read, the count displays an error. Saving or starting an import is blocked rather than silently using outdated attributes.

### Save or reset settings

| Config menu item | What it does |
|---|---|
| **Open Config File** | Opens the file for editing; creates it if needed. |
| **Save Current Settings** | Saves the current GUI settings and attributes, preserving the `[OU]` mappings. |
| **Reset Saved Settings** | After confirmation, deletes the saved configuration and clears the GUI settings. |

> **Keep a backup before resetting.**  
> Reset also removes custom OU mappings and custom attributes stored in that file.

---

## 📊 Read your results

When processing finishes, the utility displays a summary and saves a CSV in your chosen report folder:

```text
ComputerImportResults_yyyyMMdd_HHmmss.csv
```

| Report status | Meaning |
|---|---|
| **Created** | A new computer account was created in the selected OU. |
| **Already Existed - Updated** | An existing account received the configured attributes. Its location was not changed. |
| **FAILED** | That computer could not be processed. Read its **Error** column. |

The report also includes the computer name, location, and number of attributes applied. Keep it as your record of the run. **Tools → Open Reports Folder** opens the folder currently selected in the GUI.

---

You can view this guide inside the utility using **About → View User Guide (README.md)**.

## 🔎 Common questions

| What you see | What to check |
|---|---|
| **Start Import is disabled** | Check **Tools → Active Roles Prerequisites**. All components and required module commands must be available; complete any requested restart. |
| **Destination dropdown is empty** | Add entries to `[OU]` in the `config.ini` beside the script, then save the file. |
| **A destination becomes unselected** | Its mapping may have changed or been removed. Review the file and choose the intended destination again. |
| **Attributes do not match your expectations** | Save your editor changes and check the configuration path displayed in the GUI. Make sure you edited that file. |
| **The server connection fails** | Check the server name, network connection, and your access. Use **Tools → Test Active Roles Connection**. |
| **The destination cannot be validated** | Check its complete designation and your permissions. Use **Tools → Validate Destination OU**. |
| **An installer is missing** | Check the package layout below, including the exact folder and file names. |
| **The report cannot be saved** | Confirm that the report folder exists and that you can create files there. |

If you need help, provide the error message and relevant report or installer log to your support team.

---

<details>
<summary><strong>📁 Package layout — for whoever prepares the folder</strong></summary>

Keep this structure. The package can be moved to another location as long as the relative layout is preserved.

```text
ARS-Bulk-Prestage-ADComputers/
├── ARS-Bulk-Prestage-ADComputers.ps1
├── config.ini
├── README.md
└── Prerequisites/
    ├── ActiveRoles ADSI Provider - For exporting objects and reports/
    │   └── x64/
    │       ├── ADSI/
    │       │   └── _x64.msi
    │       └── SDK/
    │           └── _x64.msi
    └── ActiveRoles Powershell Module/
        └── x64/
            └── Shell/
                └── _x64.msi
```

Keep any accompanying installer files, such as CAB files, in their supplied component folders. A separately supplied launcher may also be included in the package.

</details>

---

**About this guide**  
Applies to script version **2.1**. Local syntax and behavior checks have been completed; installation and directory operations still need validation in your environment.
