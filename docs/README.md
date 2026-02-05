# Open Reports

> Your Gateway to Integrated Excel Reporting
> Unlock the full potential of your Opera PMS data with high-performance, automated Excel workbooks.

---

### Introduction

Welcome to **Open Reports**, a Microsoft Excel VSTO-based reporting toolkit tailored for the hospitality and finance sectors. This suite provides high-performance modules for dynamic reporting, SQL-driven queries, matrix mapping, and regional compliance, such as Maldives Green Tax reporting.

---

### How It Works

Open Reports utilizes specialized Excel Workbooks designed to address various reporting scenarios, from basic data extracts to complex financial models.

- **Definition Files:** All report configurations are stored in external definition files. These files are portable and can be hosted on a shared network or used locally. Each report type utilizes specific file extensions for its configuration.
- **Mapping Capabilities:** Complex reports requiring data normalization can utilize customized mapping files. These allow for the custom grouping of Opera codes (e.g., revenue centers or transaction codes), which can then be shared across multiple reports.
- **File-Based Architecture:** Unlike traditional reporting tools, Open Reports uses a simplified file structure to store definitions rather than a complex central database.
- **Integrated Connectivity:** Connection settings are stored in a dedicated connection file linked to the workbook. Once configured, the workbook maintains these links automatically.

---

### Getting Started

Open Reports is distributed as a secure Visual Studio Tools for Office (VSTO) customization. To ensure a seamless installation, the client environment must be configured to authorize these customizations.

#### System Requirements

- **Operating System:** Windows 8, 10, or 11 (32-bit and 64-bit supported).
- **Microsoft Excel:** Excel 2013 or later / Office 365 (32-bit and 64-bit supported).
- **Runtimes:** .NET Framework 4.8 and Visual Studio 2010 Tools for Office Runtime.
- **Database Client:** No separate Oracle Client installation is required; Open Reports includes an integrated Oracle client.

#### Client Setup

Before installing workbooks, run the **Client Setup Utility** as an **Administrator**. This utility only needs to be executed once per machine to establish the following security policies:

- **Certificate Installation:** Registers Hospitality Technology as a verified and trusted publisher.
- **Trusted Locations:** Authorizes Excel to execute the VSTO customization.
- **Network Authorization:** Configures Internet Options to allow for secure updates and file downloads.
- **Security Bypass:** Eliminates recurring security warnings and macro blocks.

:link: [Download Client Setup Installer](https://htmaldives.com/apps/setup_utility.exe)

---

### Available Reports

1. **Dynamic Report**
   - **Functionality:** Interactive report building with a drag-and-drop interface for column organization.
   - **Output:** Instant tabular data views based on predefined data models.
   - **File Extension:** `*.orpt`
   - **Download:** :link: [Dynamic Report Workbook](https://htmaldives.com/apps/DynamicReport/Open_Dynamic_Report.xlsx)
2. **Query Report**
   - **Functionality:** Direct execution of custom SQL queries within Excel.
   - **Output:** Results rendered into Excel tables with support for parameterized filtering.
   - **File Extension:** `*.orquery`
   - **Download:** :link: [Query Report Workbook](https://htmaldives.com/apps/QueryReport/Open_Query_Report.xlsx)
3. **Mapping Report**
   - **Functionality:** Defines pivot-style matrix structures and logic-based mapping rules.
   - **Use Case:** Ideal for revenue grouping and chart of account allocations.
   - **File Extension:** `*.orgl` and `*.ormap`
   - **Download:** :link: [Mapping Report Workbook](https://htmaldives.com/apps/MappingReport/Open_Mapping_Report.xlsx)
4. **Green Tax Report (Maldives)**
   - **Functionality:** Standadized Maldives Guest Information sheet, with option to fix sequence, and setup customized fields
   - **File Extension:** `*.orgtax`
   - **Download:** :link: [Green Tax Report Workbook](https://htmaldives.com/apps/GreenTaxReport/GreenTaxReport.xlsx)

---

### Connection Setup and Authentication

#### Creating a New Connection

Upon the initial launch, you may be prompted to select an existing connection file. If none exists, this step can be bypassed to create a new one:

1. Open the workbook and navigate to the **Reports** tab in the Excel Ribbon.
2. Select **Configuration** and enable editing using the System Admin or Support password.
3. Access the **Connection Menu** and select **Create New Connection**.
4. Define the following parameters: **Property Code, Hostname, Port, Opera Schema Name,** and **Opera Schema Password**.
5. Test the connection using the built-in client and save.
6. Multiple properties can be added for multi-property or shared environments.

> [!NOTE]
> Connection details are stored in a hidden temporary sheet within the workbook, ensuring they remain attached for future sessions.

#### Authentication

1. Navigate to the **Login** button on the Reports tab.
2. Select the relevant **Property Code** and enter your **Opera credentials**.
3. Super Admin users will have report configuration automatically enabled upon login.
4. Use the **Remember Me** option to securely cache credentials within the workbook’s temporary folder for faster access.

---

### Running Reports

1. **Loading Definitions:** Open a definition file and click **Run**. The workbook automatically remembers the last definition file used and reloads it upon opening.
2. **Parameters:** A parameter screen will appear based on the report setup. Enter the required criteria (e.g., date ranges or filters).
3. **Execution:** Click **Run** to generate the data. The active sheet will be cleared and regenerated with the new results.
4. **Persistence:** Last-used parameters are stored in a hidden row (Row 1) so they are pre-filled for subsequent runs.

---

### Technical Support

| **Resource**       | **Contact Information**                                                                       |
| ------------------ | --------------------------------------------------------------------------------------------- |
| **Email Support**  | shassan@hospitalitytechnology.com.mv                                                          |
| **Support Portal** | [https://support.hospitalitytechnology.com.mv](https://support.hospitalitytechnology.com.mv/) |

---

<button-counter></button-counter>
