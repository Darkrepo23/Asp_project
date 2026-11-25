# 🧾 Personal Information Management System

> A classic **ASP.NET Web Forms** application built with **Visual Studio 2010**, **VB.NET**, and **SQL Server**. It stores, displays, edits, and deletes personal information records through a clean card-based interface.

---

## ✨ At a glance

| | |
|---|---|
| **Type** | ASP.NET **Web Site** (not Web Application) |
| **Language** | Visual Basic .NET |
| **Framework** | .NET Framework 4.0 |
| **IDE** | Visual Studio 2010 |
| **Database** | SQL Server / SQL Server Express (`asp_website`) |
| **Data access** | `SqlConnection`, parameterized inline SQL via `App_Code\Connection.vb` |
| **UI** | Web Forms server controls + plain CSS (no frameworks) |
| **Dev server** | ASP.NET Development Server, port `52632`, path `/WebSite1` |

---

## 🚀 Features

- **14 input controls** covering identity, contact, and emergency-contact details
- **Create** records with the **Save Information** button
- **Read** records in a styled GridView sorted by newest first
- **Update** records inline with **Modify → Save → Cancel**
- **Delete** records with a confirm-style action link
- **Validation** for required names, email format, birth date, and 11-digit mobile numbers
- **Clear** button that resets the entire form
- **Responsive layout** with cards, grid, hover states, and a fade-in animation
- **Red Save button** in edit mode, tagged via `btn-save` CSS class

---

## 🧱 Project structure

```text
Visual Studio 2010\
├── Projects\WebSite1\                 ← Solution folder / Git repository
│   ├── WebSite1.sln
│   ├── WebSite1.suo
│   └── README.md                      ← this file
│
└── WebSites\WebSite1\                 ← Actual website content
    ├── Default.aspx                   ← Markup: form + GridView
    ├── Default.aspx.vb                ← Code-behind: validation & CRUD
    ├── web.config                     ← Compilation settings
    ├── style\
    │   └── StyleSheet.css             ← All styling
    └── App_Code\
        └── Connection.vb              ← SQL Server connection helper
```

> 💡 The `.sln` is a pointer: Visual Studio opens the site from `WebSites\WebSite1`, not from the solution folder. Files in `App_Code` are compiled automatically by a Web Site project.

---

## 🗃️ Database

**Database:** `asp_website`  
**Table:** `dbo.PersonalInformation`

> The schema below matches what the current code expects. If your live table differs, update the columns or the SQL in `Default.aspx.vb`.

### Columns

| # | Column | SQL type | Nullable | Form control |
|---:|---|---|:---:|---|
| 1 | `ID` | `INT IDENTITY(1,1)` | No | *(primary key, hidden)* |
| 2 | `FirstName` | `NVARCHAR(100)` | No | `txtFirstName` |
| 3 | `MiddleName` | `NVARCHAR(100)` | Yes | `txtMiddleName` |
| 4 | `LastName` | `NVARCHAR(100)` | No | `txtLastName` |
| 5 | `DisplayName` | `NVARCHAR(100)` | No | `txtDisplayName` |
| 6 | `BirthDate` | `DATE` | Yes | `txtBirthDate` |
| 7 | `Gender` | `NVARCHAR(10)` | Yes | `rblGender` |
| 8 | `CivilStatus` | `NVARCHAR(20)` | Yes | `ddlCivilStatus` |
| 9 | `Nationality` | `NVARCHAR(100)` | Yes | `txtNationality` |
| 10 | `PlaceOfBirth` | `NVARCHAR(255)` | Yes | `txtPlaceOfBirth` |
| 11 | `Email` | `NVARCHAR(255)` | Yes | `txtEmail` |
| 12 | `Contact` | `NVARCHAR(11)` | Yes | `txtContact` |
| 13 | `Address` | `NVARCHAR(500)` | Yes | `txtAddress` |
| 14 | `EmergencyName` | `NVARCHAR(200)` | Yes | `txtEmergencyName` |
| 15 | `EmergencyContact` | `NVARCHAR(11)` | Yes | `txtEmergencyContact` |
| 16 | `CreatedDate` | `DATETIME2(0)` | No | *(default: `SYSDATETIME()`)* |

### Create-table script

```sql
USE [asp_website];
GO

IF OBJECT_ID(N'dbo.PersonalInformation', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.PersonalInformation
    (
        ID                INT IDENTITY(1,1) NOT NULL
            CONSTRAINT PK_PersonalInformation PRIMARY KEY,

        FirstName         NVARCHAR(100) NOT NULL,
        MiddleName        NVARCHAR(100) NULL,
        LastName          NVARCHAR(100) NOT NULL,
        DisplayName       NVARCHAR(100) NOT NULL,
        BirthDate         DATE          NULL,

        Gender            NVARCHAR(10)  NULL
            CONSTRAINT CK_PI_Gender CHECK (Gender IN (N'Male', N'Female')),

        CivilStatus       NVARCHAR(20)  NULL
            CONSTRAINT CK_PI_CivilStatus
            CHECK (CivilStatus IN (N'', N'Single', N'Married', N'Widowed', N'Divorced')),

        Nationality       NVARCHAR(100) NULL,
        PlaceOfBirth      NVARCHAR(255) NULL,
        Email             NVARCHAR(255) NULL,

        Contact           NVARCHAR(11)  NULL
            CONSTRAINT CK_PI_Contact
            CHECK (Contact IS NULL OR
                   (LEN(Contact) = 11 AND
                    Contact LIKE N'09[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]')),

        Address           NVARCHAR(500) NULL,
        EmergencyName     NVARCHAR(200) NULL,

        EmergencyContact  NVARCHAR(11) NULL
            CONSTRAINT CK_PI_EmergencyContact
            CHECK (EmergencyContact IS NULL OR
                   (LEN(EmergencyContact) = 11 AND
                    EmergencyContact LIKE N'09[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]')),

        CreatedDate       DATETIME2(0) NOT NULL
            CONSTRAINT DF_PI_CreatedDate DEFAULT (SYSDATETIME())
    );
END
GO
```

### Connection string

Configured in `App_Code\Connection.vb`:

```vb
Public ConnectionString As String = _
    "Data Source=DESKTOP-IM9M47U\SQLEXPRESS;" &
    "Initial Catalog=asp_website;Integrated Security=True"
```

⚠️ Change the server/database name to match your machine. Windows Authentication is used, so the app runs under the current user's SQL permissions.

---

## 🧠 How the code is organized

### `Default.aspx.vb`

| Member | Purpose |
|---|---|
| `btnSubmit_Click` | Calls `Register()` |
| `Register()` | Validates input, builds the parameterized `INSERT`, then refreshes the grid |
| `btnClear_Click` / `ClearForm()` | Resets every input control and dropdown |
| `Page_Load` | Sets the birth-date textbox to HTML `type="date"`; loads records on first request |
| `LoadRecords()` | Runs `SELECT * ... ORDER BY ID DESC` and binds the GridView |
| `gvRecords_RowEditing` | Puts the selected row into edit mode |
| `gvRecords_RowCancelingEdit` | Cancels edit mode |
| `gvRecords_RowUpdating` | Validates and saves all 14 fields via `UPDATE ... WHERE ID = @ID` |
| `gvRecords_RowDeleting` | Deletes a record by its `ID` data key |
| `gvRecords_RowDataBound` | Adds the `btn-save` CSS class to the Update link in edit mode |
| `GetOptionalValue()` | Converts blank values to `DBNull.Value` |
| `ShowAlert()` | Displays safe JavaScript alert messages |

### `App_Code\Connection.vb`

A tiny helper that owns the `SqlConnection`, collects `SqlParameter`s through `AddParam()`, and exposes `Query(sql)` for both reads and writes. Results land in `Data` / `DataCount`.

---

## ✅ Validation rules

- **First Name, Last Name, Display Name** — required on insert and update
- **Email** — optional, but must look like `username@gmail.com`
- **Contact / Emergency Contact** — optional, but must be exactly `09` + 9 digits
- **Birth Date** — optional, parsed with `DateTime.TryParse`
- **Gender / Civil Status** — optional; blank values are stored as `NULL`

Client-side `oninput` filters also strip non-letters from name fields and non-digits from contact fields.

---

## 🎨 Styling notes (`style\StyleSheet.css`)

- `.card`, `.form-grid`, `.form-group` — responsive form layout
- `.btn-primary` / `.btn-secondary` — the page-level buttons
- `.records-table` — GridView table, headers, hover row
- `.records-table a` — shared action-link styling
- `.records-table a[href*="Edit"]` — blue **Modify** link
- `.records-table a.btn-save` — **red Save** link in edit mode
- `.records-table a[href*="Cancel"]` — grey **Cancel** link
- `.records-table a[href*="Delete"]` — red **Delete** link

> The Save link renders as `__doPostBack('gvRecords$...','')` with no "Update" in its href, so `RowDataBound` assigns `class="btn-save"` explicitly.

---

## ▶️ Running locally

1. Open `WebSite1.sln` in **Visual Studio 2010**.
2. Make sure a SQL Server instance is running and `asp_website` exists with the table above.
3. Adjust the connection string in `App_Code\Connection.vb` if needed.
4. Press **F5** (or **Ctrl+F5** without debugging).
5. The site opens on the ASP.NET Development Server at `http://localhost:52632/WebSite1/`.

If a CSS change doesn't appear, press **Ctrl+F5** in the browser to bypass the cache.

---

## 🔧 Known limitations

- SQL statements are inline (parameterized, but no stored procedures)
- Connection string is hard-coded instead of in `web.config`
- No paging, search, or sorting on the GridView
- No photo upload or file storage
- Age is not stored — derive it from `BirthDate` when needed
- No automated tests

## 💡 Possible improvements

- Move the connection string into `<connectionStrings>` in `web.config`
- Add search/filter and paging to the records table
- Convert the GridView action column to `TemplateField`s for finer control
- Add a photo upload field
- Add unit tests for validation helpers

---

**Personal Information Management System** — a small, practical Web Forms CRUD project.
