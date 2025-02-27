📝 Overview

This VBA project enhances CCAM database searches, allowing users to filter results based on keywords and codes. The system highlights search results, applies modifiers to pricing, and enables users to select and copy relevant results to a separate sheet. It includes a user-friendly interface with checkboxes and action buttons.

🚀 Features

✔ Optimized Search – Filters CCAM data efficiently using regular expressions for valid codes.✔ Result Highlighting – Emphasizes matching keywords in the results sheet.✔ Modifier Management – Extracts and applies pricing modifications.✔ User Selection & Copying – Checkboxes enable selecting and copying results to a "Sélection" sheet.✔ Automatic Sorting – Sorts results by modified prices for better visibility.✔ Interactive UserForm – Allows manual modifier selection and price recalculations.

🔧 How It Works

1️⃣ Enter a keyword via an input box.2️⃣ The script searches the CCAM database for matches.3️⃣ Valid codes are checked using regex validation.4️⃣ If results are found, they appear in the "Résultats" sheet.5️⃣ The user can:

✅ Select results via checkboxes.

✅ Copy selected results to the "Sélection" sheet.

✅ Modify applied price modifiers.

✅ Sort results by modified price.
6️⃣ A UserForm provides an interface for manual modifier selection and price adjustments.

📌 Prerequisites

Ensure your Excel workbook contains the following sheets:

CCAM (Database)

Résultats (Search results)

Modifiers (Pricing adjustments)

Sélection (User-selected results)

Enable macros in Excel:File > Options > Trust Center > Macro Settings > Enable Macros

🔨 Installation

1️⃣ Import the VBA Code

Open Excel.

Press ALT + F11 to open the VBA Editor.

Import the .bas module or copy and paste the code.
2️⃣ Ensure Required Sheets Exist

Create the sheets CCAM, Résultats, Modifiers, Sélection.
3️⃣ Run the Macro

Execute RechercheOptimiseeAvecFiltrageParCode to start searching.

📂 Main Functions & Subroutines

🔍 Core Search & Filtering

RechercheOptimiseeAvecFiltrageParCode() – Performs keyword search and filters by code.

HighlightKeywords() – Highlights matching keywords in results.

SortPrixModifie() – Sorts results by modified price.

⚡ Modifier & Pricing Adjustments

ExtractModifiers() – Extracts relevant modifiers.

ExtractModPrice() – Calculates adjusted prices.

✅ User Selection & Actions

AddCheckboxes() – Adds checkboxes for selection.

CopySelectedResults() – Copies selected results to the Sélection sheet.

ApplyModifiers() – Opens the modifier selection form.

🎛 UserForm Actions

UserForm_Initialize() – Populates dropdown with available codes.

cmbCode_Change() – Updates modifier list based on selected code.

btnApply_Click() – Applies selected modifiers to pricing.

🏗 Future Enhancements

🔹 Logging system to track user selections.🔹 Enhanced UI for a smoother experience.🔹 Database integration for larger datasets.

👤 Author

Keenan Guiet

📜 License: MIT
