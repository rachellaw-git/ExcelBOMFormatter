# Importing A Global Macro in Excel

This guide will help you import a **global macro** in Excel, which can be used in any workbook.

---

## Step 1: Open the Visual Basic for Applications (VBA) Editor

1. Open Excel.
2. Press `Alt + F11` to open the **VBA Editor**.
3. In the VBA Editor, locate the **Project Explorer** on the left.  
   - If you don't see it, press `Ctrl + R` to open it.

---

## Step 2: Open Your Personal Macro Workbook

Excel uses a hidden workbook called `PERSONAL.XLSB` to store global macros.

1. In the VBA Editor, look for a project named `VBAProject (PERSONAL.XLSB)`.
2. If it does **not exist**, you need to create it by recording a macro:
   1. Go back to Excel.
   2. Go to **View → Macros → Record Macro**
   ![record_macro.png](assets/record_macro.png)
   3. In the **Record Macro** dialog:
      - Set **Store macro in** to `Personal Macro Workbook`.
      ![select_personal_macro_workbook.png](assets/select_personal_macro_workbook.png)
   4. Click **OK**, then **Stop Recording**.
   ![stop_recording.png](assets/stop_recording.png)
   5. This will create the `PERSONAL.XLSB` workbook if it wasn’t already there.

---

## Step 3: Import A Macro

1. Press `Alt + F11` to open the **VBA Editor**
2. Right-click VBAProject (PERSONAL.XSLB) → Import File...
	1. Select the file to import (like a .bas file such as `FormatBomForCadDatalink.bas`)
     ![[import_file2.png]]
     2. You can double-click the created module to access the script
     3. **Save** 
![img.png](assets/imported_bas.png)

---

💡 **Tip:** You may run into an error relating to a hidden workbook  
 `PERSONAL.XLSB` is hidden by default. Do not delete or rename it, or your global macros will be lost.
- To view it: In **Excel**, **View → Unhide → select PERSONAL.XLSB**.

## Step 4: Assign a Shortcut to Your Global Macro

1. Go back to **Excel** (not the VBA Editor).  
2. Go to **View → Macros → View Macros** (or press `Alt + F8`).  
3. In the **Macro** dialog:
   - Set the **Macros in** dropdown to `All Open Workbooks` or `PERSONAL.XLSB`.
   - Select the macro you imported.  
4. Click **Options…**.
5. In the **Macro Options** dialog:
   - Enter a **shortcut key** (e.g., type `q` for `Ctrl + Q`).
6. Click **OK**, then **Close** the Macro dialog.

---

## Step 5: Test Your Macro

1. Open any Excel workbook.  
2. Press your assigned shortcut (e.g., `Ctrl + Q`).  
3. The macro should run, no matter which workbook is active.

---
# Updating the Macro
You can update the macro either by 
1) Importing the new script and deleting the old one
2) Overwriting the old script
	1) Selecting all the code and deleting it
	2) Paste in the new code

---
## Option 1: Import and Delete a Macro
Follow these steps for importing the new script and deleting the old one
1. **Open** the VBA Editor Project Explorer
2. **Right-click** VBAProject (PERSONAL.XSLB) → **Import File...**
	1. **Select** the file to import such as `FormatBomForCadDatalink.bas`
   ![[import_file2.png]]
3. Assign a Shortcut to Your Global Macro (refer to earlier section)
### Removing a Macro
Remove the old macro
1. **Open** the VBA Editor Project Explorer
2. **Right-click** the module you want to remove
3. **Select** "Remove \[module_name]"
   ![[remove_module.png]]

---

### Option 2: Overwriting the old script
1. **Open** the VBA Editor Project Explorer
2. **Double-click** the module to be updated
   ![[vba_editor_overwrite.png]]
3. **Delete** the current code. **Copy and paste** the new script in. (You can open the script file in a text editor such as Notepad)

---
v2.0.0