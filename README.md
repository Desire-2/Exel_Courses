### Excel VBA Duplicate Finder and Cleaner

This solution provides two complementary VBA macros for managing duplicates in Excel workbooks. Below is the documentation and user guide.

---

### **Macro 1: Find and Report Duplicates**  
`FindAndReportDuplicatesInColumnB_FromRow9`  
*Scans Column B (from row 9) across all sheets, highlights duplicates, and generates a report.*

#### **Features:**
1. **Case Sensitivity Option**  
   - Asks whether to perform case-sensitive checks
   - Non-case-sensitive: Converts text to lowercase for comparison

2. **Dynamic Scanning**  
   - Scans all sheets except the "Duplicate Report"
   - Starts at **Column B, Row 9** (ignores rows 1-8)
   - Skips empty cells and errors

3. **Reporting**  
   - Creates "Duplicate Report" sheet with:
     - Duplicate value
     - Occurrence count
     - Source sheet name
     - Cell address
   - Formats report with headers and auto-fits columns

4. **Visual Highlighting**  
   - Duplicates highlighted with:
     - Light red fill (`RGB(255, 199, 206)`)
     - Dark red text (`RGB(156, 0, 6)`)

5. **Performance Optimized**  
   - Disables screen updating/events during execution
   - Reports total duplicates and processing time

---

### **Macro 2: Clear Duplicate Highlights**  
`ClearOnlyDuplicateHighlights`  
*Removes highlighting applied by the first macro.*

#### **Features:**
1. **Precise Clearing**  
   - Only removes cells with *exact* light red fill (`RGB(255, 199, 206)`)
   - Resets font color to automatic
   - Skips "Duplicate Report" sheet

2. **Non-Destructive**  
   - Preserves other cell formatting
   - Doesn’t delete data or reports

---

### **User Guide**  
#### **How to Use:**
1. **Add Macros to Excel:**  
   - Press `Alt + F11` to open VBA Editor  
   - Insert new module(s) and paste both macros  

2. **Run Duplicate Check:**  
   - Execute `FindAndReportDuplicatesInColumnB_FromRow9`  
   - Choose case sensitivity when prompted  
   - Check the "Duplicate Report" sheet for results  

3. **Clear Highlights:**  
   - Execute `ClearOnlyDuplicateHighlights`  
   - Confirm completion via popup  

#### **Keyboard Shortcuts:**  
- Run Macro: `Alt + F8` → Select macro → **Run**  

#### **Important Notes:**  
- **Column B Focus:** Only scans Column B (adjust code to change column)  
- **Row 9+:** Ignores rows 1-8  
- **Highlight Safety:** Clear macro ONLY removes the exact highlight color from the finder macro  
- **Report Sheet:** Auto-deletes/recreates report on each run  

---

### **Troubleshooting**  
| Issue                     | Solution                                  |
|---------------------------|-------------------------------------------|
| Macros disabled           | Enable macros in `Trust Center Settings`  |
| Report not generated      | Check sheet name ≠ "Duplicate Report"     |
| Highlights not cleared    | Ensure exact RGB color exists in cells    |
| Slow execution            | Close other apps; reduce workbook size    |

---

### **Code Files**  
#### 1. `FindAndReportDuplicatesInColumnB_FromRow9`  
```vba
[VBA code from first file]
```

#### 2. `ClearOnlyDuplicateHighlights`  
```vba
[VBA code from second file]
```

---

**Tip:** Customize the highlight colors by modifying the `RGB` values in both macros to ensure consistent cleanup.
