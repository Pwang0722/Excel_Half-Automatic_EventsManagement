<h2 align="center">Excel: Efficient Scheduled Booking Management with VBA Macros and Formulas</h1>
</div>

### Spreadsheet Examples
- [Asia Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(Asia%20Content)%20Clean%20Template.xlsm)
- [English Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(English%20Content)%20Clean%20Template.xlsm)
---

### Outline
An example of using the FILTER function to auto-fill similar inputs across multiple worksheets, and of reformatting the worksheets and exporting them as PDF files using VBA Macros. And I have created all these VBA macros with the help of ChatGPT.

---

### Ｍethod 
Take spreadsheet [Asia Content.xlsm](https://github.com/Pwang0722/Excel_DataManagement/raw/main/(Asia%20Content)%20Clean%20Template.xlsm) as an example:
- Fill in the data under columns A to M in the sheet titled "TITLE LIST". Based on the data you have filled in, a code will be generated from a formula in column N.
- There is a formula in cell B19 in the sheets titled from "1B. ###" to "13A. ###", which retrieves the codes from column N in the "TITLE LIST" sheet and automatically fills in the data based on different requirements in each sheet.

Formula example:
  ```bash
 =IFERROR(FILTER('TITLE LIST'!A:N,('TITLE LIST'!N:N="AENG FMALLN")+('TITLE LIST'!N:N="GMAND FMALLN")+('TITLE LIST'!N:N="OMAND FMALLN")+('TITLE LIST'!N:N="OBM FMALLN")+('TITLE LIST'!N:N="ASOT ONLYALLN")+('TITLE LIST'!N:N="GSOT ONLYALLN")+('TITLE LIST'!N:N="OSOT ONLYALLN")+('TITLE LIST'!N:N="AENG FM05BN")+('TITLE LIST'!N:N="GMAND FM05BN")+('TITLE LIST'!N:N="OMAND FM05BN")+('TITLE LIST'!N:N="OBM FM05BN")+('TITLE LIST'!N:N="ASOT ONLY05BN")+('TITLE LIST'!N:N="GSOT ONLY05BN")+('TITLE LIST'!N:N="OSOT ONLY05BN")+('TITLE LIST'!N:N="GMAND FMALLY")+('TITLE LIST'!N:N="GSOT ONLYALLY")+('TITLE LIST'!N:N="GMAND FM05BY")+('TITLE LIST'!N:N="GSOT ONLY05BY")),"")
  ```
 - To prevent Excel from lagging while filling in data, I set the Calculation Options to Manual. Hence, users have to run the Calculate Now function every time they finish data entry or make changes. To make this process more convenient for everyone, I created a Macro that runs the Calculate Now function and assigned it to a button.

Macro example for function, Calculation Now:
  ```bash
  Sub CalculateWorkbook()
    Application.CalculateFull
    MsgBox ("Done. Sheets are ready to check.")
End Sub
```
- Running the Macro to tidy up multiple sheets, including hiding and deleting unnecessary columns, rows, and data.

Macro example for reformatting multiple worksheets:
  ```bash
  Sub Reformat()
Application.ScreenUpdating = False
For Each sh In Worksheets
If sh.Name <> "INDEX" And sh.Name <> "TITLE LIST" And sh.Name <> "REFORMAT" And sh.Name <> "#" And sh.Name <> "##" Then
    sh.Activate
Dim lRow As Long
Dim iCntr As Long
lRow = 45
For iCntr = lRow To 19 Step -1
    If Trim(Cells(iCntr, 1)) = "" Then
        Rows(iCntr).Delete
    End If
Next
    Range("A19:L45").Select
    Selection.Value = Selection.Value
    Range("K3:L4").Select
    Selection.ClearContents
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("4:16").Select
    Selection.Delete Shift:=xlUp
    x = ActiveSheet.UsedRange.Rows.Count
    Columns("A").Hidden = True
    Columns("G").Hidden = True
End If
Next sh
Application.ScreenUpdating = True
MsgBox ("Done! Sheets are ready to check.")
End Sub
```
- Lastly, running the Macro to export multiple sheets as separate PDF files.

Macro example for exporting PDF files:
  ```bash
  Sub AllSavePDF()
Dim fileName As String
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    If ws.Name <> "TITLE LIST" And ws.Name <> "INDEX" And ws.Name <> "##" And ws.Name <> "#" And ws.Name <> "REFORMAT" Then
    
        'Get filename from cell A2
        fileName = ws.Range("A2").Value
        
        'Add date to the filename
        fileName = fileName & "_" & Format(Date, "mm-dd-yy")
        
        'Save as PDF file
        ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
            fileName, Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    End If
Next ws
MsgBox ("Done saving ALL sheets as PDF files!")
End Sub
```

- In addition, to avoid users forgetting to run Calculation Now after data entry, I have created a reminder that will pop out every time they close the workbook.

```bash
  Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim Answer As Long
    Dim textBreakLine As String
    Dim textOne As String
    Dim textTwo As String
 
    textBreakLine = "*Producer reminder to run formula*'"
    textOne = "Yes: save & close"
    textTwo = "No: back"
    textThree = "Cancel: close without saving"
    Answer = MsgBox(textBreakLine & vbCrLf & textOne & vbCrLf & textTwo & vbCrLf & textTree, vbQuestion + vbYesNoCancel, "Close Workbook")
    Select Case Answer
        Case vbYes
            ActiveWorkbook.Save
        Case vbNo
            Cancel = True
            ThisWorkbook.Activate
    End Select
End Sub
```
---

### Tutorials
For data entry:
- [Desktop User](https://ior.ad/9aZG?iframeHash=mobilequick-1)
- [Web User](https://ior.ad/9aZm?iframeHash=mobilequick-1)

For worksheet reformatting and exporting PDF files:
- [Mac User](https://ior.ad/9b99)
- [Windows User](https://github.com/Pwang0722/Excel_DataManagement/blob/main/Steps%20(for%20Win).pdf)
---
