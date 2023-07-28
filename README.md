<h2 align="center">Excel: Events Schedules Management with VBA macros and Formulas</h1>
</div>

### Spreadsheet Example
- [Booking Sheet (password:1234).xlsm](https://github.com/Pwang0722/Excel_PerpetualCalendar_OutlookCalendar/raw/main/Booking%20Sheet%20Template%20(password-1234).xlsm)

---

### Outline
A spreadsheet with a customized perpetual calendar that allows team members to make bookings for their duties each month and update specific bookings as events to the Outlook Calendar. It involves utilizing multiple Functions, Conditional Formatting, and VBA macros to efficiently achieve the desired objectives. And created Formulas and VBA macros with the help of ChatGPT

---

### Notice
- The spreadsheet was created using Excel version 2306 on Windows 11. It may encounter unexpected errors while running VBA macros on a MAC.
- Before testing out the spreadsheet, make sure to enable VBA macros and check the Outlook reference.
  - Enable Macros: In Excel, go to File > Options > Trust Center > Trust Center Settings. Under the Macro Settings tab, select "Enable all macros" or "Enable all macros with notification" to allow the code to run.
  - Check References: In the VBA editor, go to Tools > References and ensure that the necessary Outlook reference is selected. Look for the reference starting with "Microsoft Outlook Object Library" and make sure it is checked.
  
---
### Sheet Protection 
- To avoid accidentally modifying the template worksheets, they will be protected with a password every time the workbook is opened or closed.
  
Macro example for protecting sheets:
  ```bash
Private Sub Workbook_Open()
    ProtectSheets
End Sub
---
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ProtectSheets
End Sub
---
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Dim protectedSheets As Variant
    Dim ws As Worksheet
    Dim password As String
    
    ' List of sheets to protect
    protectedSheets = Array("TEMPLATE_ALL", "Audio Out-House", "Summary", "HOLIDAYS")
    
    ' Check if the changed sheet is one of the protected sheets
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(Target.Worksheet.Name)
    On Error GoTo 0
    
    If Not ws Is Nothing And IsInArray(ws.Name, protectedSheets) Then
        ' Check if the sheet is protected
        If ws.ProtectContents Then
            ' Prompt the user to enter the password to unprotect the sheet
            password = InputBox("Enter the password to unprotect the sheet:", "Password")
            
            ' Check if the entered password matches the preset password
            If password = "1234" Then
                ' Unprotect the sheet to allow editing
                ws.Unprotect password:="1234"
            Else
                MsgBox "Incorrect password. The sheet will remain protected.", vbExclamation
                Application.EnableEvents = False
                Target.Offset(1, 0).Select ' Move to the next cell to avoid an infinite loop
                Application.EnableEvents = True
            End If
        End If
    End If
End Sub
---
Private Sub ProtectSheets()
    Dim protectedSheets As Variant
    Dim ws As Worksheet
    
    ' List of sheets to protect
    protectedSheets = Array("TEMPLATE_ALL", "Audio Out-House", "Summary", "HOLIDAYS")
    
    ' Loop through each protected sheet and protect with the preset password
    For Each sheetName In protectedSheets
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ws.Protect password:="1234", UserInterfaceOnly:=True
        End If
    Next sheetName
End Sub
---
Function IsInArray(val As Variant, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, val)) > -1)
End Function
  ```
---

### Perpetual Calendar_sheets 'TEMPLATE_ALL' & 'HOLIDAYS'
- Sheet 'TEMPLATE_ALL' is a booking sheet contains a perpetual calendar which could generate a clean monthly calendar for team members to make bookings for their projects.
- Generate a clean monthly calendar by select desired Year and Month in cells AN2 & AN5.

Macro example for generating calendar:
  ```bash
Sub GenerateBookingSheet()
    NewSheet = Range("B1").Text & " " & Range("AM1")
    Sheets("TEMPLATE_ALL").Copy Before:=Sheets(2)
    ActiveSheet.Name = NewSheet
    ActiveSheet.Select
    Range("B1:AJ79").Select
    Selection.Copy
    Range("B1:AJ79").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("AL:AV").Hidden = True
    Range("B1").Select
    MsgBox "Done"
End Sub
  ```
- Calendar is designed with 6 weeks from Monday to Friday, using DATE and WEEKDAY functions to generate the dates based on the Year and Month input in cells AN2 & AN5.
- Names of holidays are generated using the VLOOKUP function to lookup the data in sheet 'HOLIDAYS'.
  
Formula example for dates (H8):
  ```bash
=DATE(AN2,AN5,1)-WEEKDAY(DATE(AN2,AN5,1),2)+1
  ```
Formula example for holidays (B8):
  ```bash
=IFERROR(VLOOKUP($H$20,HOLIDAYS!$B$2:$D$22,3,FALSE),"")
  ```
- Use Conditional Formatting with a formula to format cells' color in grey for those dates that don't belong to the selected month.
- And formatting cells' color in magenta for holidays by using the MATCH function to match the data in sheet 'HOLIDAYS'.

Formula example for Conditional Formatting for dates:
  ```bash
=MONTH($H$8)<>$AN$5
 ```
Formula example for Conditional Formatting for holidays:
  ```bash
=MATCH($H$8,HOLIDAYS!$B$2:$B$22,0)
 ```

---

### Update events to Outlook Calendar_Sheet 'Audio Out-House'
- List all Freelance Audio items in the 'TEMPLATE_ALL' sheet under the FACILITY column in columns E, L, S, Z, and AG, and place the data in the 'Audio Out-House' sheet table B4:I35.

Macro example for listing out the data:
  ```bash
Sub ListOutFreelanceAudio()

    Range("B4:I35").Select
    Selection.ClearContents
    Range("A4:A35").EntireRow.AutoFit
    
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(2)
    
    Dim destSheet As Worksheet
    Set destSheet = ThisWorkbook.Sheets("Audio Out-House")
    
    Dim rowCounter As Long
    rowCounter = 4
    
    Dim dateValue As Date
    Dim producerValue As String
    Dim channelValue As String
    Dim projectValue As String
    Dim facilityValue As String
    Dim timeValue As Date
    Dim durationValue As String
    Dim dueDateValue As Date
    
    Dim rangeList() As String
    rangeList = Split("B8:H13,I8:O13,P8:V13,W8:AC13,AD8:AJ13,B20:H25,I20:O25,P20:V25,W20:AC25,AD20:AJ25,B32:H37,I32:O37,P32:V37,W32:AC37,AD32:AJ37,B44:H49,I44:O49,P44:V49,W44:AC49,AD44:AJ49,B56:H61,I56:O61,P56:V61,W56:AC61,AD56:AJ61,B68:H73,I68:O73,P68:V73,W68:AC73,AD68:AJ73", ",")
    
    Dim rangeObj As Range
    
    For Each rangeStr In rangeList
        
        Set rangeObj = sourceSheet.Range(rangeStr)
        dateValue = rangeObj.Cells(1, rangeObj.Columns.Count).Value
        
        For j = 0 To 4
            
            producerValue = rangeObj.Cells(2 + j, 1).Value
            channelValue = rangeObj.Cells(2 + j, 2).Value
            projectValue = rangeObj.Cells(2 + j, 3).Value
            facilityValue = rangeObj.Cells(2 + j, 4).Value
            textValue = rangeObj.Cells(2 + j, 5).Value
            durationValue = rangeObj.Cells(2 + j, 6).Value
            dueDateValue = rangeObj.Cells(2 + j, rangeObj.Columns.Count).Value
            
            If facilityValue = "Freelance Audio" Then
                destSheet.Cells(rowCounter, 2).Value = dateValue
                destSheet.Cells(rowCounter, 3).Value = producerValue
                destSheet.Cells(rowCounter, 4).Value = channelValue
                destSheet.Cells(rowCounter, 5).Value = projectValue
                destSheet.Cells(rowCounter, 6).Value = facilityValue
                destSheet.Cells(rowCounter, 7).Value = textValue
                destSheet.Cells(rowCounter, 8).Value = durationValue
                destSheet.Cells(rowCounter, 9).Value = dueDateValue
                rowCounter = rowCounter + 1
            End If
            
        Next j
        
    Next rangeStr
    destSheet.Range("B4:I" & rowCounter - 1).Sort key1:=destSheet.Range("I4:I" & rowCounter - 1), _
                                                    order1:=xlAscending, Header:=xlNo

End Sub
```
- The data from table B4:I35 will be automatically reformatted into Outlook acceptable format for events and placed in table K4:O35 using multiple formulas.
- The events listed in table K4:O35 could be uploaded to 'Outsource Audio' calendar in Outlook.

Macro example for updating events to Outlook Canlendar :
  ```bash
Sub UpdateOusourceAudioCalendar()
    Dim olApp As Outlook.Application
    Dim olApt As Outlook.AppointmentItem
    Dim olCalFolder As Outlook.MAPIFolder
    Dim i As Long

    Set olApp = CreateObject("Outlook.Application")
    Set olCalFolder = olApp.Session.GetDefaultFolder(olFolderCalendar).Folders("Outsource Audio")

    ' Loop through all rows in range K4:O35
    For i = 4 To 35
        If Sheets("Audio Out-House").Range("K" & i).Value <> "" Then ' Check if row has a title

            ' Create new appointment item
            Set olApt = olCalFolder.Items.Add(olAppointmentItem)

            ' Set appointment properties
            olApt.Subject = Sheets("Audio Out-House").Range("K" & i).Value ' Title
            olApt.location = Sheets("Audio Out-House").Range("N" & i).Value ' Location
            olApt.Categories = Sheets("Audio Out-House").Range("O" & i).Value ' Color Category
            olApt.Body = "" ' Body

            ' Set start and end date/time
            If IsDate(Sheets("Audio Out-House").Range("L" & i).Value) And IsDate(Sheets("Audio Out-House").Range("M" & i).Value) Then
                olApt.Start = Sheets("Audio Out-House").Range("L" & i).Value
                olApt.End = Sheets("Audio Out-House").Range("M" & i).Value
            End If

            ' Save appointment
            olApt.Save
        End If
    Next i

    ' Clean up
    Set olApt = Nothing
    Set olCalFolder = Nothing
    Set olApp = Nothing
    MsgBox "Done"
End Sub
```
- The events in 'Outsource Audio' calendar also can be deleted by selecting desired Year and Month in cells R9 & R11.

Macro example for deleting events in Outlook Calendar.:
  ```bash
Sub DeleteEventsInMonth()
    Dim olApp As Object
    Dim olCalFolder As Object
    Dim olItems As Object
    Dim i As Long
    Dim month As Integer
    Dim year As Integer
    Dim startDate As Date
    Dim endDate As Date
    
    ' Get the month and year from the worksheet
    month = Worksheets("Audio Out-House").Range("R11").Value
    year = Worksheets("Audio Out-House").Range("R9").Value
    
    ' Set the start and end dates for the selected month
    startDate = DateSerial(year, month, 1)
    endDate = DateSerial(year, month + 1, 1) - 1
    
    ' Get the Outlook calendar folder
    Set olApp = CreateObject("Outlook.Application")
    Set olCalFolder = olApp.Session.GetDefaultFolder(9).Folders("Outsource Audio")
    
    ' Use Restrict method to only retrieve calendar items within the selected month and year
    Set olItems = olCalFolder.Items.Restrict("[Start] >= '" & startDate & "' AND [End] <= '" & endDate & "'")
    
    ' Loop through each event in the calendar folder and delete it
    For i = olItems.Count To 1 Step -1
        olItems.Item(i).Delete
    Next i
    
    ' Display a message box to indicate that the events were deleted successfully
    MsgBox "Done"
End Sub
```


---
