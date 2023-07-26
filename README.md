<h2 align="center">Excel: Scheduled Booking Management with VBA Macros and Formulas</h1>
</div>

### Spreadsheet Example
- [Booking Sheet.xlsm](https://github.com/Pwang0722/Excel_PerpetualCalendar_OutlookCalendar/raw/main/Booking%20Sheet%20Template.xlsm)

---

### Outline
A spreadsheet with a customized perpetual calendar that allows team members to make bookings for their duties each month and update specific bookings to the Outlook Calendar.
It involves utilizing multiple Functions, Conditional Formatting, and VBA Macros to efficiently achieve the desired objectives.

---

### Ｍethod 
- Fill in the data under columns A to M in the sheet titled "TITLE LIST". Based on the data you have filled in, a code will be generated from a formula in column N.
- There is a formula in cell B19 in the sheets titled from "1B. ###" to "13A. ###", which retrieves the codes from column N in the "TITLE LIST" sheet and automatically fills in the data based on different requirements in each sheet.

Formula example:
  ```bash
 =IFERROR(FILTER('TITLE LIST'!A:N,('TITLE LIST'!N:N="AENG FMALLN")+('TITLE LIST'!N:N="GMAND FMALLN")+('TITLE LIST'!N:N="OMAND FMALLN")+('TITLE LIST'!N:N="OBM FMALLN")+('TITLE LIST'!N:N="ASOT ONLYALLN")+('TITLE LIST'!N:N="GSOT ONLYALLN")+('TITLE LIST'!N:N="OSOT ONLYALLN")+('TITLE LIST'!N:N="AENG FM05BN")+('TITLE LIST'!N:N="GMAND FM05BN")+('TITLE LIST'!N:N="OMAND FM05BN")+('TITLE LIST'!N:N="OBM FM05BN")+('TITLE LIST'!N:N="ASOT ONLY05BN")+('TITLE LIST'!N:N="GSOT ONLY05BN")+('TITLE LIST'!N:N="OSOT ONLY05BN")+('TITLE LIST'!N:N="GMAND FMALLY")+('TITLE LIST'!N:N="GSOT ONLYALLY")+('TITLE LIST'!N:N="GMAND FM05BY")+('TITLE LIST'!N:N="GSOT ONLY05BY")),"")
  ```

---
