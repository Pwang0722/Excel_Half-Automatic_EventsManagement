<h2 align="center">Excel: Scheduled Booking Management with VBA Macros and Formulas</h1>
</div>

### Spreadsheet Example
- [Booking Sheet.xlsm (password:1234)](https://github.com/Pwang0722/Excel_PerpetualCalendar_OutlookCalendar/raw/main/Booking%20Sheet%20Template.xlsm)

---

### Outline
A spreadsheet with a customized perpetual calendar that allows team members to make bookings for their duties each month and update specific bookings to the Outlook Calendar. It involves utilizing multiple Functions, Conditional Formatting, and VBA Macros to efficiently achieve the desired objectives. And created Formulas and VBA Macros with the help of ChatGPT


---

### Ｍethod 
- To Avoid
Formula example:
  ```bash
 =IFERROR(FILTER('TITLE LIST'!A:N,('TITLE LIST'!N:N="AENG FMALLN")+('TITLE LIST'!N:N="GMAND FMALLN")+('TITLE LIST'!N:N="OMAND FMALLN")+('TITLE LIST'!N:N="OBM FMALLN")+('TITLE LIST'!N:N="ASOT ONLYALLN")+('TITLE LIST'!N:N="GSOT ONLYALLN")+('TITLE LIST'!N:N="OSOT ONLYALLN")+('TITLE LIST'!N:N="AENG FM05BN")+('TITLE LIST'!N:N="GMAND FM05BN")+('TITLE LIST'!N:N="OMAND FM05BN")+('TITLE LIST'!N:N="OBM FM05BN")+('TITLE LIST'!N:N="ASOT ONLY05BN")+('TITLE LIST'!N:N="GSOT ONLY05BN")+('TITLE LIST'!N:N="OSOT ONLY05BN")+('TITLE LIST'!N:N="GMAND FMALLY")+('TITLE LIST'!N:N="GSOT ONLYALLY")+('TITLE LIST'!N:N="GMAND FM05BY")+('TITLE LIST'!N:N="GSOT ONLY05BY")),"")
  ```

---
