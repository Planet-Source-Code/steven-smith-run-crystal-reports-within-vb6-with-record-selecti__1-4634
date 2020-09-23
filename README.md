<div align="center">

## Run Crystal Reports within VB6 with Record Selecti


</div>

### Description

THis short piece of code illustrates how to imbed Crystal Reports into VB. While simple, it does not seem to be documented well anywhere.
 
### More Info
 
If using record selection (printing a selected record) v_choice is a variable to be passed to the record selection part of Crystal

Crystal OXC from VB application is on Form

You will have a menu option or command button to launch report


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Steven Smith](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/steven-smith.md)
**Level**          |Unknown
**User Rating**    |4.0 (28 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/steven-smith-run-crystal-reports-within-vb6-with-record-selecti__1-4634/archive/master.zip)

### API Declarations

v_choice used for record selection (assign value to v_choice that matches field value)


### Source Code

```
'' if using record selection for report on siingle record
' set v_choice as public string
' store your record selction field choice to v_choice
'*******************************************
'Add the crystal ocx object to form (will be named CrystalReport1)
' you can pass record selection
''NOTE
''Create the report in Crystal first and place the report in the same directory as your database.
'' Set the report location to same as database in Crystal
''This part is run from menu or command button
CrystalReport1.ReportSource = crptReport
CrystalReport1.ReportFileName = reportpath & "\YOUR_REPORT_NAME.rpt"
'***This line only is using single record selection
CrystalReport1.ReplaceSelectionFormula ("{TABLENAME.FIELDNAME} =" & "'" & v_choice & "'")
'*********
CrystalReport1.WindowState = crptMaximized
CrystalReport1.PrintReport
CrystalReport1.PageZoom (50)
```

