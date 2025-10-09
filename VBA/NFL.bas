Attribute VB_Name = "NFL"
Option Explicit
Private conn As New Connection
Sub getConnection()

conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & ActiveWorkbook.FullName & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"

End Sub
Private Sub freezeTopPane(activeWindow As Window)
    With activeWindow 'ActiveWindow
        .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub
Sub saveXLSX()
Dim i As Integer: i = 0
Dim fileName As String
Dim sourceWorkbook As Workbook
Dim currentWorkbook As Workbook
Dim xl As New Excel.Application
Dim r As Range
On Error GoTo errorHandler

fileName = ActiveWorkbook.path & "\" & Dir(ActiveWorkbook.path & "\FanDuel*.csv")

xl.Workbooks.Open (fileName)

xl.ActiveWorkbook.SaveAs Replace(fileName, "csv", "xlsx"), FileFormat:=xlOpenXMLWorkbook

With xl.ActiveSheet.Sort
    .SortFields.Clear
    .SortFields.Add key:=xl.Worksheets(1).Cells.Find(What:="Team")
    .SortFields.Add key:=xl.Worksheets(1).Cells.Find(What:="Position")
    .SortFields.Add key:=xl.Worksheets(1).Cells.Find(What:="Salary"), Order:=xlDescending
    .SetRange xl.Range("A1:Q" & xl.Cells(Rows.Count, 1).End(xlUp).row)
    .header = xlYes
    .Apply
End With

If xl.Cells(1, 16).Value <> "Points" Then xl.Cells(1, 16).Value = "Points"
If xl.Cells(1, 17).Value <> "Projected Points" Then xl.Cells(1, 17).Value = "Projected Points"

For Each r In xl.Range("F2:F" & xl.Cells(Rows.Count, 1).End(xlUp).row)
    r.Offset(, 11).Value = r.Value2
Next

xl.Range("A1").CurrentRegion.EntireColumn.AutoFit

freezeTopPane xl.activeWindow
xl.Sheets(1).Activate

xl.ActiveWorkbook.Save
xl.Visible = True


Set xl = Nothing
Exit Sub

errorHandler:
    xl.Workbooks(Dir(ActiveWorkbook.path & "\FanDuel*.csv")).Close
    xl.Workbooks.Open Replace(fileName, "csv", "xlsx")
    xl.Visible = True
    
    Set xl = Nothing

End Sub
Sub createTierWB()
Dim fileName As String: fileName = "Tier.xlsm"
    
If Dir(ActiveWorkbook.path & "\" & fileName) = "" Then
    Workbooks.Add.SaveAs fileName:=ActiveWorkbook.path & "\" & fileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    If Worksheets.Count < 9 Then Worksheets.Add Count:=9 - Worksheets.Count
    ActiveWorkbook.Save
Else
    Workbooks.Open (ActiveWorkbook.path & "\" & fileName)
End If

attachModule
createFanduel
createStats
createTier
createSearch
createMatrix
createRandomLineup
createEventProcedure

With Worksheets("Search")
    .Activate
    .Range("C2").Activate
End With

ActiveWorkbook.Save

End Sub
Sub createScoreProjections()
Dim SheetName As String: SheetName = "Projections"
Dim fileName As String: fileName = "Scoring Projections.xlsm"
Dim fileExists As Boolean: fileExists = True
Dim spArr
Dim i As Double
    
If Dir(ActiveWorkbook.path & "\" & fileName) = "" Then
    fileExists = False
    Workbooks.Add.SaveAs fileName:=ActiveWorkbook.path & "\" & fileName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    
End If
    
createWSProjections fileName, SheetName, "D"
    

'End If

'populateWorksheet fileName, fileExists, sheetName

End Sub
Sub createRetro()
Dim wb As Workbook
Dim fileName As String: fileName = "Tier.xlsm"
Dim ws As Worksheet
Dim conn As New ADODB.Connection
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim arr
Dim coll As New Collection
Dim printArr
Dim strArr() As String
Dim i As Long
Dim j As Long

Set wb = Workbooks.Open(ActiveWorkbook.path & "\" & fileName)

Set ws = Worksheets(7)
With ws
    .Activate
    .Name = "Retro"
End With

ws.Range("A1:W" & ws.Cells(Rows.Count, 1).End(xlDown).row).Clear
createHeaders ws.Name

'getPlayerArray
SQL = "SELECT * " & _
      "FROM [FanDuel$]" & _
      "WHERE [Points] > 0 " & _
      "ORDER BY [Points] DESC"
          
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & wb.FullName & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"

rs.Open SQL, conn

If rs.EOF Then Exit Sub


ws.Range("A2").CopyFromRecordset rs
arr = ws.Range("A2:S" & ws.Cells(Rows.Count, 1).End(xlUp).row)
ws.Range("A2:S" & ws.Cells(Rows.Count, 1).End(xlUp).row).ClearContents

'Process Data
Set coll = getCollection(arr, 6)

ReDim printArr(1 To coll.Count, 22)

For i = 1 To coll.Count
    strArr = Split(coll(i), "_")
    
    For j = 0 To UBound(strArr)
        printArr(i, j) = strArr(j)
    Next j
Next i

'Process Array
With ws
    Range("A1").Activate
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    
    If Application.CountA(.Range("W:W")) > 1 Then
        .Range("A2:W" & Application.CountA(.Range("W:W"))).ClearContents
    End If
    
    'createHeaders wb, .Name
    .Range("A2:W" & UBound(printArr) + 1).Value = printArr
    .Range("A1").CurrentRegion.EntireColumn.AutoFit
    freezeTopPane activeWindow
    
    If .AutoFilterMode = False Then .Range("A1").AutoFilter
    .Range("A1").CurrentRegion.EntireColumn.AutoFit

    With .Sort
        .SortFields.Clear
        .SortFields.Add key:=Range("Q2"), Order:=xlDescending
        .SetRange Range("A2:W" & UBound(printArr) + 1)
        .Apply
    End With

End With

With ws
    .Activate
    .Range("A2").Activate
End With

ActiveWorkbook.Save

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub
Private Sub createMatrix()
'Dim wb As Workbook
Dim ws As Worksheet
Dim rs As New Recordset
Dim SQL As String
Dim arr
Dim i As Long
Dim j As Long

Set ws = Worksheets(2)
With ws
    .Activate
    .Name = "Matrix"
End With

createHeaders ws.Name
ws.Range("A2:L" & ws.Cells(Rows.Count, 1).End(xlDown).row).ClearContents

getConnection

SQL = "SELECT [Nickname], [Position]&':'&[Team] &' '&[Injury Indicator] " & _
      "FROM [FanDuel$] " & _
      "WHERE [Tier] = 1 " & _
      "ORDER BY [Salary] DESC"

rs.Open SQL, conn
'arr = wb.Worksheets("Search").Range("B1:B" & wb.Worksheets("Search").Cells(Rows.Count, 1).End(xlUp).row).Value
arr = Application.Transpose(rs.GetRows)
ReDim Preserve arr(1 To UBound(arr), 1 To 9)

With ws
    'Set player matrix
    'createHeaders wb, ws.Name
    For i = 1 To UBound(arr)
        arr(i, 3) = "=COUNTIFS(Tier!F:F,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 4) = "=COUNTIFS(Tier!G:G,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 5) = "=COUNTIFS(Tier!H:H,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 6) = "=COUNTIFS(Tier!I:I,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 7) = "=COUNTIFS(Tier!J:J,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 8) = "=COUNTIFS(Tier!K:K,$B$" & i + 1 & ",Tier!$L:$L,"">0"")"
        arr(i, 9) = "=SUM(C" & i + 1 & ":H" & i + 1 & ")"
    Next i

    .Range("A2").Resize(UBound(arr), UBound(arr, 2)).Value = arr
    .Cells(i + 1, 1).Value = "Totals"
    .Cells(i + 1, 9).Formula = "=SUM(I2:I" & i & ")"
    
    'Set key matrix
    .Cells(2, 11).Value = 12
    .Cells(3, 11).Value = 13
    .Cells(4, 11).Value = 14
    .Cells(5, 11).Value = 23
    .Cells(6, 11).Value = 24
    .Cells(2, 12).FormulaArray = "=COUNTIFS(Tier!C:C,$J$2" & ",Tier!$L:$L,1)"
    .Cells(3, 12).FormulaArray = "=COUNTIFS(Tier!C:C,$J$3" & ",Tier!$L:$L,1)"
    .Cells(4, 12).FormulaArray = "=COUNTIFS(Tier!C:C,$J$4" & ",Tier!$L:$L,1)"
    .Cells(5, 12).FormulaArray = "=COUNTIFS(Tier!C:C,$J$5" & ",Tier!$L:$L,1)"
    .Cells(6, 12).FormulaArray = "=COUNTIFS(Tier!C:C,$J$6" & ",Tier!$L:$L,1)"
End With

freezeTopPane activeWindow
Range("A1:L1").CurrentRegion.EntireColumn.AutoFit

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing


End Sub
Private Sub createStats()
'Dim wb As Workbook
Dim ws As Worksheet
'Dim conn As New ADODB.Connection '
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim dict As New Scripting.Dictionary
Dim team
Dim cmd As New ADODB.Command
Dim param As ADODB.parameter
Dim i As Long
Dim data
Dim j As Long
Dim arr

Set ws = Worksheets(4)
With ws
    .Activate
    .Name = "Stats"
End With

'Import Player Positions for Dictionary
SQL = "SELECT [Nickname], [Position] & ':' & [Team] & ' ' & [Injury Indicator] " & _
      "FROM [FanDuel$] " & _
      "WHERE [Tier] = 1"

'conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
'          "Data Source=" & wb.FullName & ";" & _
'          "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"
getConnection
rs.Open SQL, conn

'Create Player/Position Dictionary
Do Until rs.EOF
    dict.Add key:=rs.fields(0).Value, Item:=rs.fields(1).Value
    rs.MoveNext
Loop
rs.Close

'Get Teams
SQL = "SELECT DISTINCT [Team] " & _
      "FROM [FanDuel$] "
   
rs.Open SQL, conn

If Not rs.EOF Then
    team = rs.GetRows()
End If

rs.Close
conn.Close

'Import player stats
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=C:\UCF_Challenges\NFL-Fantasy\2025\;" & _
          "Extended Properties=""text;HDR=YES;FMT=Delimited"";"

Set cmd.ActiveConnection = conn
cmd.CommandText = "SELECT * " & _
                  "FROM FantasyPros_Fantasy_Football_Points_HALF.csv " & _
                  "WHERE team IN (?,?) " & _
                  "ORDER BY 4, cint([#])"
      
Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, team(0, 0))
cmd.Parameters.Append param
Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, team(0, 1))
cmd.Parameters.Append param
      
'Set rs = cmd.Execute
rs.Open cmd

'Populate Stats Sheet - Headers
For i = 0 To rs.fields.Count - 1
    Sheets("Stats").Cells(1, i + 1).Value = rs.fields(i).Name
Next

'Populate Stats Sheet - Data
data = rs.GetRows()
With Sheets("Stats")
'    .Range("A2").CopyFromRecordset rs '-- Pasted numerics as text
    For i = 0 To UBound(data, 2)
        For j = 0 To UBound(data, 1)
            If Not data(j, i) = "-" Then .Cells(i + 2, j + 1).Value = data(j, i)
        Next j
    Next i
    arr = .Range("B2:B" & .Cells(Rows.Count, 1).End(xlUp).row).Value
    For i = LBound(arr) To UBound(arr)
        If dict.Exists(arr(i, 1)) Then .Range("C" & i + 1).Value = dict.Item(arr(i, 1))
    Next
'    .Range("A1").CurrentRegion.EntireColumn.AutoFit
    
    If .AutoFilterMode = False Then .Range("B1:D" & .Cells(Rows.Count, 1).End(xlUp).row).AutoFilter
    .Range("A1").CurrentRegion.EntireColumn.AutoFit

End With

freezeTopPane activeWindow
'ActiveWorkbook.Save

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
Set dict = Nothing

End Sub
Private Sub createTier()
Dim ws As Worksheet
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim data
Dim arr
Dim coll As New Collection
Dim printArr
Dim strArr() As String
Dim i As Long
Dim j As Long

Set ws = Worksheets(3)
With ws
    .Activate
    .Name = "Tier"
End With

createHeaders ws.Name
ws.Range("A2:W" & ws.Cells(Rows.Count, 1).End(xlDown).row).ClearContents

getConnection

SQL = "SELECT * " & _
      "FROM [FanDuel$]" & _
      "WHERE [Tier] = 1 " & _
      "ORDER BY [Salary] DESC"
          
rs.Open SQL, conn

If Not rs.EOF Then
    
    data = rs.GetRows
    ReDim arr(1 To UBound(data, 2) + 1, 1 To UBound(data, 1) + 1)
    
    For i = LBound(data, 2) To UBound(data, 2)
        For j = LBound(data, 1) To UBound(data, 1)
            If Not IsNull(data(j, 1)) Then
                arr(i + 1, j + 1) = data(j, i)
            End If
        Next j
    Next i
End If

'Process Data
Set coll = getCollection(arr, 6)

ReDim printArr(1 To coll.Count, 22)

For i = 1 To coll.Count
    strArr = Split(coll(i), "_")
    
    For j = 0 To UBound(strArr)
        printArr(i, j) = strArr(j)
    Next j
Next i

'Process Array
With ws
    Range("A1").Activate
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    
    If Application.CountA(.Range("W:W")) > 1 Then
        .Range("A2:W" & Application.CountA(.Range("W:W"))).ClearContents
    End If
    
    'createHeaders wb, .Name
    .Range("A2:W" & UBound(printArr) + 1).Value = printArr
    .Range("A1").CurrentRegion.EntireColumn.AutoFit
    freezeTopPane activeWindow
    
    If .AutoFilterMode = False Then .Range("A1").AutoFilter
    .Range("A1").CurrentRegion.EntireColumn.AutoFit

    With .Sort
        .SortFields.Clear
        .SortFields.Add key:=Range("P1"), Order:=xlDescending
        .SetRange Range("A2:W" & UBound(printArr) + 1)
        .Apply
    End With
End With

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub
Private Sub createSearch()
Dim ws As Worksheet
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim arr
Dim i As Long
Dim j As Integer
Dim btn As Button

Set ws = Worksheets(1)
With ws
    .Activate
    .Name = "Search"
End With

createHeaders ws.Name
ws.Range("A2:AA" & ws.Cells(Rows.Count, 1).End(xlDown).row).ClearContents

'getPlayerArray
SQL = "SELECT * " & _
      "FROM [FanDuel$]" & _
      "WHERE [Tier] = 1 " & _
      "ORDER BY [Salary] DESC"

getConnection
rs.Open SQL, conn

If Not rs.EOF Then
    arr = rs.GetRows
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData
    For i = 0 To UBound(arr, 2)
        Range("A" & i + 2).Value = arr(16, i) 'Projected Points
        Range("A" & i + 2).NumberFormat = "##.0"
        Range("B" & i + 2).Value = arr(1, i) & ":" & arr(10, i) & " " & arr(12, i) 'Position:Team inj ind
        
        For j = 3 To 5
            With Cells(i + 2, j).Validation
                .Delete
                .Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:="=$B$" & i + 2
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        Next j
    Next i
    
    If ws.Buttons.Count > 0 Then ws.Buttons.Delete
        
    Set btn = ActiveSheet.Buttons.Add(0.5 * Range("A1").Width, Range("A" & UBound(arr, 2) + 4).Top, Range("A1").Width, Range("A18").Height * 1.5)
    With btn
        .Caption = "Search"
        .OnAction = "Tier.xlsm!search"
        .Placement = xlFreeFloating
    End With
    
    If ActiveSheet.AutoFilterMode = False Then Range("A1:AB1").AutoFilter
    Range("A1:AB1").CurrentRegion.EntireColumn.AutoFit
        
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub
Sub createRandomLineup()
Dim ws As Worksheet
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim arr
Dim i As Long
Dim j As Integer
Dim btn As Button

Set ws = Worksheets(6)
With ws
    .Activate
    .Name = "Random Lineup"
End With

createHeaders ws.Name

'getPlayerArray
SQL = "SELECT [Nickname], [Position] & ':' & [Team] &' '&[Injury Indicator]" & _
      "FROM [FanDuel$]" & _
      "WHERE [Tier] = 1 " & _
      "ORDER BY [Salary] DESC"
          
getConnection
rs.Open SQL, conn

If Not rs.EOF Then
    arr = rs.GetRows
    For i = 0 To UBound(arr, 2)
        Range("A" & i + 2).Value = arr(0, i)
        Range("B" & i + 2).Value = arr(1, i)
        
        For j = 3 To 5
            With Cells(i + 2, j).Validation
                .Delete
                .Add Type:=xlValidateList, _
                           AlertStyle:=xlValidAlertStop, _
                           Operator:=xlBetween, _
                           Formula1:="=$B$" & i + 2
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = True
            End With
        Next j
    Next i

    Range("A1").CurrentRegion.EntireColumn.AutoFit

    If ws.Buttons.Count > 0 Then ws.Buttons.Delete
        
    Set btn = ws.Buttons.Add(0.5 * Range("A1").Width, Range("A" & UBound(arr, 2) + 4).Top, 0.75 * Range("A1").Width, Range("A18").Height * 1.5)
    With btn
        .Caption = "Create"
        .OnAction = "Tier.xlsm!getRandomLineup"
        .Placement = xlFreeFloating
    End With
    
    Set btn = ws.Buttons.Add(0.5 * Range("A1").Width, Range("A" & UBound(arr, 2) + 6).Top, 0.75 * Range("A1").Width, Range("A18").Height * 1.5)
    With btn
        .Caption = "Clear"
        .OnAction = "Tier.xlsm!removeRandomLineup"
        .Placement = xlFreeFloating
    End With
End If

Range("C:K").ColumnWidth = Columns("B").ColumnWidth
Range("L:Q").ColumnWidth = Columns("A").ColumnWidth

Range("C2").Activate

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub
Private Sub createFanduel()
Dim ws As Worksheet
Dim path As String
Dim fileName As String
Dim conn As New ADODB.Connection
Dim rsSchema As New ADODB.Recordset
Dim SheetName As String
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim i As Long
Dim data As Variant
Dim j As Long
Dim arr

Set ws = Worksheets(5)
With ws
    .Activate
    .Name = "FanDuel"
End With

path = ActiveWorkbook.path
fileName = Dir(ActiveWorkbook.path & "\FanDuel*.xlsx")

'Get Sheet Name
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & path & "\" & fileName & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"

Set rsSchema = conn.OpenSchema(adSchemaTables)

Do Until rsSchema.EOF
    If InStr(rsSchema.fields("TABLE_NAME"), "FanDuel") Then
        SheetName = rsSchema.fields("TABLE_NAME")
        Exit Do
    End If
    rsSchema.MoveNext
Loop

'Import Player Positions
SQL = "SELECT * " & _
      "FROM [" & SheetName & "]"
          
'Set rs = Conn.Execute(SQL)
rs.Open SQL, conn

'Populate Headers
For i = 0 To rs.fields.Count - 1
    ws.Cells(1, i + 1).Value = rs.fields(i).Name
Next

'Populate Data
With ws
    .Range("A2").CopyFromRecordset rs
    .Cells(1, 19).Value = "FPPG Rank"
    .Range("A1").CurrentRegion.Sort Key1:=.Cells(1, 15), _
                                    Order1:=xlAscending, _
                                    key2:=.Cells(1, 6), _
                                    Order2:=xlDescending, _
                                    header:=xlYes
    For i = 2 To .Cells(Rows.Count, 1).End(xlUp).row
        .Cells(i, 19).Value = i - 1
    Next i

    'Rank Salary
    .Cells(1, 18).Value = "Salary Rank"
    .Range("A1").CurrentRegion.Sort Key1:=.Cells(1, 15), _
                                    Order1:=xlDescending, _
                                    key2:=.Cells(1, 8), _
                                    Order2:=xlDescending, _
                                    header:=xlYes
    For i = 2 To .Cells(Rows.Count, 1).End(xlUp).row
        .Cells(i, 18).Value = i - 1
    Next i
        
    .Range("A1").CurrentRegion.EntireColumn.AutoFit
End With

freezeTopPane activeWindow
'ActiveWorkbook.Save

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub

Sub createCSV()
Dim header As String: header = "MVP - 1.5X Points,AnyFLEX,AnyFLEX,AnyFLEX,AnyFLEX"
Dim path As String: path = ActiveWorkbook.path
Dim outfile
Dim xl As New Excel.Application

outfile = FreeFile

If Dir(ActiveWorkbook.path & "\my_picks.csv") = "" Then
    Open ActiveWorkbook.path & "\my_picks.csv" For Output As #outfile
    Print #outfile, header
    Close outfile
End If

Workbooks.Open ActiveWorkbook.path & "\my_picks.csv", ReadOnly:=False
'xl.Visible = True

Set xl = Nothing

End Sub
Private Sub createHeaders(psheetName As String)
'Dim fileName As String: fileName = pfileName
Dim btn As Button
Dim cell As Range
Dim sheetNum As Integer: sheetNum = 3
Dim i As Integer
Dim str As String
Dim arr
Dim strArr() As String

Select Case psheetName 'If psheetName = "Tier" Or psheetName = "Retro" Then
    Case "Tier", "Retro"
        ReDim strArr(1 To 23)
            strArr(1) = ""
            strArr(2) = ""
            strArr(3) = "key"
            strArr(4) = "salary_rank"
            strArr(5) = "fppg_rank"
            strArr(6) = "MVP_pos"
            strArr(7) = "p2_pos"
            strArr(8) = "p3_pos"
            strArr(9) = "p4_pos"
            strArr(10) = "p5_pos"
            strArr(11) = "p6_pos"
            strArr(12) = "select"
            strArr(13) = "team_cnt"
            strArr(14) = "total_salary"
            strArr(15) = "total_fppg"
            strArr(16) = "total_ppts"
            strArr(17) = "total_pts"
            strArr(18) = "MVP_name"
            strArr(19) = "p2_name"
            strArr(20) = "p3_name"
            strArr(21) = "p4_name"
            strArr(22) = "p5_name"
            strArr(23) = "p6_name"
    Case "Search"
        ReDim strArr(1 To 28)
        strArr(1) = "PPTS"
        strArr(2) = "Position"
        strArr(3) = "MVP"
        strArr(4) = "Include"
        strArr(5) = "Exclude"
        strArr(6) = ""
        strArr(7) = ""
        strArr(8) = "key"
        strArr(9) = "salary_rank"
        strArr(10) = "fppg_rank"
        strArr(11) = "MVP_pos"
        strArr(12) = "p2_pos"
        strArr(13) = "p3_pos"
        strArr(14) = "p4_pos"
        strArr(15) = "p5_pos"
        strArr(16) = "p6_pos"
        strArr(17) = "select"
        strArr(18) = "team_cnt"
        strArr(19) = "total_salary"
        strArr(20) = "total_fppg"
        strArr(21) = "total_ppts"
        strArr(22) = "total_pts"
        strArr(23) = "MVP_name"
        strArr(24) = "p2_name"
        strArr(25) = "p3_name"
        strArr(26) = "p4_name"
        strArr(27) = "p5_name"
        strArr(28) = "p6_name"
    Case "Matrix"
        ReDim strArr(1 To 12)
        strArr(1) = "Nickname"
        strArr(2) = "Position"
        strArr(3) = "MVP"
        strArr(4) = "p2_pos"
        strArr(5) = "p3_pos"
        strArr(6) = "p4_pos"
        strArr(7) = "p5_pos"
        strArr(8) = "p6_pos"
        strArr(9) = "Total"
        strArr(10) = ""
        strArr(11) = "key"
        strArr(12) = "Total"
    Case "Random Lineup"
        ReDim strArr(1 To 19)
        strArr(1) = "Nickname"
        strArr(2) = "Position"
        strArr(3) = "MVP"
        strArr(4) = "Flex"
        strArr(5) = "Exclude"
        strArr(6) = ""
        strArr(7) = "MVP_pos"
        strArr(8) = "p2_pos"
        strArr(9) = "p3_pos"
        strArr(10) = "p4_pos"
        strArr(11) = "p5_pos"
        strArr(12) = "p6_pos"
        strArr(13) = "total_ppts"
        strArr(14) = "MVP_name"
        strArr(15) = "p2_name"
        strArr(16) = "p3_name"
        strArr(17) = "p4_name"
        strArr(18) = "p5_name"
        strArr(19) = "p6_name"
    End Select

Sheets(psheetName).Range(Cells(1, 1), Cells(1, UBound(strArr))).Value = strArr
   
End Sub
Sub createWSProjections(pfileName, psheetName, pPosition)
Dim fileName As String: fileName = pfileName
Dim btn As Button
Dim sheetNum As Integer: sheetNum = 1
Dim row As Double
Dim spArr
Dim currentPosition As String
Dim previousPosition As String
Dim i As Long
Dim dArr(1) As Integer
Dim kArr(1) As Integer
Dim fArr(1) As Integer

'spArr = getPlayerArray(fileName)
       
With Workbooks(pfileName)
    .Sheets(sheetNum).Name = psheetName
    For i = 2 To UBound(spArr)
        currentPosition = spArr(i, 2)
        If currentPosition = "D" Then
            If currentPosition <> previousPosition Then
                row = 1
                With .Sheets(psheetName).Cells(row, 1)
                    .Value = "D"
                    .Offset(, 1).Value = "DE/SF"
                    .Offset(, 2).Value = "DE/RTD"
                    .Offset(, 3).Value = "DE/FRTD"
                    .Offset(, 4).Value = "DE/BRTD"
                    .Offset(, 5).Value = "DE/XPR"
                    .Offset(, 6).Value = "DE/PA0"
                    .Offset(, 7).Value = "DE/PA1-6"
                    .Offset(, 8).Value = "DE/PA7-13"
                    .Offset(, 9).Value = "DE/PA14-20"
                    .Offset(, 10).Value = "DE/PA28-34"
                    .Offset(, 11).Value = "DE/PA35+"
                    .Offset(, 12).Value = "FR"
                    .Offset(, 13).Value = "DE/I"
                    .Offset(, 14).Value = "S"
                    .Offset(, 15).Value = "DE/B"
                    .Offset(, 20).Value = "Game Points"
                    .Offset(, 21).Value = "Fantasy Points"
                End With
                Rows(row).Font.Bold = True
                dArr(0) = row + 1
            End If
            
            row = row + 1
            dArr(1) = row
            Cells(row, 1).Value = spArr(i, 2) & ":" & spArr(i, 11)
            previousPosition = currentPosition
        ElseIf currentPosition = "K" Then
            If currentPosition <> previousPosition Then
                row = row + 2
                With .Sheets(psheetName).Cells(row, 1)
                    .Value = "K"
                    .Offset(, 1).Value = "FGu20"
                    .Offset(, 2).Value = "FGu30"
                    .Offset(, 3).Value = "FGu40"
                    .Offset(, 4).Value = "FGu50"
                    .Offset(, 5).Value = "FGo50"
                    .Offset(, 6).Value = "XP"
                    .Offset(, 7).Value = "2PC/S"
                    .Offset(, 8).Value = "2PC/P"
                    .Offset(, 9).Value = "FU/L"
                    .Offset(, 10).Value = "I"
                End With
                Rows(row).Font.Bold = True
                kArr(0) = row + 1
            End If
            
            row = row + 1
            kArr(1) = row
            Cells(row, 1).Value = spArr(i, 2) & ":" & spArr(i, 11)
            previousPosition = currentPosition
        Else
            If currentPosition = "QB" And currentPosition <> previousPosition Then
                row = row + 2
                With .Sheets(psheetName).Cells(row, 1)
                    .Value = "FLEX"
                    .Offset(, 1).Value = "RuTD"
                    .Offset(, 2).Value = "ReTD"
                    .Offset(, 3).Value = "FU/TD"
                    .Offset(, 4).Value = "KR/TD"
                    .Offset(, 5).Value = "PR/TD"
                    .Offset(, 6).Value = "2PC/S"
                    .Offset(, 7).Value = "PaTD"
                    .Offset(, 8).Value = "PaY"
                    .Offset(, 9).Value = "300+ PaY Gm"
                    .Offset(, 10).Value = "RuY"
                    .Offset(, 11).Value = "100+ RuY GM"
                    .Offset(, 12).Value = "Re"
                    .Offset(, 13).Value = "ReY"
                    .Offset(, 14).Value = "100+ ReY GM"
                    .Offset(, 15).Value = "2PC/P"
                    .Offset(, 16).Value = "I"
                    .Offset(, 17).Value = "FR"
                    .Offset(, 18).Value = "FU/L"
                End With
                Rows(row).Font.Bold = True
                fArr(0) = row + 1
            End If
            
            row = row + 1
            fArr(1) = row
            Cells(row, 1).Value = spArr(i, 2) & ":" & spArr(i, 11)
            previousPosition = currentPosition
        End If
    Next
    
    Cells.ColumnWidth = 7.5
    
    .Sheets(psheetName).Activate
    .Save
End With

End Sub
Private Sub createEventProcedure()
Dim ws As Worksheet
Dim VBProj As VBIDE.VBProject
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule
Dim LineNum As Long
Dim LineCnt As Long

Set VBProj = ActiveWorkbook.VBProject
Set VBComp = VBProj.VBComponents(Worksheets("Search").codename)
Set CodeMod = VBComp.CodeModule

If CodeMod.CountOfLines = 0 Then
    With CodeMod
        LineNum = .CreateEventProc("Change", "Worksheet")
        LineNum = LineNum + 1
        .InsertLines LineNum, "If Target.Column = 17 Then"
        .InsertLines LineNum + 1, "For Each cell In Target"
        .InsertLines LineNum + 2, "If cell.Column = 17 Then"
        .InsertLines LineNum + 3, "Worksheets(3).Cells(Worksheets(3).Range(""$A:$A"").Find(What:=cell.Offset(, -11), LookAt:=xlWhole).row, 12) = Target.Value"
        .InsertLines LineNum + 4, "End If"
        .InsertLines LineNum + 5, "Next"
        .InsertLines LineNum + 6, "End If"
    End With
End If
'ActiveWorkbook.VBProject.VBE.MainWindow.Visible = False

End Sub
Private Function getCollection(ByRef arr, pnum As Integer) As Collection
Dim i As Long
Dim key(5) As Integer: key(0) = 0
Dim team As String
Dim teamCnt(5) As Integer
Dim id(5) As String
Dim salary(5) As Double
Dim mvpSalary(5) As Double
Dim fppg(5) As Double
Dim points(5) As Double
Dim ppts(5) As Double
Dim pos(5) As String
Dim sKey(5) As Integer
Dim key2 As Integer
Dim fKey(5) As Integer
Dim j As Integer
Dim k As Integer
Dim l As Integer
Dim m As Integer
Dim n As Integer
Dim o As Integer
Dim nArr
Dim row  As Double: row = 0
Dim groupRow  As Double: groupRow = 0
Dim str As String
Dim coll As New Collection
Dim maxRows As Long: maxRows = 100000

For i = 1 To 5
    key(0) = arr(i, 18)
    'key(1) = key(0)
    team = arr(i, 11)
    teamCnt(0) = 0
    id(0) = arr(i, 1) & ":" & arr(i, 4)
    salary(0) = arr(i, 8)
    mvpSalary(0) = arr(i, 9)
    fppg(0) = arr(i, 6)
    points(0) = arr(i, 16)
    ppts(0) = arr(i, 17)
    pos(0) = arr(i, 2) & ":" & arr(i, 11) & " " & arr(i, 13)
    sKey(0) = arr(i, 18)
    fKey(0) = arr(i, 19)
    If arr(i, 11) = team Then teamCnt(0) = teamCnt(0) + 1
                
    For j = i + 1 To UBound(arr)
        key(1) = arr(j, 18)
        teamCnt(1) = 0
        id(1) = arr(j, 1) & ":" & arr(j, 4)
        salary(1) = arr(j, 8)
        mvpSalary(1) = arr(j, 9)
        fppg(1) = arr(j, 6)
        points(1) = arr(j, 16)
        ppts(1) = arr(j, 17)
        pos(1) = arr(j, 2) & ":" & arr(j, 11) & " " & arr(j, 13)
        sKey(1) = arr(i, 18)
        fKey(1) = arr(j, 19)
        If arr(j, 11) = team Then teamCnt(1) = teamCnt(1) + 1

        For k = j + 1 To UBound(arr)
            teamCnt(2) = 0
            id(2) = arr(k, 1) & ":" & arr(k, 4)
            salary(2) = arr(k, 8)
            mvpSalary(2) = arr(k, 9)
            fppg(2) = arr(k, 6)
            points(2) = arr(k, 16)
            ppts(2) = arr(k, 17)
            pos(2) = arr(k, 2) & ":" & arr(k, 11) & " " & arr(k, 13)
            sKey(2) = arr(k, 18)
            fKey(2) = arr(k, 19)
            If arr(k, 11) = team Then teamCnt(2) = teamCnt(2) + 1

            For l = k + 1 To UBound(arr)
                teamCnt(3) = 0
                id(3) = arr(l, 1) & ":" & arr(l, 4)
                salary(3) = arr(l, 8)
                mvpSalary(3) = arr(l, 9)
                fppg(3) = arr(l, 6)
                points(3) = arr(l, 16)
                ppts(3) = arr(l, 17)
                pos(3) = arr(l, 2) & ":" & arr(l, 11) & " " & arr(l, 13)
                sKey(3) = arr(l, 18)
                fKey(3) = arr(l, 19)
                If arr(l, 11) = team Then teamCnt(3) = teamCnt(3) + 1

                For m = l + 1 To UBound(arr)
                    teamCnt(4) = 0
                    id(4) = arr(m, 1) & ":" & arr(m, 4)
                    salary(4) = arr(m, 8)
                    mvpSalary(4) = arr(m, 9)
                    fppg(4) = arr(m, 6)
                    points(4) = arr(m, 16)
                    ppts(4) = arr(m, 17)
                    pos(4) = arr(m, 2) & ":" & arr(m, 11) & " " & arr(m, 13)
                    sKey(4) = arr(m, 18)
                    fKey(4) = arr(m, 19)
                    If arr(m, 11) = team Then teamCnt(4) = teamCnt(4) + 1

                    For o = m + 1 To UBound(arr)
                        teamCnt(5) = 0
                        id(5) = arr(o, 1) & ":" & arr(o, 4)
                        salary(5) = arr(o, 8)
                        mvpSalary(5) = arr(o, 9)
                        fppg(5) = arr(o, 6)
                        points(5) = arr(o, 16)
                        ppts(5) = arr(o, 17)
                        pos(5) = arr(o, 2) & ":" & arr(o, 11) & " " & arr(o, 13)
                        sKey(5) = arr(o, 18)
                        fKey(5) = arr(o, 19)
                        If arr(o, 11) = team Then teamCnt(5) = teamCnt(5) + 1
                    
                    If salary(0) + salary(1) + salary(2) + salary(3) + salary(4) + mvpSalary(5) <= 60000 And _
                       WorksheetFunction.Sum(teamCnt) > 0 And _
                       WorksheetFunction.Sum(teamCnt) < 6 Then
                           For n = 1 To pnum
                                Select Case n
                                    Case 1
                                        nArr = Array(0, 1, 2, 3, 4, 5)
                                    Case 2
                                        nArr = Array(1, 0, 2, 3, 4, 5)
                                    Case 3
                                        nArr = Array(2, 1, 0, 3, 4, 5)
                                    Case 4
                                        nArr = Array(3, 1, 2, 0, 4, 5)
                                    Case 5
                                        nArr = Array(4, 1, 2, 3, 0, 5)
                                    Case 6
                                        nArr = Array(5, 4, 1, 2, 3, 0)
                                End Select
                                
                                If mvpSalary(nArr(0)) + salary(nArr(1)) + salary(nArr(2)) + salary(nArr(3)) + salary(nArr(4)) + salary(nArr(5)) <= 60000 Then
                                str = CStr(row) & "_" & _
                                      CStr(groupRow) & "_" & _
                                      key(0) & key(1) & "_" & _
                                      sKey(nArr(0)) & "_" & _
                                      fKey(nArr(0)) & "_" & _
                                      pos(nArr(0)) & "_" & _
                                      pos(nArr(1)) & "_" & _
                                      pos(nArr(2)) & "_" & _
                                      pos(nArr(3)) & "_" & _
                                      pos(nArr(4)) & "_" & _
                                      pos(nArr(5)) & "_" & _
                                      "" & "_" & _
                                      CStr(teamCnt(nArr(0)) + teamCnt(nArr(1)) + teamCnt(nArr(2)) + teamCnt(nArr(3)) + teamCnt(nArr(4)) + teamCnt(nArr(5))) & "_" & _
                                      CStr(mvpSalary(nArr(0)) + salary(nArr(1)) + salary(nArr(2)) + salary(nArr(3)) + salary(nArr(4)) + salary(nArr(5))) & "_" & _
                                      CStr(Round(fppg(nArr(0)) + fppg(nArr(1)) + fppg(nArr(2)) + fppg(nArr(3)) + fppg(nArr(4)) + fppg(nArr(5)), 2)) & "_" & _
                                      CStr(Round(1.5 * ppts(nArr(0)) + ppts(nArr(1)) + ppts(nArr(2)) + ppts(nArr(3)) + ppts(nArr(4)) + ppts(nArr(5)), 2)) & "_" & _
                                      CStr(Round(1.5 * points(nArr(0)) + points(nArr(1)) + points(nArr(2)) + points(nArr(3)) + points(nArr(4)) + points(nArr(5)), 2)) & "_" & _
                                      id(nArr(0)) & "_" & _
                                      id(nArr(1)) & "_" & _
                                      id(nArr(2)) & "_" & _
                                      id(nArr(3)) & "_" & _
                                      id(nArr(4)) & "_" & _
                                      id(nArr(5))

                                coll.Add Item:=str
                                row = row + 1
                                If row = maxRows Then Exit For
                                End If
                           Next n
                        groupRow = groupRow + 1
                    End If
                    If row = maxRows Then Exit For
                    Next o
                    If row = maxRows Then Exit For
                Next m
                If row = maxRows Then Exit For
            Next l
            If row = maxRows Then Exit For
        Next k
        If row = maxRows Then Exit For
    Next j
    If row = maxRows Then Exit For
Next i

Set getCollection = coll

End Function
Sub InsertDynamicModuleAndCode()
        Dim VBProj As Object ' VBIDE.VBProject
        Dim VBComp As Object ' VBIDE.VBComponent
        Dim CodeMod As Object ' VBIDE.CodeModule
        Dim ModuleName As String
        Dim CodeToAdd As String

        ' Set the project object
        Set VBProj = ActiveWorkbook.VBProject

        ' Define the new module name
        ModuleName = "NewDynamicModule"

        ' Add a new standard module
        Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule)
        VBComp.Name = ModuleName

        ' Get the CodeModule object of the new module
        Set CodeMod = VBComp.CodeModule

        ' Define the code to be added
        CodeToAdd = "Sub MyDynamicMacro()" & vbCrLf & _
                    "    MsgBox ""This is a dynamically created macro!""" & vbCrLf & _
                    "End Sub"

        ' Add the code to the module
        CodeMod.AddFromString CodeToAdd

        ' Clean up objects
        Set CodeMod = Nothing
        Set VBComp = Nothing
        Set VBProj = Nothing

        MsgBox "Module '" & ModuleName & "' created and code inserted successfully."
    End Sub
'Explanation of the Code:
'Set VBProj = ActiveWorkbook.VBProject: This line gets a reference to the VBA project of the active workbook.
'Set VBComp = VBProj.VBComponents.Add(vbext_ct_StdModule): This adds a new standard module to the VBA project. vbext_ct_StdModule specifies the type of component to add (a standard module).
'VBComp.Name = ModuleName: This assigns a name to the newly created module.
'Set CodeMod = VBComp.CodeModule: This obtains the CodeModule object associated with the new module, allowing manipulation of its code.
'CodeMod.AddFromString CodeToAdd: This method inserts the specified string of VBA code into the module. vbCrLf is used for line breaks within the code string.
'

Sub AddCodeToThisWorkbook()

        Dim wb As Workbook
        Dim cm As CodeModule
        Dim sCode As String

        Set wb = ThisWorkbook ' Refers to the workbook containing this code
        Set cm = wb.VBProject.VBComponents("ThisWorkbook").CodeModule

        ' Define the code to be inserted
        sCode = "Private Sub Workbook_Open()" & vbCrLf & _
                "    MsgBox ""Welcome to this workbook!""" & vbCrLf & _
                "End Sub"

        ' Insert the code at the end of the module
        cm.InsertLines cm.CountOfLines + 1, sCode

        MsgBox "Code added to ThisWorkbook module successfully!", vbInformation

    End Sub
Sub attachModule()
Dim ModulePath As String
Dim VBProj As Object
Dim VBComp As Object

Set VBProj = ActiveWorkbook.VBProject

For Each VBComp In ActiveWorkbook.VBProject.VBComponents
    If VBComp.Name = "NFLButtons" Then Exit Sub
Next VBComp

Set VBProj = ActiveWorkbook.VBProject
ModulePath = ActiveWorkbook.path & "\VBA\NFLButtons.bas"
VBProj.VBComponents.Import ModulePath

End Sub
