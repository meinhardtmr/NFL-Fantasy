Attribute VB_Name = "NFLButtons"
Private conn As New Connection
Sub getConnection()
Dim wb As Workbook

Set wb = Workbooks("Tier.xlsm")

conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & wb.FullName & ";" & _
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
Sub search()
Dim ws As Worksheet
Dim i As Long, _
    l As Long, _
    j As Long: j = 1
Dim p As Integer
Dim cnt As Integer: cnt = 0
Dim wb As Workbook
Dim arrSearch, _
    arrSource, _
    arrPrint, _
    strArr, _
    pptsArr
Dim cell As Range
Dim collPrint As New Collection
Dim strInclude As String, _
    strExclude As String, _
    strPrint As String
Dim matchCnt As Integer
Dim r
Dim d As New Scripting.Dictionary
Dim rs As New Recordset
Dim cmd As New Command
Dim param As parameter
Dim MVP As String
Dim include As String
Dim exclude As String
Dim includeNum As Long
Dim excludeNum As Long
Dim recCount As Long
Dim arr

Set wb = ActiveWorkbook
Set ws = Worksheets("Search")
ws.Range("F2").Activate

Call getConnection

SQL = "SELECT [Position], [MVP], [Include], [Exclude] " & _
      "FROM [Search$]"
      
rs.CursorLocation = adUseClient
rs.Open SQL, conn

Do Until rs.EOF
    If Len(rs.fields("MVP").Value) > 0 And MVP = "" Then MVP = rs.fields("mvp").Value
    
    If Len(rs.fields("Include").Value) > 0 And include = "" Then
        include = "'" & rs.fields("Include").Value & "'"
        includeNum = includeNum + 1
    ElseIf Len(rs.fields("Include").Value) > 0 Then
        include = Left(include, Len(include) - 1) & " " & rs.fields("Include") & "'"
        includeNum = includeNum + 1
    End If
    
    If Len(rs.fields("Exclude").Value) > 0 And exclude = "" Then
        exclude = "'" & rs.fields("Exclude").Value & "'"
    ElseIf Len(rs.fields("Exclude").Value) > 0 Then
        exclude = Left(exclude, Len(exclude) - 1) & " " & rs.fields("Exclude").Value & "'"
    End If
    
    rs.MoveNext
Loop

rs.Close

'Query Tier Database
SQL = "SELECT F1, F2 " & _
             ",[key]" & _
             ",[salary_rank]" & _
             ",[fppg_rank]" & _
             ",[MVP_pos]" & _
             ",[p2_pos]" & _
             ",[p3_pos]" & _
             ",[p4_pos]" & _
             ",[p5_pos]" & _
             ",[p6_pos]" & _
             ",[select]" & _
             ",[team_cnt]" & _
             ",[total_salary]" & _
             ",[total_fppg]" & _
             ",[total_ppts]" & _
             ",[total_pts]" & _
             ",[mvp_name]" & _
             ",[p2_name]" & _
             ",[p3_name]" & _
             ",[p4_name]" & _
             ",[p5_name]" & _
             ",[p6_name] " & _
      "FROM [Tier$] " & _
      "WHERE mvp_pos = "

If Len(MVP) > 0 Then
    SQL = SQL & "?"
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, MVP)
    cmd.Parameters.Append param
Else
    SQL = SQL & " mvp_pos"
End If

If Len(include) > 0 Then
    SQL = SQL & " AND iif(instr(?,p2_pos)>0,1,0) + " & _
                     "iif(instr(?,p3_pos)>0,1,0) + " & _
                     "iif(instr(?,p4_pos)>0,1,0) + " & _
                     "iif(instr(?,p5_pos)>0,1,0) + " & _
                     "iif(instr(?,p6_pos)>0,1,0) = " & flexNum
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, include)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, include)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, include)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, include)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, include)
    cmd.Parameters.Append param
End If

If Len(exclude) > 0 Then
    SQL = SQL & " AND iif(instr(?,mvp_pos)>0,1,0) + " & _
                     "iif(instr(?,p2_pos)>0,1,0) + " & _
                     "iif(instr(?,p3_pos)>0,1,0) + " & _
                     "iif(instr(?,p4_pos)>0,1,0) + " & _
                     "iif(instr(?,p5_pos)>0,1,0) + " & _
                     "iif(instr(?,p6_pos)>0,1,0) = 0"
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
End If

Set cmd.ActiveConnection = conn
cmd.CommandText = SQL

rs.CursorLocation = adUseClient

rs.Open cmd
recCount = rs.RecordCount
arr = rs.GetRows
ReDim arrPrint(recCount + 1, rs.fields.Count - 1)

For i = 0 To recCount
    For j = 0 To UBound(arrPrint, 2)
        arrPrint(i, j) = arr(j, i)
    Next
Next

'If collPrint.Count = 0 Then
'    Sheets("Search").Range("E2:AA" & Sheets("Search").Cells(Rows.Count, 1).End(xlDown).row).Clear
'Else
'    ReDim arrPrint(1 To collPrint.Count, 22)
'    For i = 1 To collPrint.Count
'        strArr = Split(collPrint(i), "_")
'
'        For j = 0 To UBound(strArr)
'            arrPrint(i, j) = strArr(j)
'        Next j
'
'    Next i
'
'    If Sheets("Search").FilterMode Then Sheets("Search").ShowAllData
'    Sheets("Search").Range("E2:AA" & Sheets("Search").Cells(Rows.Count, 1).End(xlDown).row).Clear
'    'Sheets("Search").Range("E2:AA" & UBound(arrPrint) + 1).Value = arrPrint
    Sheets("Search").Range("F2").Resize(UBound(arrPrint), UBound(arrPrint, 2) + 1).Value = arrPrint
'
'    Sheets("Search").Range("A1").CurrentRegion.EntireColumn.AutoFit
'    freezeTopPane activeWindow
'    If Worksheets("Search").AutoFilterMode = False Then Sheets("Search").Range("C1").AutoFilter
'
    'Sort Worksheet
'    With Worksheets("Search")
'        .Sort.SortFields.Clear
'        .Range("E2:AA" & UBound(arrPrint)).Sort Key1:=.Cells(1, 20), _
'                                                Order1:=xlDescending, _'
'                                                header:=xlNo'
'    End With
'End If

'ActiveWorkbook.Save
conn.Close
Set conn = Nothing

End Sub
Sub getRandomLineup()
Dim wb As Workbook
Dim ws As Worksheet
Dim conn As New ADODB.Connection
Dim SQL As String
Dim rs As New ADODB.Recordset
Dim arr(0, 1 To 14)
Dim data
Dim i As Long
Dim MVP As String
Dim flex As String
Dim exclude As String
Dim cmd As New ADODB.Command
Dim param As ADODB.parameter
Dim random As Double
Dim flexNum As Integer: flexNum = 0
Dim dict As New Scripting.Dictionary
Dim recCount As Double
Dim strWhat As String

Set wb = ActiveWorkbook
Set ws = ActiveSheet

'getPlayerArray
SQL = "SELECT [Player], [MVP], [Flex], [Exclude] " & _
      "FROM [Random Lineup$]"

conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & wb.FullName & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=YES;"";"
          
rs.Open SQL, conn

Do Until rs.EOF
    If Len(rs.fields("MVP").Value) > 0 And MVP = "" Then MVP = rs.fields("mvp").Value
    
    If Len(rs.fields("Flex").Value) > 0 And flex = "" Then
        flex = "'" & rs.fields("Flex").Value & "'"
        flexNum = flexNum + 1
    ElseIf Len(rs.fields("Flex").Value) > 0 Then
        flex = Left(flex, Len(flex) - 1) & " " & rs.fields("Flex") & "'"
        flexNum = flexNum + 1
    End If
    
    If Len(rs.fields("Exclude").Value) > 0 And exclude = "" Then
        exclude = "'" & rs.fields("Exclude").Value & "'"
    ElseIf Len(rs.fields("Exclude").Value) > 0 Then
        exclude = Left(exclude, Len(exclude) - 1) & " " & rs.fields("Exclude").Value & "'"
    End If
    
    rs.MoveNext
Loop

rs.Close

'Seed the Randomizer
Randomize
 
'Query Tier Database
SQL = "SELECT F1" & _
             ",mvp_pos" & _
             ",p2_pos" & _
             ",p3_pos" & _
             ",p4_pos" & _
             ",p5_pos" & _
             ",p6_pos" & _
             ",total_ppts" & _
             ",mvp_name" & _
             ",p2_name" & _
             ",p3_name" & _
             ",p4_name" & _
             ",p5_name" & _
             ",p6_name " & _
      "FROM [Tier$] " & _
      "WHERE ([select] IS NULL Or [select] <> 0) " & _
      " AND mvp_pos = "

If Len(MVP) > 0 Then
    SQL = SQL & "?"
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, MVP)
    cmd.Parameters.Append param
Else
    SQL = SQL & " mvp_pos"
End If

If Len(flex) > 0 Then
    SQL = SQL & " AND iif(instr(?,p2_pos)>0,1,0) + " & _
                     "iif(instr(?,p3_pos)>0,1,0) + " & _
                     "iif(instr(?,p4_pos)>0,1,0) + " & _
                     "iif(instr(?,p5_pos)>0,1,0) + " & _
                     "iif(instr(?,p6_pos)>0,1,0) = " & flexNum
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 50, flex)
    cmd.Parameters.Append param
End If

If Len(exclude) > 0 Then
    SQL = SQL & " AND iif(instr(?,mvp_pos)>0,1,0) + " & _
                     "iif(instr(?,p2_pos)>0,1,0) + " & _
                     "iif(instr(?,p3_pos)>0,1,0) + " & _
                     "iif(instr(?,p4_pos)>0,1,0) + " & _
                     "iif(instr(?,p5_pos)>0,1,0) + " & _
                     "iif(instr(?,p6_pos)>0,1,0) = 0"
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarChar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
End If

Set cmd.ActiveConnection = conn
cmd.CommandText = SQL

rs.CursorLocation = adUseClient

rs.Open cmd
recCount = rs.RecordCount

If Not rs.EOF Then
    data = Application.WorksheetFunction.Transpose(rs.GetRows)
    If recCount > 1 Then
        random = WorksheetFunction.RandBetween(1, (UBound(data)))
        For i = 1 To UBound(arr, 2)
            arr(0, i) = data(random, i)
        Next
        strWhat = data(random, 1)
    Else
        For i = 1 To UBound(arr, 2)
            arr(0, i) = data(i)
        Next
        strWhat = data(1)
    End If
    
    With ws
        .Range("F" & .Cells(Rows.Count, 6).End(xlUp).row + 1 & ":S" & .Cells(Rows.Count, 6).End(xlUp).row + 1).Value = arr
        'mark select on Search worksheet
        'Worksheets(1).Cells(Worksheets(1).Range("$E:$E").Find(What:=strWhat, LookAt:=xlWhole).row, 16) = 1
        Worksheets(3).Cells(Worksheets(3).Range("$A:$A").Find(What:=strWhat, LookAt:=xlWhole).row, 12) = 0
        'Worksheets(3).Cells(Worksheets(3).Range(""$A:$A"").Find(What:=cell.Offset(, -11), LookAt:=xlWhole).row, 12) = Target.Value
    End With
Else
    MsgBox "No Lineups Found"
End If

ws.Range("G" & ws.Cells(Rows.Count, 6).End(xlUp).row).Activate

rs.Close
Set rs = Nothing
Set conn = Nothing

End Sub
Sub removeRandomLineup()
Dim wb As Workbook
Dim ws As Worksheet
Dim conn As New Connection
Dim cmd As New Command
Dim base_SQL As String
Dim SQL As String
Dim rs As New Recordset
Dim param As ADODB.parameter
Dim inClause As String
Dim recCount As Integer
Set wb = ActiveWorkbook
Set ws = ActiveSheet

'Update Tier to set select NULL
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" & _
          "Data Source=" & wb.FullName & ";" & _
          "Extended Properties=""Excel 12.0 Xml;HDR=YES;ReadOnly=False"";"

cmd.ActiveConnection = conn

SQL = "SELECT [F6] FROM [Random Lineup$] WHERE [F6] IS NOT NULL"
cmd.CommandText = SQL

rs.CursorLocation = adUseClient
rs.Open cmd
recCount = rs.RecordCount

base_SQL = "UPDATE[Tier$] " & _
      "SET [select] = NULL " & _
      "WHERE [F1] IN "
'cmd.CommandText = SQL

'Set param = cmd.CreateParameter("", adInteger, adParamInput, 50)
'cmd.Parameters.Append param

Do Until rs.EOF
    If Len(inClause) = 0 Then
        inClause = "(" & rs.fields("F6").Value & ")"
    Else
        inClause = Left(inClause, Len(inClause) - 1) & "," & rs.fields("F6").Value & ")"
    End If
    
    If rs.AbsolutePosition Mod 100 = 0 Or rs.AbsolutePosition = recCount Then
        SQL = base_SQL & inClause
        cmd.CommandText = SQL
        cmd.Execute
        SQL = base_SQL
        inClause = ""
    End If
    
    rs.MoveNext
Loop

ws.Range("F2:S" & ws.Cells(Rows.Count, 6).End(xlDown).row).ClearContents
ws.Range("F2").Activate

End Sub
Sub saveLineup()
Dim wb As Workbook
Dim ws As Worksheet
Dim strMVP As String
Dim strPositions As String
Dim arr
Dim i As Long
Dim foundCell As Range

Set wb = ActiveWorkbook
Set ws = ActiveSheet

With ws
    strMVP = .Range("C2")
    strPositions = .Range("C3") & " " & .Range("C4") & " " & .Range("C5") & " " & .Range("C6") & " " & .Range("C7")
End With

'Worksheets(3).Cells(Worksheets(3).Range("$L:$L").Find(What:=ws.Range("B2"), LookAt:=xlWhole).row, 12) = ""
    
arr = Sheets("Tier").Range("F2:K" & Sheets("Tier").Cells(Rows.Count, 6).End(xlUp).row).Value

For i = 1 To UBound(arr)
    If InStr(strMVP, arr(i, 1)) * _
       InStr(strPositions, arr(i, 2)) * _
       InStr(strPositions, arr(i, 3)) * _
       InStr(strPositions, arr(i, 4)) * _
       InStr(strPositions, arr(i, 5)) * _
       InStr(strPositions, arr(i, 6)) Then
        
        Set foundCell = Worksheets("Tier").Range("$L:$L").Find(What:=ws.Range("B2"), LookAt:=xlWhole)
        If Not Worksheets("Tier").Range("$L:$L").Find(What:=ws.Range("B2"), LookAt:=xlWhole) Is Nothing Then
            foundCell = ""
        End If
        With ws
            Sheets("Tier").Cells(i + 1, 12) = .Range("B2")
        End With
        Exit For
    End If
Next

End Sub


