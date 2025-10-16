Attribute VB_Name = "NFLButtons"
Private conn As Object
Const adParamInput = 1
Const adVarchar = 200
Const adInteger = 3
Const adUseClient = 3

Sub getConnection()
Set conn = CreateObject("ADODB.Connection")
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
Sub search()
Dim ws As Worksheet
Dim d As Object
Dim rs As Object
Dim MVP As String
Dim include As String
Dim exclude As String
Dim includeNum As Long
Dim arr
Dim arrPrint
Dim strP1 As String


Set ws = Worksheets("Search")
Range("F2").Activate

Set d = CreateObject("Scripting.Dictionary")
Set rs = CreateObject("ADODB.Recordset")

getConnection

'Create Dictionary
SQL = "SELECT [PPTS], [Position] " & _
      "FROM [Search$] " & _
      "WHERE [PPTS] IS NOT NULL"

rs.Open SQL, conn

Do While Not rs.EOF
    d(rs.fields("Position").Value) = Round(rs.fields("PPTS").Value, 1)
    rs.MoveNext
Loop

rs.Close

'Get SQL parameters
SQL = "SELECT [MVP], [Include], [Exclude] " & _
      "FROM [Search$] " & _
      "WHERE [MVP] IS NOT NULL OR [Include] IS NOT NULL OR [Exclude] IS NOT NULL"

'rs.CursorLocation = adUseClient
rs.Open SQL, conn

Do Until rs.EOF
    If Len(rs.fields("MVP").Value) > 0 And MVP = "" Then MVP = "'" & rs.fields("mvp").Value & "'"
    
    If Len(rs.fields("Include").Value) > 0 And include = "" Then
        include = "'" & rs.fields("Include").Value & "'"
        includeNum = 1
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
    SQL = SQL & MVP
Else
    SQL = SQL & "mvp_pos"
End If

If Len(include) > 0 Then
    If Len(MVP) = 0 Then
        strP1 = "iif(InStr(" & include & ",[mvp_pos])>0,1,0) + "
    End If
    SQL = SQL & " AND " & strP1 & "iif(instr(" & include & ",[p2_pos])>0,1,0) + " & _
                     "iif(instr(" & include & ",[p3_pos])>0,1,0) + " & _
                     "iif(instr(" & include & ",[p4_pos])>0,1,0) + " & _
                     "iif(instr(" & include & ",[p5_pos])>0,1,0) + " & _
                     "iif(instr(" & include & ",[p6_pos])>0,1,0) = " & includeNum
End If
Debug.Print SQL
If Len(exclude) > 0 Then
    SQL = SQL & " AND iif(instr(" & exclude & ",[mvp_pos])>0,1,0) + " & _
                     "iif(instr(" & exclude & ",[p2_pos])>0,1,0) + " & _
                     "iif(instr(" & exclude & ",[p3_pos])>0,1,0) + " & _
                     "iif(instr(" & exclude & ",[p4_pos])>0,1,0) + " & _
                     "iif(instr(" & exclude & ",[p5_pos])>0,1,0) + " & _
                     "iif(instr(" & exclude & ",[p6_pos])>0,1,0) = 0"
End If

'rs.CursorLocation = adUseClient
rs.Open SQL, conn

If Not rs.EOF Then
    arr = rs.GetRows
    ReDim arrPrint(UBound(arr, 2), UBound(arr))

    For i = 0 To UBound(arr, 2)
        For j = 0 To UBound(arr)
            arrPrint(i, j) = arr(j, i)
        Next
        arrPrint(i, 15) = d(arrPrint(i, 5)) * 1.5 + d(arrPrint(i, 6)) + d(arrPrint(i, 7)) + d(arrPrint(i, 8)) + d(arrPrint(i, 9)) + d(arrPrint(i, 10))
    Next

    If UBound(arrPrint) >= 0 Then
        With ws
            If .FilterMode Then Sheets("Search").ShowAllData
            .Range("F2:AB" & Sheets("Search").Cells(Rows.Count, 6).End(xlDown).row).Clear
            .Range("F2").Resize(UBound(arrPrint) + 1, UBound(arrPrint, 2) + 1).Value = arrPrint

            .Range("A1").CurrentRegion.EntireColumn.AutoFit
            freezeTopPane activeWindow
            If .AutoFilterMode = False Then Sheets("Search").Range("C1").AutoFilter
                
            'Sort Worksheet
            .Sort.SortFields.Clear
            .Range("F2:AB" & UBound(arrPrint) + 2).Sort Key1:=.Cells(1, 21), _
                                                    Order1:=xlDescending, _
                                                    header:=xlNo '
        End With
    End If
Else
    Sheets("Search").Range("F2:AB" & Sheets("Search").Cells(Rows.Count, 1).End(xlDown).row).Clear
    MsgBox "No Lineups Found"
End If

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
Set d = Nothing

End Sub
Sub getRandomLineup()
Dim ws As Worksheet
Dim SQL As String
Dim rs As Object
Dim arr(0, 1 To 14)
Dim data
Dim MVP As String
Dim flex As String
Dim exclude As String
Dim cmd As Object
Dim param As Object
Dim random As Double
Dim flexNum As Integer: flexNum = 0
Dim dict As Object
Dim strP1 As String

'Const adParamInput = 1
'Const adVarchar = 200
'Const adInteger = 3
'Const adUseClient = 3

Set ws = ActiveSheet

Set rs = CreateObject("ADODB.Recordset")
Set cmd = CreateObject("ADODB.Command")
Set dict = CreateObject("Scripting.Dictionary")

'getPlayerArray
SQL = "SELECT [MVP], [Flex], [Exclude] " & _
      "FROM [Random Lineup$] " & _
      "WHERE [MVP] IS NOT NULL OR [Flex] IS NOT NULL OR [Exclude] IS NOT NULL"

getConnection

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
SQL = "SELECT top 25 [F1]" & _
             ",[mvp_pos]" & _
             ",[p2_pos]" & _
             ",[p3_pos]" & _
             ",[p4_pos]" & _
             ",[p5_pos]" & _
             ",[p6_pos]" & _
             ",[total_ppts]" & _
             ",[mvp_name]" & _
             ",[p2_name]" & _
             ",[p3_name]" & _
             ",[p4_name]" & _
             ",[p5_name]" & _
             ",[p6_name] " & _
      "FROM [Tier$] " & _
      "WHERE [select] IS NULL AND [F1] NOT IN (SELECT [F6] FROM [Random Lineup$] WHERE [F6] IS NOT NULL) " & _
      "AND mvp_pos = "

If Len(MVP) > 0 Then
    SQL = SQL & "?"
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, MVP)
    cmd.Parameters.Append param
Else
    SQL = SQL & "mvp_pos"
End If

If Len(flex) > 0 Then
    If Len(MVP) = 0 Then
        strP1 = "iif(instr(?,[mvp_pos])>0,1,0) + "
        Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
        cmd.Parameters.Append param
    End If
    SQL = SQL & " AND " & strP1 & "iif(instr(?,[p2_pos])>0,1,0) + " & _
                     "iif(instr(?,[p3_pos])>0,1,0) + " & _
                     "iif(instr(?,[p4_pos])>0,1,0) + " & _
                     "iif(instr(?,[p5_pos])>0,1,0) + " & _
                     "iif(instr(?,[p6_pos])>0,1,0) = " & flexNum
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 50, flex)
    cmd.Parameters.Append param
End If

If Len(exclude) > 0 Then
    SQL = SQL & " AND iif(instr(?,[mvp_pos])>0,1,0) + " & _
                     "iif(instr(?,[p2_pos])>0,1,0) + " & _
                     "iif(instr(?,[p3_pos])>0,1,0) + " & _
                     "iif(instr(?,[p4_pos])>0,1,0) + " & _
                     "iif(instr(?,[p5_pos])>0,1,0) + " & _
                     "iif(instr(?,[p6_pos])>0,1,0) = 0"
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
    Set param = cmd.CreateParameter("", adVarchar, adParamInput, 100, exclude)
    cmd.Parameters.Append param
End If

Set cmd.ActiveConnection = conn
cmd.CommandText = SQL

rs.CursorLocation = adUseClient

rs.Open cmd
'recCount = rs.RecordCount

If rs.RecordCount > 0 Then
    data = Application.Transpose(rs.GetRows)
    If rs.RecordCount > 1 Then
        random = WorksheetFunction.RandBetween(1, (UBound(data)))
        For i = 1 To UBound(arr, 2)
            arr(0, i) = data(random, i)
        Next
    Else
        For i = 1 To UBound(arr, 2)
            arr(0, i) = data(i)
        Next
    End If
    
    Range("F" & Cells(Rows.Count, 6).End(xlUp).row + 1 & ":S" & Cells(Rows.Count, 6).End(xlUp).row + 1).Value = arr
    Range("F1").CurrentRegion.EntireColumn.AutoFit

Else
    MsgBox "No Lineups Found"
End If

Range("G" & ws.Cells(Rows.Count, 6).End(xlUp).row).Activate

rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing

End Sub
Sub clearRandomLineup()
With Worksheets("Random Lineup")
    .Range("F2:S" & .Cells(Rows.Count, 6).End(xlDown).row).ClearContents
    .Range("F2").Activate
End With
End Sub
Sub saveLineup()
'Dim wb As Workbook
Dim ws As Worksheet
Dim strMVP As String
Dim strPositions As String
Dim arr
Dim i As Long

'Set wb = ActiveWorkbook
Set ws = ActiveSheet

With ws
    strMVP = .Range("C2")
    strPositions = .Range("C3") & " " & .Range("C4") & " " & .Range("C5") & " " & .Range("C6") & " " & .Range("C7")
End With
 
arr = Sheets("Tier").Range("F2:K" & Sheets("Tier").Cells(Rows.Count, 6).End(xlUp).row).Value

For i = 1 To UBound(arr)
    If InStr(strMVP, arr(i, 1)) * _
       InStr(strPositions, arr(i, 2)) * _
       InStr(strPositions, arr(i, 3)) * _
       InStr(strPositions, arr(i, 4)) * _
       InStr(strPositions, arr(i, 5)) * _
       InStr(strPositions, arr(i, 6)) Then
       
        Sheets("Tier").Cells(i + 1, 12) = ws.Range("B2")
        Exit For
    End If
    If i = UBound(arr) Then
        MsgBox "No Lineups Found"
    End If
Next i

End Sub

Sub clearLineup()
    With Worksheets("Lineup Manager")
        .Range("B2").Activate
        .Range("B2", "C2:C7").ClearContents
    End With
End Sub
