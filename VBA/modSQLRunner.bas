Attribute VB_Name = "modSQLRunner"
Option Compare Database
Option Explicit

'small change saved
'small change 2 saved
'small change 3 saved
'small change 4 saved
'small change 5 saved
'small change 6 saved


Public Sub RunSqlScriptFromFile(ByVal scriptPath As String)

    Dim fso As Object
    Dim ts As Object
    Dim sqlText As String
    Dim statements() As String
    Dim stmt As Variant
    Dim cleanedStmt As String
    Dim db As DAO.Database

    On Error GoTo ErrHandler

    If Len(Dir(scriptPath)) = 0 Then
        MsgBox "Script file not found:" & vbCrLf & scriptPath, vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(scriptPath, 1)

    sqlText = ts.ReadAll
    ts.Close

    statements = Split(sqlText, ";")

    For Each stmt In statements
        cleanedStmt = Trim(CStr(stmt))

        If cleanedStmt <> "" Then
            If Left$(cleanedStmt, 2) <> "--" Then
                db.Execute cleanedStmt, dbFailOnError
            End If
        End If
    Next stmt

    MsgBox "SQL script completed successfully.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error running script:" & vbCrLf & Err.Description, vbCritical
End Sub

Public Sub TestRunRebuildScript()

    Dim scriptPath As String
    scriptPath = CurrentProject.Path & "\DBScripts\BuildTables.sql"
    'C:\Users\mrRob\OneDrive\Documents\2026\Job Companies\DBScripts

    'RunSqlScriptFromFileSafe scriptPath
    RunSqlScriptFromFile scriptPath
    

End Sub
'////////////////////////

 

Public Sub ExportAllUserTableDataToSql()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim exportTables As Collection
    Dim tableName As Variant
    Dim outputPath As String
    Dim ff As Integer

    On Error GoTo ErrHandler

    Set db = CurrentDb
    Set exportTables = GetExportableTables(db)

    outputPath = CurrentProject.Path & "\AllUserTableData.sql"
    ff = FreeFile

    Open outputPath For Output As #ff

    Print #ff, "-- Data export generated " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #ff, ""

    ' Optional: clear existing rows first
    ' Done in reverse order to reduce relationship issues
    WriteDeleteStatements exportTables, ff

    Print #ff, ""

    ' Write inserts in forward order
    For Each tableName In exportTables
        ExportTableRows CStr(tableName), ff
        Print #ff, ""
    Next tableName

    Close #ff

    MsgBox "Export complete: " & outputPath, vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    If ff > 0 Then Close #ff
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

Private Function GetExportableTables(ByVal db As DAO.Database) As Collection

    Dim results As New Collection
    Dim tdf As DAO.TableDef

    For Each tdf In db.TableDefs
        If IsUserTable(tdf) Then
            results.Add tdf.Name
        End If
    Next tdf

    Set GetExportableTables = results
End Function

Private Function IsUserTable(ByVal tdf As DAO.TableDef) As Boolean

    Dim nameText As String
    nameText = tdf.Name

    ' Skip system tables
    If Left$(nameText, 4) = "MSys" Then
        IsUserTable = False
        Exit Function
    End If

    ' Skip temporary/import leftovers
    If Left$(nameText, 1) = "~" Then
        IsUserTable = False
        Exit Function
    End If

    ' Skip linked tables if desired
    ' Comment this block out if you want linked tables included
    If (tdf.Attributes And dbAttachedTable) <> 0 Then
        IsUserTable = False
        Exit Function
    End If

    If (tdf.Attributes And dbAttachedODBC) <> 0 Then
        IsUserTable = False
        Exit Function
    End If

    IsUserTable = True
End Function

Private Sub WriteDeleteStatements(ByVal tableList As Collection, ByVal ff As Integer)

    Dim i As Long

    Print #ff, "-- Optional cleanup"

    For i = tableList.Count To 1 Step -1
        Print #ff, "DELETE FROM [" & tableList(i) & "];"
    Next i
End Sub

Private Sub ExportTableRows(ByVal tableName As String, ByVal ff As Integer)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    Dim fieldList As String
    Dim valueList As String
    Dim sqlLine As String

    On Error GoTo ErrHandler

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM [" & tableName & "]", dbOpenSnapshot)

    Print #ff, "-- Table: " & tableName

    Do While Not rs.EOF

        fieldList = ""
        valueList = ""

        For Each fld In rs.Fields
            If ShouldExportField(fld) Then

                If Len(fieldList) > 0 Then
                    fieldList = fieldList & ", "
                    valueList = valueList & ", "
                End If

                fieldList = fieldList & "[" & fld.Name & "]"
                valueList = valueList & SqlLiteral(rs.Fields(fld.Name).Value, fld.Type)
            End If
        Next fld

        If Len(fieldList) > 0 Then
            sqlLine = "INSERT INTO [" & tableName & "] (" & fieldList & ") VALUES (" & valueList & ");"
            Print #ff, sqlLine
            Debug.Print sqlLine
        End If

        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrHandler:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Err.Raise Err.Number, "ExportTableRows(" & tableName & ")", Err.Description
End Sub

Private Function ShouldExportField(ByVal fld As DAO.Field) As Boolean

    Select Case fld.Type
        Case dbAttachment, dbComplexByte, dbComplexInteger, dbComplexLong, _
             dbComplexSingle, dbComplexDouble, dbComplexGUID
            ShouldExportField = False

        Case Else
            ShouldExportField = True
    End Select
End Function

Private Function SqlLiteral(ByVal v As Variant, ByVal fieldType As Integer) As String

    If IsNull(v) Then
        SqlLiteral = "Null"
        Exit Function
    End If

    Select Case fieldType

        Case dbText, dbMemo, dbChar
            SqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"

        Case dbDate
            SqlLiteral = "#" & Format$(CDate(v), "yyyy-mm-dd hh:nn:ss") & "#"

        Case dbBoolean
            If CBool(v) Then
                SqlLiteral = "True"
            Else
                SqlLiteral = "False"
            End If

        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDecimal
            SqlLiteral = Replace(CStr(v), ",", ".")

        Case dbGUID
            SqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"

        Case Else
            SqlLiteral = "'" & Replace(CStr(v), "'", "''") & "'"
    End Select
End Function



Public Function StripSqlComments(ByVal sqlText As String) As String
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim result As String
    
    lines = Split(sqlText, vbCrLf)
    
    For i = LBound(lines) To UBound(lines)
        lineText = lines(i)
        
        If Trim$(lineText) <> "" Then
            If Left$(Trim$(lineText), 2) <> "--" Then
                result = result & lineText & vbCrLf
            End If
        Else
            result = result & vbCrLf
        End If
    Next i
    
    StripSqlComments = result
End Function

Public Function NormalizeSqlText(ByVal sqlText As String) As String
    Dim s As String
    
    s = Nz(sqlText, "")
    s = Replace(s, vbCrLf, vbLf)
    s = Replace(s, vbCr, vbLf)
    s = Replace(s, vbLf, vbCrLf)
    
    NormalizeSqlText = Trim$(s)
End Function

Public Sub ExecuteSqlText(ByVal sqlText As String)
    On Error GoTo ErrHandler
    
    Dim db As DAO.Database
    Dim oneStmt As String
    
    oneStmt = Trim$(Nz(sqlText, ""))
    
    If Len(oneStmt) = 0 Then
        MsgBox "There is no SQL to run.", vbExclamation
        Exit Sub
    End If
    
    Set db = CurrentDb
    db.Execute oneStmt, dbFailOnError
    
    MsgBox "SQL executed successfully.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "SQL execution failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub


Public Sub RunSqlFromFormTextBox(ByVal frm As Form, ByVal textBoxName As String)
    Dim sqlText As String
    
    sqlText = Nz(frm.Controls(textBoxName).Value, "")
    ExecuteSqlText sqlText
End Sub



Public Function GetFirstWord(ByVal sqlText As String) As String

    Dim s As String
    Dim p As Long

    s = Trim$(sqlText)

    If Len(s) = 0 Then
        GetFirstWord = ""
        Exit Function
    End If

    p = InStr(1, s, " ")
    If p > 0 Then
        GetFirstWord = Left$(s, p - 1)
    Else
        GetFirstWord = s
    End If

End Function

Public Sub RunSqlSelectFromForm(ByVal frm As Form, ByVal textBoxName As String)

    Dim sqlText As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Const TEMP_QUERY_NAME As String = "zz_tmp_SQLRunnerResults"

    On Error GoTo ErrHandler

    sqlText = Trim$(Nz(frm.Controls(textBoxName).Value, ""))

    If Len(sqlText) = 0 Then
        MsgBox "There is no SQL to run.", vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb

    On Error Resume Next
    db.QueryDefs.Delete TEMP_QUERY_NAME
    Err.Clear
    On Error GoTo ErrHandler

    Set qdf = db.CreateQueryDef(TEMP_QUERY_NAME, sqlText)

    DoCmd.OpenQuery TEMP_QUERY_NAME, acViewNormal, acReadOnly
    Exit Sub

ErrHandler:
    MsgBox "SELECT failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub



Public Sub RunSqlBatchFromForm(ByVal frm As Form, ByVal textBoxName As String)

    Dim sqlText As String

    sqlText = Nz(frm.Controls(textBoxName).Value, "")
    RunSqlBatchText sqlText

End Sub

Public Sub RunSqlBatchText(ByVal sqlText As String)

    Dim statements As Collection
    Dim stmt As Variant
    Dim db As DAO.Database
    Dim firstWord As String
    Dim lastSelectSql As String
    Dim executedCount As Long
    Dim stmtText As String

    On Error GoTo ErrHandler

    sqlText = NormalizeSqlText(sqlText)
    sqlText = StripSqlComments(sqlText)

    If Len(Trim$(sqlText)) = 0 Then
        MsgBox "There is no SQL to run.", vbExclamation
        Exit Sub
    End If

    Set statements = SplitSqlStatements(sqlText)

    If statements.Count = 0 Then
        MsgBox "There is no SQL to run.", vbExclamation
        Exit Sub
    End If

    Set db = CurrentDb

    For Each stmt In statements
    stmtText = CStr(stmt)
    stmtText = Replace(stmtText, vbCr, "")
    stmtText = Replace(stmtText, vbLf, "")
    stmtText = Trim$(stmtText)
    
    If Len(stmtText) > 0 Then
        firstWord = UCase$(GetFirstWord(stmtText))

        Select Case firstWord
            Case "SELECT", "TRANSFORM"
                lastSelectSql = stmtText

            Case "INSERT", "UPDATE", "DELETE", "CREATE", "DROP", "ALTER"
                db.Execute stmtText, dbFailOnError
                executedCount = executedCount + 1

            Case Else
                Debug.Print "Skipped fragment: [" & stmtText & "]"
        End Select
    End If
Next stmt

    If Len(lastSelectSql) > 0 Then
        ShowSqlResults lastSelectSql
    Else
        MsgBox executedCount & " statement(s) executed successfully.", vbInformation
    End If

    Exit Sub

ErrHandler:
    MsgBox "SQL batch failed." & vbCrLf & vbCrLf & _
           "Statement:" & vbCrLf & stmtText & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Public Function SplitSqlStatements(ByVal sqlText As String) As Collection

    Dim results As New Collection
    Dim i As Long
    Dim ch As String
    Dim currentText As String
    Dim inSingleQuote As Boolean

    currentText = ""
    inSingleQuote = False

    For i = 1 To Len(sqlText)
        ch = Mid$(sqlText, i, 1)

        Select Case ch
            Case "'"
                currentText = currentText & ch

                If inSingleQuote Then
                    If i < Len(sqlText) And Mid$(sqlText, i + 1, 1) = "'" Then
                        currentText = currentText & "'"
                        i = i + 1
                    Else
                        inSingleQuote = False
                    End If
                Else
                    inSingleQuote = True
                End If

            Case ";"
                If inSingleQuote Then
                    currentText = currentText & ch
                Else
                    If Len(Trim$(currentText)) > 0 Then
                        results.Add Trim$(currentText)
                    End If
                    currentText = ""
                End If

            Case Else
                currentText = currentText & ch
        End Select
    Next i

    If Len(Trim$(currentText)) > 0 Then
        results.Add Trim$(currentText)
    End If

    Set SplitSqlStatements = results
End Function

Public Sub ShowSqlResults(ByVal sqlText As String)

    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Const TEMP_QUERY_NAME As String = "zz_tmp_SQLRunnerResults"

    On Error GoTo ErrHandler

    Set db = CurrentDb

    On Error Resume Next
    db.QueryDefs.Delete TEMP_QUERY_NAME
    Err.Clear
    On Error GoTo ErrHandler

    Set qdf = db.CreateQueryDef(TEMP_QUERY_NAME, sqlText)
    DoCmd.OpenQuery TEMP_QUERY_NAME, acViewNormal, acReadOnly

    Exit Sub

ErrHandler:
    MsgBox "SELECT failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub

 
