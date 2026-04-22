Attribute VB_Name = "modDBManip"
Option Explicit

Public Sub ExportDatabaseSchema()

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim fileNum As Integer
    Dim outputPath As String
    
    Set db = CurrentDb
    fileNum = FreeFile
    
    outputPath = CurrentProject.Path & "\DatabaseSchema.txt"
    
    Open outputPath For Output As #fileNum
    
    Print #fileNum, "ACCESS DATABASE SCHEMA"
    Print #fileNum, "======================="
    Print #fileNum, ""
    
    For Each tdf In db.TableDefs
        
        ' Skip system tables
        If Not IsSystemTable(tdf.Name) Then
            
            Print #fileNum, "TABLE: " & tdf.Name
            
            For Each fld In tdf.Fields
                
                Print #fileNum, "  - " & fld.Name & _
                    " | Type=" & FieldTypeName(fld.Type) & _
                    " | Size=" & fld.Size & _
                    " | Required=" & fld.Required & _
                    " | Auto=" & IsAutoNumber(fld)
                    
            Next fld
            
            Print #fileNum, ""
            
        End If
        
    Next tdf
    
    Close #fileNum
    
    MsgBox "Schema exported to:" & vbCrLf & outputPath, vbInformation

End Sub



Private Function IsSystemTable(ByVal tableName As String) As Boolean
    IsSystemTable = (Left(tableName, 4) = "MSys")
End Function

Private Function IsAutoNumber(fld As DAO.Field) As Boolean
    IsAutoNumber = (fld.Attributes And dbAutoIncrField) <> 0
End Function



Private Function FieldTypeName(ByVal typeCode As Integer) As String

    Select Case typeCode
        Case dbBoolean: FieldTypeName = "YESNO"
        Case dbByte: FieldTypeName = "BYTE"
        Case dbInteger: FieldTypeName = "INTEGER"
        Case dbLong: FieldTypeName = "LONG"
        Case dbCurrency: FieldTypeName = "CURRENCY"
        Case dbSingle: FieldTypeName = "SINGLE"
        Case dbDouble: FieldTypeName = "DOUBLE"
        Case dbDate: FieldTypeName = "DATETIME"
        Case dbText: FieldTypeName = "TEXT"
        Case dbLongBinary: FieldTypeName = "BINARY"
        Case dbMemo: FieldTypeName = "MEMO"
        Case Else: FieldTypeName = "UNKNOWN(" & typeCode & ")"
    End Select

End Function


