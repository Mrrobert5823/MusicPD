Attribute VB_Name = "modRenderDebug"
Option Compare Database
Option Explicit

Public Sub DebugPrintScoreStructure(Optional ByVal pScoreID As Long = 1)

    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset

    Dim curMeasure As Variant
    Dim curStaff As String
    Dim curVoice As String

    Set db = CurrentDb
    Set qd = db.QueryDefs("qryScoreRenderNoteEvents")
    qd.Parameters("pScoreID") = pScoreID
    Set rs = qd.OpenRecordset(dbOpenSnapshot)

    If rs.EOF Then
        Debug.Print "No rows found."
        GoTo CleanExit
    End If

    curMeasure = Null
    curStaff = ""
    curVoice = ""

    Debug.Print String(70, "=")
    Debug.Print "Score Structure for ScoreID=" & pScoreID
    Debug.Print String(70, "=")

    Do While Not rs.EOF

        If Nz(curMeasure, 0) <> Nz(rs!MeasureNumber, 0) Then
            curMeasure = rs!MeasureNumber
            curStaff = ""
            curVoice = ""

            Debug.Print
            Debug.Print "Measure " & rs!MeasureNumber
            Debug.Print String(30, ".")
        End If

        If curStaff <> Nz(rs!StaffName, "") Then
            curStaff = Nz(rs!StaffName, "")
            curVoice = ""
            Debug.Print "  Staff: " & curStaff
        End If

        If curVoice <> Nz(rs!VoiceName, "") Then
            curVoice = Nz(rs!VoiceName, "")
            Debug.Print "    Voice: " & curVoice
        End If

        Debug.Print "      Beat " & _
                    Format(Nz(rs!beatStart, 0), "0.##") & _
                    " | Dur " & Nz(rs!durationName, "") & _
                    " | Note " & Nz(rs!pitchLetter, "Rest") & _
                    Nz(rs!Accidental, "") & _
                    Nz(rs!octaveNum, "")

        rs.MoveNext
    Loop

CleanExit:
    On Error Resume Next
    rs.Close
    Set rs = Nothing
    Set qd = Nothing
    Set db = Nothing

End Sub

