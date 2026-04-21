Attribute VB_Name = "modSvgRender"
Option Compare Database
Option Explicit

Private Type TRenderNote
    MeasureNumber As Long
    StaffName As String
    VoiceName As String
    beatStart As Double
    durationName As String
    pitchLetter As String
    Accidental As String
    octaveNum As Long
End Type

Public Sub ExportScoreSvgHtml(Optional ByVal pScoreID As Long = 1)

    Dim htmlPath As String
    Dim html As String
    
    htmlPath = CurrentProject.Path & "\score_preview.html"
    
    html = BuildScoreHtml(pScoreID)
    SaveTextFile htmlPath, html
    
    Application.FollowHyperlink htmlPath
    
    Debug.Print "Preview written to:" & vbCrLf & htmlPath

End Sub

Private Function BuildScoreHtml(ByVal pScoreID As Long) As String

    Dim rsHeader As DAO.Recordset
    Dim titleText As String
    Dim subtitleText As String
    Dim tempoVal As Variant
    Dim tsTop As Variant
    Dim tsBottom As Variant
    Dim keySig As String
    
    Set rsHeader = CurrentDb.OpenRecordset( _
        "SELECT ScoreTitle, Subtitle, Tempo, TimeSigTop, TimeSigBottom, KeySignature " & _
        "FROM tblScore WHERE ScoreID = " & pScoreID, dbOpenSnapshot)
    
    If rsHeader.EOF Then
        BuildScoreHtml = "<html><body><h1>Score not found.</h1></body></html>"
        rsHeader.Close
        Set rsHeader = Nothing
        Exit Function
    End If
    
    titleText = Nz(rsHeader!ScoreTitle, "")
    subtitleText = Nz(rsHeader!Subtitle, "")
    tempoVal = Nz(rsHeader!Tempo, "")
    tsTop = Nz(rsHeader!TimeSigTop, "")
    tsBottom = Nz(rsHeader!TimeSigBottom, "")
    keySig = Nz(rsHeader!KeySignature, "")
    
    rsHeader.Close
    Set rsHeader = Nothing
    
    BuildScoreHtml = ""
    BuildScoreHtml = BuildScoreHtml & "<!DOCTYPE html>"
    BuildScoreHtml = BuildScoreHtml & "<html><head><meta charset='utf-8'>"
    BuildScoreHtml = BuildScoreHtml & "<title>" & HtmlEncode(titleText) & "</title>"
    BuildScoreHtml = BuildScoreHtml & "<style>"
    BuildScoreHtml = BuildScoreHtml & "body{font-family:Arial,Helvetica,sans-serif;margin:24px;background:#fff;color:#000;}"
    BuildScoreHtml = BuildScoreHtml & ".title{font-size:28px;font-weight:700;text-align:center;margin-bottom:6px;}"
    BuildScoreHtml = BuildScoreHtml & ".subtitle{font-size:16px;text-align:center;margin-bottom:10px;}"
    BuildScoreHtml = BuildScoreHtml & ".meta{font-size:14px;margin-bottom:18px;}"
    BuildScoreHtml = BuildScoreHtml & "svg{display:block;}"
    BuildScoreHtml = BuildScoreHtml & "</style></head><body>"
    
    BuildScoreHtml = BuildScoreHtml & "<div class='title'>" & HtmlEncode(titleText) & "</div>"
    If Len(subtitleText) > 0 Then
        BuildScoreHtml = BuildScoreHtml & "<div class='subtitle'>" & HtmlEncode(subtitleText) & "</div>"
    End If
    
    BuildScoreHtml = BuildScoreHtml & "<div class='meta'>"
    BuildScoreHtml = BuildScoreHtml & "Tempo: " & HtmlEncode(CStr(tempoVal))
    BuildScoreHtml = BuildScoreHtml & " | Time: " & HtmlEncode(CStr(tsTop)) & "/" & HtmlEncode(CStr(tsBottom))
    BuildScoreHtml = BuildScoreHtml & " | Key: " & HtmlEncode(keySig)
    BuildScoreHtml = BuildScoreHtml & "</div>"
    
    BuildScoreHtml = BuildScoreHtml & BuildScoreSvg(pScoreID)
    
    BuildScoreHtml = BuildScoreHtml & "</body></html>"

End Function

 
Private Sub DrawFiveLineStaff(ByRef svg As String, ByVal x As Double, ByVal yTop As Double, ByVal widthVal As Double, ByVal lineGap As Double)

    Dim i As Long
    Dim yVal As Double
    
    For i = 0 To 4
        yVal = yTop + (i * lineGap)
        
        svg = svg & "<line x1='" & CStr(x) & _
                    "' y1='" & CStr(yVal) & _
                    "' x2='" & CStr(x + widthVal) & _
                    "' y2='" & CStr(yVal) & _
                    "' stroke='blue' stroke-width='4' />"
    Next i

End Sub
 

Private Sub DrawTimeSignature(ByRef svg As String, ByVal x As Double, ByVal yTop As Double, ByVal tsTop As Long, ByVal tsBottom As Long, ByVal lineGap As Double)
    svg = svg & "<text x='" & FormatNum(x) & "' y='" & FormatNum(yTop + lineGap * 1.5) & "' font-size='22'>" & tsTop & "</text>"
    svg = svg & "<text x='" & FormatNum(x) & "' y='" & FormatNum(yTop + lineGap * 3.5) & "' font-size='22'>" & tsBottom & "</text>"
End Sub

Private Function GetNoteX(ByVal measureStartX As Double, ByVal noteSpacingWidth As Double, ByVal beatStart As Double) As Double
    Dim innerLeft As Double
    innerLeft = measureStartX + 50
    
    Select Case beatStart
        Case 1
            GetNoteX = innerLeft + 10
        Case 2
            GetNoteX = innerLeft + (noteSpacingWidth * 0.33)
        Case 3
            GetNoteX = innerLeft + (noteSpacingWidth * 0.66)
        Case 4
            GetNoteX = innerLeft + noteSpacingWidth
        Case Else
            GetNoteX = innerLeft + ((beatStart - 1) / 3) * noteSpacingWidth
    End Select
End Function

Private Function GetTrebleNoteY(ByVal pitch As String, ByVal octave As Long, _
                                ByVal staffTop As Double, ByVal lineGap As Double) As Double

    ' Treble clef reference: bottom line = E4
    Dim refPitch As String
    Dim refOctave As Long
    
    refPitch = "E"
    refOctave = 4
    
    Dim steps As Long
    steps = GetNoteStepOffset(pitch, octave, refPitch, refOctave)
    
    ' Each step = half a line spacing
    GetTrebleNoteY = staffTop + (4 * lineGap) - (steps * (lineGap / 2))

End Function



Private Function StaffYFromReference(ByVal pitchLetter As String, ByVal octaveNum As Long, _
                                     ByVal refY As Double, ByVal halfStepY As Double, _
                                     ByVal refLetter As String, ByVal refOctave As Long) As Double
    Dim noteIndex As Long
    Dim refIndex As Long
    Dim diff As Long
    
    noteIndex = NaturalNoteIndex(UCase$(pitchLetter), octaveNum)
    refIndex = NaturalNoteIndex(UCase$(refLetter), refOctave)
    diff = noteIndex - refIndex
    
    StaffYFromReference = refY - (diff * halfStepY)
End Function

Private Function NaturalNoteIndex(ByVal pitchLetter As String, ByVal octaveNum As Long) As Long
    Dim stepVal As Long
    
    Select Case pitchLetter
        Case "C": stepVal = 0
        Case "D": stepVal = 1
        Case "E": stepVal = 2
        Case "F": stepVal = 3
        Case "G": stepVal = 4
        Case "A": stepVal = 5
        Case "B": stepVal = 6
        Case Else: stepVal = 0
    End Select
    
    NaturalNoteIndex = octaveNum * 7 + stepVal
End Function

Private Sub DrawNote(ByRef svg As String, ByVal x As Double, ByVal y As Double, ByVal durationName As String, ByVal stemUp As Boolean)

    Dim rx As Double
    Dim ry As Double
    Dim stemX As Double
    Dim stemY2 As Double
    Dim fillColor As String
    Dim needsStem As Boolean
    
    rx = 7
    ry = 5
    
    Select Case LCase$(Trim$(durationName))
        Case "whole"
            fillColor = "white"
            needsStem = False
        
        Case "half"
            fillColor = "white"
            needsStem = True
        
        Case Else
            fillColor = "black"
            needsStem = True
    End Select
    
    svg = svg & "<ellipse cx='" & FormatNum(x) & "' cy='" & FormatNum(y) & _
                "' rx='" & rx & "' ry='" & ry & "' fill='" & fillColor & _
                "' stroke='black' stroke-width='1.8' />"
    
    If needsStem Then
        If stemUp Then
            stemX = x + rx
            stemY2 = y - 34
            svg = svg & "<line x1='" & FormatNum(stemX) & "' y1='" & FormatNum(y) & _
                        "' x2='" & FormatNum(stemX) & "' y2='" & FormatNum(stemY2) & _
                        "' stroke='black' stroke-width='1.5' />"
        Else
            stemX = x - rx
            stemY2 = y + 34
            svg = svg & "<line x1='" & FormatNum(stemX) & "' y1='" & FormatNum(y) & _
                        "' x2='" & FormatNum(stemX) & "' y2='" & FormatNum(stemY2) & _
                        "' stroke='black' stroke-width='1.5' />"
        End If
    End If

End Sub

Private Function FormatNum(ByVal d As Double) As String
    FormatNum = Replace(Format(d, "0.###"), ",", ".")
End Function

Public Sub SaveTextFile(ByVal filePath As String, ByVal fileText As String)
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Output As #ff
    Print #ff, fileText
    Close #ff
End Sub

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function

Private Function GetBeatX(ByVal measureLeft As Double, ByVal beatStart As Double, ByVal measureWidth As Double) As Double

    Dim innerLeft As Double
    innerLeft = measureLeft + 40
    
    Select Case beatStart
        Case 1
            GetBeatX = innerLeft + 10
        Case 2
            GetBeatX = innerLeft + (measureWidth * 0.33)
        Case 3
            GetBeatX = innerLeft + (measureWidth * 0.66)
        Case 4
            GetBeatX = innerLeft + (measureWidth * 0.9)
        Case Else
            GetBeatX = innerLeft
    End Select

End Function

Private Function BuildVoiceStaffSvg(ByVal pScoreID As Long, ByVal pVoiceName As String, _
                                    ByVal pStaffName As String, ByVal pSectionTop As Double, _
                                    ByVal pCaption As String, ByVal pIsTreble As Boolean) As String

    Dim qd As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim svg As String
    Dim sqlText As String
    
    Dim sectionLeft As Double
    Dim sectionWidth As Double
    Dim sectionHeight As Double
    
    Dim staffLeft As Double
    Dim staffTop As Double
    Dim measureWidth As Double
    Dim lineGap As Double
    Dim noteSpacingWidth As Double
    Dim measureCount As Long
    
    Dim xBar As Double
    Dim xMeasureStart As Double
    Dim noteX As Double
    Dim noteY As Double
    Dim i As Long
    
    sectionLeft = 20
    sectionWidth = 1220
    sectionHeight = 220
    
    staffLeft = 120
    staffTop = pSectionTop + 70
    measureWidth = 240
    lineGap = 12
    noteSpacingWidth = measureWidth - 70
    measureCount = 4
    
    svg = ""
    
    ' section border so we can see the block clearly
    svg = svg & "<rect x='" & FormatNum(sectionLeft) & _
                "' y='" & FormatNum(pSectionTop) & _
                "' width='" & FormatNum(sectionWidth) & _
                "' height='" & FormatNum(sectionHeight) & _
                "' fill='none' stroke='#cccccc' stroke-width='1' />"
    
    ' caption
    svg = svg & "<text x='20' y='" & FormatNum(pSectionTop + 28) & _
                "' font-size='22' font-weight='bold'>" & HtmlEncode(pCaption) & "</text>"
    
    ' optional clef placeholder text
    If pIsTreble Then
        svg = svg & "<text x='20' y='" & FormatNum(staffTop + 34) & _
                    "' font-size='42'>Treble</text>"
    Else
        svg = svg & "<text x='20' y='" & FormatNum(staffTop + 34) & _
                    "' font-size='42'>Bass</text>"
    End If
    
    ' staff lines
    DrawFiveLineStaff svg, staffLeft, staffTop, measureWidth * measureCount, lineGap
    Dim i As Long
    Dim xBar As Double

    For i = 0 To measureCount
        xBar = staffLeft + (i * measureWidth)
        
        svg = svg & "<line x1='" & CStr(xBar) & _
                    "' y1='" & CStr(staffTop) & _
                    "' x2='" & CStr(xBar) & _
                    "' y2='" & CStr(staffTop + 4 * lineGap) & _
                    "' stroke='black' stroke-width='2' />"
    Next i
    
    ' time signature
    DrawTimeSignature svg, staffLeft + 8, staffTop, 4, 4, lineGap
    
    ' bar lines
    For i = 0 To measureCount
        xBar = staffLeft + (i * measureWidth)
        svg = svg & "<line x1='" & FormatNum(xBar) & _
                    "' y1='" & FormatNum(staffTop) & _
                    "' x2='" & FormatNum(xBar) & _
                    "' y2='" & FormatNum(staffTop + 4 * lineGap) & _
                    "' stroke='black' stroke-width='1.8' />"
    Next i
    
    sqlText = ""
    sqlText = sqlText & "SELECT m.MeasureNumber, ne.BeatStart, ne.DurationName, "
    sqlText = sqlText & "ne.PitchLetter, ne.Accidental, ne.OctaveNum, ne.EventOrder, ne.NoteEventID "
    sqlText = sqlText & "FROM (tblMeasure AS m "
    sqlText = sqlText & "INNER JOIN tblNoteEvent AS ne ON m.MeasureID = ne.MeasureID) "
    sqlText = sqlText & "INNER JOIN tblStaff AS st ON ne.StaffID = st.StaffID "
    sqlText = sqlText & "WHERE m.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND st.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.VoiceName = '" & Replace(pVoiceName, "'", "''") & "' "
    sqlText = sqlText & "AND st.StaffName = '" & Replace(pStaffName, "'", "''") & "' "
    sqlText = sqlText & "AND Nz(ne.IsVisible, True) = True "
    sqlText = sqlText & "AND Nz(st.IsVisible, True) = True "
    sqlText = sqlText & "ORDER BY m.MeasureNumber, ne.BeatStart, ne.EventOrder, ne.NoteEventID;"
    
    Set qd = CurrentDb.CreateQueryDef("")
    qd.SQL = sqlText
    Set rs = qd.OpenRecordset(dbOpenSnapshot)
    
    Do While Not rs.EOF
        
        xMeasureStart = staffLeft + ((Nz(rs!MeasureNumber, 1) - 1) * measureWidth)
        noteX = GetNoteX(xMeasureStart, noteSpacingWidth, Nz(rs!beatStart, 1))
        
        If pIsTreble Then
            noteY = GetTrebleNoteY(Nz(rs!pitchLetter, ""), Nz(rs!octaveNum, 4), staffTop, lineGap)
            DrawNote svg, noteX, noteY, Nz(rs!durationName, "quarter"), True
        Else
            noteY = GetBassNoteY(Nz(rs!pitchLetter, ""), Nz(rs!octaveNum, 3), staffTop, lineGap)
            DrawNote svg, noteX, noteY, Nz(rs!durationName, "whole"), False
        End If
        
        ' red debug dot
        svg = svg & "<circle cx='" & FormatNum(noteX) & _
                    "' cy='" & FormatNum(noteY) & _
                    "' r='2.5' fill='red' />"
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set qd = Nothing
    
    BuildVoiceStaffSvg = svg

End Function


Private Function BuildScoreSvg(ByVal pScoreID As Long) As String

    Dim svg As String
    Dim i As Long
    Dim xBar As Double
    
    Dim staffLeft As Double
    Dim staffWidth As Double
    Dim lineGap As Double
    Dim measureWidth As Double
    
    Dim lhTop As Double
    Dim rh1Top As Double
    Dim rh2Top As Double
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    
    Dim measureLeft As Double
    Dim noteX As Double
    Dim noteY As Double
    
    staffLeft = 120
    staffWidth = 960
    lineGap = 12
    measureWidth = 240
    
    lhTop = 120
    rh1Top = 390
    rh2Top = 660
    
    svg = ""
    svg = svg & "<svg xmlns='http://www.w3.org/2000/svg' width='1300' height='900'>"
    
    Set db = CurrentDb
    
    ' =========================
    ' SECTION 1 - LEFT HAND
    ' =========================
    svg = svg & "<rect x='20' y='40' width='1220' height='220' fill='none' stroke='gray' stroke-width='1' />"
    svg = svg & "<text x='20' y='68' font-size='22' font-weight='bold'>LEFT HAND</text>"
    
    DrawFiveLineStaff svg, staffLeft, lhTop, staffWidth, lineGap
    DrawBassClef svg, 70, lhTop
    
    For i = 0 To 4
        xBar = staffLeft + (i * measureWidth)
        svg = svg & "<line x1='" & CStr(xBar) & _
                    "' y1='" & CStr(lhTop) & _
                    "' x2='" & CStr(xBar) & _
                    "' y2='" & CStr(lhTop + (4 * lineGap)) & _
                    "' stroke='black' stroke-width='2' />"
    Next i
    DrawChordSymbols svg, pScoreID, staffLeft, lhTop, measureWidth
    
    sqlText = ""
    sqlText = sqlText & "SELECT m.MeasureNumber, ne.BeatStart, ne.PitchLetter, ne.OctaveNum, ne.DurationName "
    sqlText = sqlText & "FROM (tblMeasure AS m "
    sqlText = sqlText & "INNER JOIN tblNoteEvent AS ne ON m.MeasureID = ne.MeasureID) "
    sqlText = sqlText & "WHERE m.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.VoiceName = 'LH_Shells' "
    sqlText = sqlText & "ORDER BY m.MeasureNumber, ne.BeatStart, ne.EventOrder, ne.NoteEventID;"
    
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    
    Do While Not rs.EOF
        measureLeft = staffLeft + ((rs!MeasureNumber - 1) * measureWidth)
        noteX = GetBeatX(measureLeft, Nz(rs!beatStart, 1), measureWidth)
        noteY = GetBassNoteY(rs!pitchLetter, rs!octaveNum, lhTop, lineGap)
        
        DrawLedgerLines svg, noteX, noteY, lhTop, lineGap
        DrawNoteWithStem svg, noteX, noteY, Nz(rs!durationName, "whole"), _
                         ShouldStemUp(noteY, lhTop, lineGap)
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    ' =========================
    ' SECTION 2 - RIGHT HAND OPTION 1
    ' =========================
    svg = svg & "<rect x='20' y='310' width='1220' height='220' fill='none' stroke='gray' stroke-width='1' />"
    svg = svg & "<text x='20' y='338' font-size='22' font-weight='bold'>RIGHT HAND OPTION 1</text>"
    
    DrawFiveLineStaff svg, staffLeft, rh1Top, staffWidth, lineGap
    DrawTrebleClef svg, 62, rh1Top
    
    For i = 0 To 4
        xBar = staffLeft + (i * measureWidth)
        svg = svg & "<line x1='" & CStr(xBar) & _
                    "' y1='" & CStr(rh1Top) & _
                    "' x2='" & CStr(xBar) & _
                    "' y2='" & CStr(rh1Top + (4 * lineGap)) & _
                    "' stroke='black' stroke-width='2' />"
    Next i
    DrawChordSymbols svg, pScoreID, staffLeft, rh1Top, measureWidth
    
    sqlText = ""
    sqlText = sqlText & "SELECT m.MeasureNumber, ne.BeatStart, ne.PitchLetter, ne.OctaveNum, ne.DurationName "
    sqlText = sqlText & "FROM (tblMeasure AS m "
    sqlText = sqlText & "INNER JOIN tblNoteEvent AS ne ON m.MeasureID = ne.MeasureID) "
    sqlText = sqlText & "WHERE m.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.VoiceName = 'RH_Option1' "
    sqlText = sqlText & "ORDER BY m.MeasureNumber, ne.BeatStart, ne.EventOrder, ne.NoteEventID;"
    
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    
    Do While Not rs.EOF
        measureLeft = staffLeft + ((rs!MeasureNumber - 1) * measureWidth)
        noteX = GetBeatX(measureLeft, Nz(rs!beatStart, 1), measureWidth)
        noteY = GetTrebleNoteY(rs!pitchLetter, rs!octaveNum, rh1Top, lineGap)
        
        DrawLedgerLines svg, noteX, noteY, rh1Top, lineGap
        DrawNoteWithStem svg, noteX, noteY, Nz(rs!durationName, "whole"), _
                         ShouldStemUp(noteY, rh1Top, lineGap)
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    
    ' =========================
    ' SECTION 3 - RIGHT HAND OPTION 2
    ' =========================
    svg = svg & "<rect x='20' y='580' width='1220' height='220' fill='none' stroke='gray' stroke-width='1' />"
    svg = svg & "<text x='20' y='608' font-size='22' font-weight='bold'>RIGHT HAND OPTION 2</text>"
    
    DrawFiveLineStaff svg, staffLeft, rh2Top, staffWidth, lineGap
    DrawTrebleClef svg, 62, rh2Top
    
    For i = 0 To 4
        xBar = staffLeft + (i * measureWidth)
        svg = svg & "<line x1='" & CStr(xBar) & _
                    "' y1='" & CStr(rh2Top) & _
                    "' x2='" & CStr(xBar) & _
                    "' y2='" & CStr(rh2Top + (4 * lineGap)) & _
                    "' stroke='black' stroke-width='2' />"
    Next i
    DrawChordSymbols svg, pScoreID, staffLeft, rh2Top, measureWidth
    
    sqlText = ""
    sqlText = sqlText & "SELECT m.MeasureNumber, ne.BeatStart, ne.PitchLetter, ne.OctaveNum, ne.DurationName "
    sqlText = sqlText & "FROM (tblMeasure AS m "
    sqlText = sqlText & "INNER JOIN tblNoteEvent AS ne ON m.MeasureID = ne.MeasureID) "
    sqlText = sqlText & "WHERE m.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "AND ne.VoiceName = 'RH_Option2' "
    sqlText = sqlText & "ORDER BY m.MeasureNumber, ne.BeatStart, ne.EventOrder, ne.NoteEventID;"
    
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    
    Do While Not rs.EOF
        measureLeft = staffLeft + ((rs!MeasureNumber - 1) * measureWidth)
        noteX = GetBeatX(measureLeft, Nz(rs!beatStart, 1), measureWidth)
        noteY = GetTrebleNoteY(rs!pitchLetter, rs!octaveNum, rh2Top, lineGap)
        
        DrawLedgerLines svg, noteX, noteY, rh2Top, lineGap
        DrawNoteWithStem svg, noteX, noteY, Nz(rs!durationName, "whole"), _
                         ShouldStemUp(noteY, rh2Top, lineGap)
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    
    svg = svg & "</svg>"
    
    BuildScoreSvg = svg

End Function

Private Function GetNoteStepOffset(ByVal pitch As String, ByVal octave As Long, _
                                   ByVal refPitch As String, ByVal refOctave As Long) As Long

    Dim noteOrder As String
    noteOrder = "CDEFGAB"
    
    Dim pitchIndex As Long
    Dim refIndex As Long
    
    pitchIndex = InStr(1, noteOrder, pitch, vbTextCompare) - 1
    refIndex = InStr(1, noteOrder, refPitch, vbTextCompare) - 1
    
    GetNoteStepOffset = ((octave - refOctave) * 7) + (pitchIndex - refIndex)

End Function
Private Function GetBassNoteY(ByVal pitch As String, ByVal octave As Long, _
                             ByVal staffTop As Double, ByVal lineGap As Double) As Double

    ' Bass clef reference: bottom line = G2
    Dim refPitch As String
    Dim refOctave As Long
    
    refPitch = "G"
    refOctave = 2
    
    Dim steps As Long
    steps = GetNoteStepOffset(pitch, octave, refPitch, refOctave)
    
    ' Each step = half a line spacing
    GetBassNoteY = staffTop + (4 * lineGap) - (steps * (lineGap / 2))

End Function


Private Sub DrawLedgerLines(ByRef svg As String, ByVal x As Double, ByVal y As Double, _
                            ByVal staffTop As Double, ByVal lineGap As Double)

    Dim staffBottom As Double
    staffBottom = staffTop + (4 * lineGap)
    
    Dim ledgerY As Double
    
    ' Below staff
    If y > staffBottom Then
        ledgerY = staffBottom + (lineGap / 2)
        
        Do While ledgerY <= y
            svg = svg & "<line x1='" & CStr(x - 14) & _
                        "' y1='" & CStr(ledgerY) & _
                        "' x2='" & CStr(x + 14) & _
                        "' y2='" & CStr(ledgerY) & _
                        "' stroke='black' stroke-width='2' />"
            ledgerY = ledgerY + (lineGap / 2)
        Loop
    End If
    
    ' Above staff
    If y < staffTop Then
        ledgerY = staffTop - (lineGap / 2)
        
        Do While ledgerY >= y
            svg = svg & "<line x1='" & CStr(x - 10) & _
                        "' y1='" & CStr(ledgerY) & _
                        "' x2='" & CStr(x + 10) & _
                        "' y2='" & CStr(ledgerY) & _
                        "' stroke='black' stroke-width='2' />"
            ledgerY = ledgerY - (lineGap / 2)
        Loop
    End If

End Sub

Private Sub DrawNoteWithStem(ByRef svg As String, ByVal x As Double, ByVal y As Double, _
                             ByVal durationName As String, ByVal stemUp As Boolean)

    Dim rx As Double
    Dim ry As Double
    Dim stemX As Double
    Dim stemY2 As Double
    Dim needsStem As Boolean
    
    rx = 7
    ry = 5
    
    svg = svg & "<ellipse cx='" & CStr(x) & _
                "' cy='" & CStr(y) & _
                "' rx='" & CStr(rx) & _
                "' ry='" & CStr(ry) & _
                "' fill='white' stroke='black' stroke-width='2' />"
    
    needsStem = (LCase$(Trim$(durationName)) <> "whole")
    
    If needsStem Then
        If stemUp Then
            stemX = x + rx
            stemY2 = y - 34
        Else
            stemX = x - rx
            stemY2 = y + 34
        End If
        
        svg = svg & "<line x1='" & CStr(stemX) & _
                    "' y1='" & CStr(y) & _
                    "' x2='" & CStr(stemX) & _
                    "' y2='" & CStr(stemY2) & _
                    "' stroke='black' stroke-width='2' />"
    End If

End Sub

Private Function ShouldStemUp(ByVal noteY As Double, ByVal staffTop As Double, ByVal lineGap As Double) As Boolean

    Dim middleLine As Double
    middleLine = staffTop + (2 * lineGap)
    
    ' Below middle ? stem up
    ' Above middle ? stem down
    ShouldStemUp = (noteY > middleLine)

End Function


Private Sub DrawTrebleClef(ByRef svg As String, ByVal x As Double, ByVal staffTop As Double)

    Dim cx As Double
    Dim yTop As Double
    
    cx = x + 18
    yTop = staffTop
    
    ' Main vertical spine
    svg = svg & "<path d='M " & CStr(cx) & " " & CStr(yTop - 18) & _
                " C " & CStr(cx - 2) & " " & CStr(yTop + 4) & _
                ", " & CStr(cx - 2) & " " & CStr(yTop + 34) & _
                ", " & CStr(cx) & " " & CStr(yTop + 64) _
                & "' fill='none' stroke='black' stroke-width='2.2' stroke-linecap='round' />"
    
    ' Upper hook
    svg = svg & "<path d='M " & CStr(cx) & " " & CStr(yTop - 10) & _
                " C " & CStr(cx + 12) & " " & CStr(yTop - 4) & _
                ", " & CStr(cx + 10) & " " & CStr(yTop + 10) & _
                ", " & CStr(cx - 2) & " " & CStr(yTop + 12) _
                & "' fill='none' stroke='black' stroke-width='2.2' stroke-linecap='round' />"
    
    ' Main loop around G line area
    svg = svg & "<ellipse cx='" & CStr(cx + 4) & _
                "' cy='" & CStr(yTop + 28) & _
                "' rx='9' ry='7' fill='none' stroke='black' stroke-width='2.2' />"
    
    ' Lower tail
    svg = svg & "<path d='M " & CStr(cx + 2) & " " & CStr(yTop + 34) & _
                " C " & CStr(cx + 14) & " " & CStr(yTop + 42) & _
                ", " & CStr(cx + 10) & " " & CStr(yTop + 58) & _
                ", " & CStr(cx - 4) & " " & CStr(yTop + 56) _
                & "' fill='none' stroke='black' stroke-width='2.2' stroke-linecap='round' />"

End Sub

Private Sub DrawBassClef(ByRef svg As String, ByVal x As Double, ByVal staffTop As Double)

    Dim cx As Double
    Dim cy As Double
    
    cx = x + 16
    cy = staffTop + 24
    
    svg = svg & "<path d='M " & CStr(cx + 10) & " " & CStr(cy - 14) & _
                " C " & CStr(cx - 6) & " " & CStr(cy - 12) & _
                ", " & CStr(cx - 10) & " " & CStr(cy + 8) & _
                ", " & CStr(cx + 6) & " " & CStr(cy + 8) & _
                " C " & CStr(cx + 16) & " " & CStr(cy + 8) & _
                ", " & CStr(cx + 18) & " " & CStr(cy - 2) & _
                ", " & CStr(cx + 8) & " " & CStr(cy - 8) _
                & "' fill='none' stroke='black' stroke-width='2.6' stroke-linecap='round' />"
    
    svg = svg & "<circle cx='" & CStr(cx + 22) & _
                "' cy='" & CStr(cy - 5) & _
                "' r='2.2' fill='black' />"
    
    svg = svg & "<circle cx='" & CStr(cx + 22) & _
                "' cy='" & CStr(cy + 5) & _
                "' r='2.2' fill='black' />"

End Sub

Private Sub DrawChordSymbols(ByRef svg As String, ByVal pScoreID As Long, _
                             ByVal staffLeft As Double, ByVal staffTop As Double, _
                             ByVal measureWidth As Double)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sqlText As String
    
    Dim measureLeft As Double
    Dim chordX As Double
    Dim chordY As Double
    
    Set db = CurrentDb
    
    sqlText = ""
    sqlText = sqlText & "SELECT m.MeasureNumber, cs.ChordText "
    sqlText = sqlText & "FROM (tblMeasure AS m "
    sqlText = sqlText & "INNER JOIN tblChordSymbol AS cs ON m.MeasureID = cs.MeasureID) "
    sqlText = sqlText & "WHERE m.ScoreID = " & pScoreID & " "
    sqlText = sqlText & "ORDER BY m.MeasureNumber, cs.BeatStart, cs.DisplayOrder, cs.ChordSymbolID;"
    
    Set rs = db.OpenRecordset(sqlText, dbOpenSnapshot)
    
    chordY = staffTop - 18
    
    Do While Not rs.EOF
        
        measureLeft = staffLeft + ((rs!MeasureNumber - 1) * measureWidth)
        chordX = measureLeft + (measureWidth / 2)
        
        svg = svg & "<text x='" & CStr(chordX) & _
                    "' y='" & CStr(chordY) & _
                    "' font-size='18' font-weight='bold' text-anchor='middle'>" & _
                    HtmlEncode(Nz(rs!ChordText, "")) & "</text>"
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing

End Sub

