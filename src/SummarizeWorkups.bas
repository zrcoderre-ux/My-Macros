Attribute VB_Name = "SummarizeWorkups"
Option Explicit

' ============================================================
'  CONFIGURATION — edit these before running
' ============================================================
Private Const API_KEY            As String = "API KEY"
Private Const MODEL              As String = "claude-haiku-4-5-20251001"
Private Const ONEDRIVE_ROOT      As String = "C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\"
Private Const SUMMARIES_PATH     As String = "C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\Summaries\"
Private Const LOG_FILE           As String = "C:\Users\ZCoderre\OneDrive - Los Angeles Superior Court\Summaries\SummaryLog.txt"
Private Const MAX_DOCS           As Long = 15      ' 0 = no limit; 3 = test run
Private Const DELAY_SECONDS      As Long = 5         ' seconds between API calls
Private Const RATE_LIMIT_RETRIES As Long = 3         ' retries on rate-limit error

' ============================================================
'  WINDOWS API — for Sleep
' ============================================================
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' ============================================================
'  ENTRY POINT
' ============================================================
Public Sub SummarizeWorkups()
    Dim lProcessed As Long, lSkipped As Long, lErrors As Long
    lProcessed = 0: lSkipped = 0: lErrors = 0

    On Error GoTo FatalError

    ' Suppress all Word dialogs and macro prompts for unattended runs
    Application.DisplayAlerts = wdAlertsNone
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    ' Ensure output folder exists
    If Dir(SUMMARIES_PATH, vbDirectory) = "" Then MkDir SUMMARIES_PATH

    Call ProcessFolder(ONEDRIVE_ROOT, lProcessed, lSkipped, lErrors)

    GoTo CleanUp

FatalError:
    Call AppendLog("FATAL ERROR: " & Err.Number & " - " & Err.Description)

CleanUp:
    Application.DisplayAlerts = wdAlertsAll
    Application.AutomationSecurity = msoAutomationSecurityByUI
    Call AppendLog("=== DONE | Processed: " & lProcessed & " | Skipped: " & lSkipped & " | Errors: " & lErrors & " ===")
    MsgBox "Done." & vbCrLf & "Processed: " & lProcessed & vbCrLf & _
           "Skipped:   " & lSkipped & vbCrLf & "Errors:    " & lErrors, vbInformation, "SummarizeWorkups"
End Sub

' ============================================================
'  RECURSIVE FOLDER WALK
' ============================================================
Private Sub ProcessFolder(sFolder As String, lProcessed As Long, lSkipped As Long, lErrors As Long)
    Dim oFSO    As Object
    Dim oFolder As Object
    Dim oSub    As Object
    Dim oFile   As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    If Not oFSO.FolderExists(sFolder) Then Exit Sub
    Set oFolder = oFSO.GetFolder(sFolder)

    ' Skip excluded folders
    Dim sName As String
    sName = LCase(oFolder.Name)
    Dim aExclude As Variant
    aExclude = Array("documents", "macros", "mcle", "meetings", "pictures", _
                     "templates", "writing samples", "word doc tech ideas", "summaries")
    Dim i As Integer
    For i = 0 To UBound(aExclude)
        If sName = aExclude(i) Then Exit Sub
    Next i
    If InStr(sName, "microsoft") > 0 Then Exit Sub

    ' Process files in this folder
    For Each oFile In oFolder.Files
        If MAX_DOCS > 0 And (lProcessed + lErrors) >= MAX_DOCS Then Exit Sub

        Call ProcessFile(oFile.Path, lProcessed, lSkipped, lErrors)
    Next oFile

    ' Recurse into subfolders
    For Each oSub In oFolder.SubFolders
        If MAX_DOCS > 0 And (lProcessed + lErrors) >= MAX_DOCS Then Exit Sub
        Call ProcessFolder(oSub.Path, lProcessed, lSkipped, lErrors)
    Next oSub
End Sub

' ============================================================
'  SINGLE FILE HANDLER
' ============================================================
Private Sub ProcessFile(sPath As String, lProcessed As Long, lSkipped As Long, lErrors As Long)
    Dim oFSO     As Object
    Dim sFile    As String
    Dim sBase    As String
    Dim sOutFile As String

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    sFile = oFSO.GetFileName(sPath)
    sBase = oFSO.GetBaseName(sPath)

    ' --- Gate 1: must be .docx
    If LCase(oFSO.GetExtensionName(sPath)) <> "docx" Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Gate 2: no temp files
    If Left(sFile, 1) = "~" Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Gate 3: must contain " vs " (case-insensitive, spaces required)
    If InStr(1, LCase(sFile), " vs ") = 0 Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Gate 4: skip files that are already summaries
    If InStr(1, LCase(sFile), "(summary)") > 0 Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Gate 5: skip files created within the last 2 days
    If DateDiff("d", oFSO.GetFile(sPath).DateCreated, Now) < 2 Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Gate 6: summary already exists
    sOutFile = SUMMARIES_PATH & sBase & " (Summary).docx"
    If oFSO.FileExists(sOutFile) Then lSkipped = lSkipped + 1: Exit Sub

    ' --- Open document
    Dim oDoc As Document
    On Error GoTo OpenError
    Set oDoc = Documents.Open(FileName:=sPath, ReadOnly:=True, Visible:=False, _
                              AddToRecentFiles:=False)
    On Error GoTo 0

    ' --- Gate 7: skip mail merge documents
    If oDoc.MailMerge.MainDocumentType <> wdNotAMergeDocument Then
        oDoc.Close SaveChanges:=False
        Call AppendLog("SKIP (mail merge): " & sFile)
        lSkipped = lSkipped + 1
        Exit Sub
    End If

    ' --- Gate 8: 50-page limit
    Dim lPages As Long
    lPages = oDoc.ComputeStatistics(wdStatisticPages)
    If lPages > 50 Then
        oDoc.Close SaveChanges:=False
        Call AppendLog("SKIP (>" & lPages & " pages): " & sFile)
        lSkipped = lSkipped + 1
        Exit Sub
    End If

    ' --- Extract text
    Dim sText As String
    sText = oDoc.Range.text
    oDoc.Close SaveChanges:=False

    ' --- Gate 9: minimum text
    If Len(sText) < 50 Then
        Call AppendLog("SKIP (too short): " & sFile)
        lSkipped = lSkipped + 1
        Exit Sub
    End If

    ' --- Sanitize text: convert Word smart characters to plain ASCII
    '     This prevents garbled â€" style output in summaries
    sText = NormalizeSmartChars(sText)

    ' --- Call API (with rate-limit retry)
    Dim sSummary As String
    Dim nAttempt As Long
    For nAttempt = 1 To RATE_LIMIT_RETRIES + 1
        sSummary = CallClaudeAPI(sText)

        If Left(sSummary, 11) = "RATE_LIMIT:" Then
            If nAttempt <= RATE_LIMIT_RETRIES Then
                Call AppendLog("Rate limit hit for " & sFile & " — waiting 60s (attempt " & nAttempt & ")")
                Sleep 60000
            Else
                Call AppendLog("ERROR (rate limit, all retries exhausted): " & sFile)
                lErrors = lErrors + 1
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next nAttempt

    ' --- Gate 10: API returned empty
    If Len(Trim(sSummary)) = 0 Then
        Call AppendLog("ERROR (empty API response): " & sFile)
        lErrors = lErrors + 1
        Exit Sub
    End If

    ' --- Gate 11: API flagged the document as incomplete / unable to summarize
    If IsIncompleteDocumentResponse(sSummary) Then
        Call AppendLog("SKIP (incomplete source document): " & sFile)
        lSkipped = lSkipped + 1
        Exit Sub
    End If

    ' --- Write summary
    Call WriteSummaryDoc(sSummary, sOutFile)
    Call AppendLog("OK: " & sFile)
    lProcessed = lProcessed + 1

    ' --- Pace API calls
    Sleep DELAY_SECONDS * 1000
    Exit Sub

OpenError:
    Call AppendLog("ERROR (could not open): " & sFile & " — " & Err.Description)
    lErrors = lErrors + 1
End Sub

' ============================================================
'  DETECT INCOMPLETE DOCUMENT RESPONSES
'  Returns True if the API response indicates the source
'  document was blank/incomplete rather than a real summary.
' ============================================================
Private Function IsIncompleteDocumentResponse(sSummary As String) As Boolean
    Dim sLower As String
    sLower = LCase(sSummary)

    ' Phrases that signal Claude could not summarize due to missing content
    Dim aPhrases As Variant
    aPhrases = Array( _
        "cannot provide the requested summary", _
        "i cannot provide", _
        "i'm unable to provide", _
        "i am unable to provide", _
        "document is incomplete", _
        "minute order is incomplete", _
        "the critical sections", _
        "are blank or missing", _
        "please provide the full document", _
        "please provide the complete", _
        "would need the complete", _
        "i would need a complete", _
        "to write an accurate", _
        "insufficient information to summarize", _
        "not enough information to summarize" _
    )

    Dim i As Integer
    For i = 0 To UBound(aPhrases)
        If InStr(sLower, aPhrases(i)) > 0 Then
            IsIncompleteDocumentResponse = True
            Exit Function
        End If
    Next i

    IsIncompleteDocumentResponse = False
End Function

' ============================================================
'  NORMALIZE SMART / UNICODE CHARACTERS TO PLAIN ASCII
'  Converts Word's curly quotes, em-dashes, ellipses, etc.
'  to their plain equivalents so the API output is clean.
' ============================================================
Private Function NormalizeSmartChars(sIn As String) As String
    Dim s As String
    s = sIn

    ' Em dash (U+2014) and en dash (U+2013) -> hyphen
    s = Replace(s, ChrW(8212), "-")   ' em dash
    s = Replace(s, ChrW(8211), "-")   ' en dash

    ' Curly/smart quotes -> straight quotes
    s = Replace(s, ChrW(8220), Chr(34))  ' left double quote
    s = Replace(s, ChrW(8221), Chr(34))  ' right double quote
    s = Replace(s, ChrW(8216), "'")     ' left single quote
    s = Replace(s, ChrW(8217), "'")     ' right single quote

    ' Ellipsis (U+2026) -> three dots
    s = Replace(s, ChrW(8230), "...")

    ' Non-breaking space (U+00A0) -> regular space
    s = Replace(s, ChrW(160), " ")

    ' Bullet (U+2022) -> asterisk
    s = Replace(s, ChrW(8226), "*")

    ' Section sign (U+00A7) - keep as-is; it's meaningful in legal text
    ' Paragraph sign (U+00B6) -> keep
    ' Registered trademark, copyright - strip
    s = Replace(s, ChrW(174), "")     ' (R)
    s = Replace(s, ChrW(169), "")     ' (C)

    NormalizeSmartChars = s
End Function

' ============================================================
'  CLAUDE API CALL
' ============================================================
Private Function CallClaudeAPI(sDocText As String) As String
    Const SUMMARY_PROMPT As String = _
        "You are summarizing a court order for a research attorney." & vbCrLf & _
        "Write exactly one paragraph of three to five sentences." & vbCrLf & _
        "Open with the court's ruling or decision using active voice and naming the parties by role " & _
        "and name (e.g. The Court granted Defendant Jane Smith's motion...)." & vbCrLf & _
        "Do not open with background or case history. Lead with the ruling." & vbCrLf & _
        "Do not use bullet points, headings, or any formatting." & vbCrLf & _
        "Return only the summary paragraph and nothing else." & vbCrLf & _
        "If the document is blank, heavily redacted, or missing the court's ruling and reasoning, " & _
        "respond only with: INCOMPLETE_DOCUMENT"

    Dim sBody As String
    sBody = "{""model"":""" & MODEL & """," & _
            """max_tokens"":1024," & _
            """messages"":[{""role"":""user"",""content"":""" & _
            EscapeJSON(SUMMARY_PROMPT & vbCrLf & vbCrLf & sDocText) & _
            """}]}"

    Dim oHTTP As Object
    Set oHTTP = CreateObject("MSXML2.XMLHTTP")

    On Error GoTo HttpError
    oHTTP.Open "POST", "https://api.anthropic.com/v1/messages", False
    oHTTP.setRequestHeader "Content-Type", "application/json"
    oHTTP.setRequestHeader "x-api-key", API_KEY
    oHTTP.setRequestHeader "anthropic-version", "2023-06-01"
    oHTTP.send sBody
    On Error GoTo 0

    Dim sResp As String
    sResp = oHTTP.responseText

    ' Check for rate limit error
    If InStr(sResp, "rate_limit_error") > 0 Then
        CallClaudeAPI = "RATE_LIMIT: " & sResp
        Exit Function
    End If

    ' Log on non-200
    If oHTTP.Status <> 200 Then
        Call AppendLog("API HTTP " & oHTTP.Status & ": " & Left(sResp, 300))
        CallClaudeAPI = ""
        Exit Function
    End If

    ' Extract summary text
    Dim sText As String
    sText = ExtractJSONField(sResp, "text")

    ' Handle INCOMPLETE_DOCUMENT signal from the model
    If Trim(sText) = "INCOMPLETE_DOCUMENT" Then
        CallClaudeAPI = "INCOMPLETE_DOCUMENT"
        Exit Function
    End If

    ' Strip backslash-escaped quotes the API sometimes inserts
    sText = Replace(sText, Chr(92) & Chr(34), Chr(34))
    ' Unescape JSON escape sequences that may survive into the summary text
    sText = Replace(sText, "\\""", """")
    sText = Replace(sText, "\\\\", "\\")
    sText = Replace(sText, "\\n", " ")
    sText = Replace(sText, "\\r", "")
    sText = Replace(sText, "\\t", " ")
    CallClaudeAPI = Trim(sText)
    Exit Function

HttpError:
    Call AppendLog("HTTP error: " & Err.Description)
    CallClaudeAPI = ""
End Function

' ============================================================
'  JSON HELPERS
' ============================================================
Private Function EscapeJSON(s As String) As String
    Dim sOut As String
    Dim c    As String
    Dim nAsc As Integer
    Dim i    As Long

    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, "/", "\/")
    s = Replace(s, Chr(8), "\b")
    s = Replace(s, Chr(12), "\f")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "\r")
    s = Replace(s, Chr(9), "\t")

    ' Strip any remaining control characters (< ASCII 32)
    sOut = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = "\" Then
            ' Already escaped — include backslash and next char together
            sOut = sOut & c
        Else
            nAsc = asc(c)
            If nAsc >= 32 Then
                sOut = sOut & c
            End If
            ' Characters < 32 that aren't already escaped are silently dropped
        End If
    Next i

    EscapeJSON = sOut
End Function

Private Function ExtractJSONField(sJSON As String, sField As String) As String
    ' Try both "field": "value" (spaced) and "field":"value" (no space)
    Dim sKey    As String
    Dim nStart  As Long
    Dim nEnd    As Long

    sKey = """" & sField & """: """
    nStart = InStr(sJSON, sKey)
    If nStart = 0 Then
        sKey = """" & sField & """:"""
        nStart = InStr(sJSON, sKey)
    End If
    If nStart = 0 Then Exit Function

    nStart = nStart + Len(sKey)
    nEnd = nStart
    Do While nEnd <= Len(sJSON)
        If Mid(sJSON, nEnd, 1) = """" Then
            ' Check it's not escaped
            Dim nBackslashes As Long
            nBackslashes = 0
            Dim j As Long
            j = nEnd - 1
            Do While j >= 1 And Mid(sJSON, j, 1) = "\"
                nBackslashes = nBackslashes + 1
                j = j - 1
            Loop
            If nBackslashes Mod 2 = 0 Then Exit Do
        End If
        nEnd = nEnd + 1
    Loop

    ExtractJSONField = Mid(sJSON, nStart, nEnd - nStart)
End Function

' ============================================================
'  RESTORE SMART QUOTES FOR OUTPUT
' ============================================================
Private Function RestoreSmartQuotes(s As String) As String
    Dim i      As Long
    Dim c      As String
    Dim cPrev  As String
    Dim sOut   As String

    sOut = ""
    cPrev = " "   ' treat start-of-string as preceded by a space

    For i = 1 To Len(s)
        c = Mid(s, i, 1)

        If c = Chr(34) Then   ' straight double quote
            ' Left quote after space, tab, newline, or open bracket/paren
            If InStr(" " & Chr(9) & Chr(10) & Chr(13) & "([{", cPrev) > 0 Then
                sOut = sOut & ChrW(8220)  ' left double quote
            Else
                sOut = sOut & ChrW(8221)  ' right double quote
            End If

        ElseIf c = "'" Then   ' straight apostrophe / single quote
            ' Left single quote after space, tab, newline, or open bracket/paren
            If InStr(" " & Chr(9) & Chr(10) & Chr(13) & "([{", cPrev) > 0 Then
                sOut = sOut & ChrW(8216)  ' left single quote
            Else
                sOut = sOut & ChrW(8217)  ' right single quote / apostrophe
            End If

        Else
            sOut = sOut & c
        End If

        cPrev = c
    Next i

    RestoreSmartQuotes = sOut
End Function

' ============================================================
'  WRITE SUMMARY DOCX
' ============================================================
Private Sub WriteSummaryDoc(sSummary As String, sOutPath As String)
    Dim oDoc As Document
    Set oDoc = Documents.Add(Visible:=False)
    oDoc.Range.text = RestoreSmartQuotes(sSummary)
    oDoc.SaveAs2 FileName:=sOutPath, FileFormat:=wdFormatXMLDocument
    oDoc.Close SaveChanges:=False
End Sub

' ============================================================
'  LOGGING
' ============================================================
Private Sub AppendLog(sMsg As String)
    Dim iFile As Integer
    iFile = FreeFile
    Open LOG_FILE For Append As #iFile
    Print #iFile, Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & sMsg
    Close #iFile
End Sub

