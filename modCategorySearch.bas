Attribute VB_Name = "modCategorySearch"
'==============================================================================
' Module  : modCategorySearch
' Purpose : Search engine for the Category Search tool (Classic Outlook, VBA).
'
' HOW IT STAYS FAST
'   1. Folder.GetTable reads only the columns we ask for - mail items are
'      NEVER opened, so 10,000+ mails are fine.
'   2. Categories, sender, subject, attachment flag and date are all checked
'      in VBA on the table rows, with an EXACT category-name match (so "HR"
'      never matches "HRS"). No fragile server-side LIKE filter is needed.
'   3. A DASL filter is used ONLY when you tick "search body".
'
' ENTRY POINT : ShowCategorySearch  (add it to the Quick Access Toolbar)
'==============================================================================
Option Explicit

'--- Win32 time-zone lookup (fallback if Row.UTCToLocal is unavailable) -------
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" (lpTZI As TIME_ZONE_INFORMATION) As Long

'--- DASL property names ------------------------------------------------------
Private Const P_KEYWORDS  As String = "urn:schemas-microsoft-com:office:office#Keywords"
Private Const P_SUBJECT   As String = "urn:schemas:httpmail:subject"
Private Const P_FROMNAME  As String = "urn:schemas:httpmail:fromname"
Private Const P_FROMEMAIL As String = "urn:schemas:httpmail:fromemail"
Private Const P_RECEIVED  As String = "urn:schemas:httpmail:datereceived"
Private Const P_HASATTACH As String = "urn:schemas:httpmail:hasattachment"
Private Const P_BODY      As String = "urn:schemas:httpmail:textdescription"

'--- Result grid column positions --------------------------------------------
Public Const RC_SUBJECT    As Long = 0
Public Const RC_SENDER     As Long = 1
Public Const RC_RECEIVED   As Long = 2
Public Const RC_CATEGORIES As Long = 3
Public Const RC_FOLDER     As Long = 4
Public Const RC_ENTRYID    As Long = 5   'hidden column
Public Const RC_STOREID    As Long = 6   'hidden column
Public Const RC_COUNT      As Long = 7

'--- Search options filled in by the UserForm --------------------------------
Public Type SearchOptions
    Cats              As Variant   'array of selected category names
    CatCount          As Long
    MatchAll          As Boolean   'True = AND, False = OR
    EntireMailbox     As Boolean
    IncludeSubfolders As Boolean
    IncludeShared     As Boolean
    OnlyAttachments   As Boolean
    UseDate           As Boolean
    DateFrom          As Date
    DateTo            As Date
    Sender            As String
    Keyword           As String
    SearchBody        As Boolean
End Type

'--- Shared state -------------------------------------------------------------
Public g_Results     As Variant   'final 2-D array (row, col), sorted newest first
Public g_ResultCount As Long
Public g_Status      As Object    'Label on the search form (progress text)

Private m_Buf()   As String       'growing buffer (col, row)
Private m_Cap     As Long
Private m_Count   As Long
Private m_Skip    As String       'EntryIDs of folders to skip (Deleted, Junk)
Private m_Failed  As Long
Private m_Scanned As Long       'diagnostics
Private m_WithCats As Long
Private m_Sample  As String
Private m_FirstErr As String
Private m_LastFolder As String   'remembered between saves
Private m_Step As String       'which statement was running (for error text)

'==============================================================================
' ENTRY POINT
'==============================================================================
Public Sub ShowCategorySearch()
    UserForms.Add("frmCategorySearch").Show vbModeless
End Sub

'==============================================================================
' MAIN SEARCH
'==============================================================================
Public Function ExecuteSearch(ByRef opt As SearchOptions) As Long
    Dim folders As Collection
    Dim fld As Outlook.Folder
    Dim dasl As String
    Dim idx As Long, total As Long

    On Error GoTo EH

    m_Cap = 1000: m_Count = 0: m_Failed = 0
    m_Scanned = 0: m_WithCats = 0: m_Sample = "": m_FirstErr = "": m_Step = ""
    ReDim m_Buf(0 To RC_COUNT - 1, 0 To m_Cap - 1)

    dasl = BuildDasl(opt)
    Set folders = GetFoldersToSearch(opt)
    total = folders.Count

    For Each fld In folders
        idx = idx + 1
        SetStatus "Folder " & idx & " of " & total & ": " & fld.Name & _
                  "   (" & m_Count & " found so far)"
        If Not SearchOneFolder(fld, opt, dasl) Then m_Failed = m_Failed + 1
    Next fld

    FinalizeResults

    SetStatus "Done: " & m_Count & " email(s) in " & total & " folder(s). Scanned " & _
              m_Scanned & " mails, " & m_WithCats & " had categories." & _
              IIf(m_Failed > 0, "  " & m_Failed & " folder(s) unreadable, first: " & m_FirstErr, "") & _
              IIf(m_Count = 0 And Len(m_Sample) > 0, "  Sample category text: [" & m_Sample & "]", "")
    ExecuteSearch = m_Count
    Exit Function
EH:
    MsgBox "Search failed: " & Err.Description, vbExclamation, "Category Search"
    ExecuteSearch = 0
End Function

Public Sub OpenResult(ByVal rowIdx As Long)
    Dim itm As Object
    On Error GoTo EH
    If rowIdx < 0 Or rowIdx >= g_ResultCount Then Exit Sub
    Set itm = Application.Session.GetItemFromID(g_Results(rowIdx, RC_ENTRYID), _
                                                g_Results(rowIdx, RC_STOREID))
    itm.Display
    Exit Sub
EH:
    MsgBox "Could not open this item (moved or deleted?): " & Err.Description, _
           vbExclamation, "Category Search"
End Sub

Public Sub ExportResultsToExcel()
    Dim xl As Object, wb As Object, ws As Object
    Dim data() As Variant, r As Long, s As String
    If g_ResultCount = 0 Then
        MsgBox "Nothing to export.", vbInformation, "Category Search"
        Exit Sub
    End If
    On Error GoTo EH
    ReDim data(1 To g_ResultCount, 1 To 5)
    For r = 0 To g_ResultCount - 1
        data(r + 1, 1) = SafeCell(CStr(g_Results(r, RC_SUBJECT)))
        data(r + 1, 2) = SafeCell(CStr(g_Results(r, RC_SENDER)))
        s = CStr(g_Results(r, RC_RECEIVED))
        If IsDate(s) Then data(r + 1, 3) = CDate(s) Else data(r + 1, 3) = s
        data(r + 1, 4) = SafeCell(CStr(g_Results(r, RC_CATEGORIES)))
        data(r + 1, 5) = SafeCell(CStr(g_Results(r, RC_FOLDER)))
    Next r
    Set xl = CreateObject("Excel.Application")
    xl.Visible = True
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Category Search"
    ws.Range("A1:E1").Value = Array("Subject", "Sender", "Received", "Categories", "Folder")
    ws.Range("A2").Resize(g_ResultCount, 5).Value = data
    ws.Range("A1:E1").Font.Bold = True
    ws.Columns("C").NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Columns("A:E").AutoFit
    ws.Range("A1:E1").AutoFilter
    Exit Sub
EH:
    MsgBox "Export failed: " & Err.Description, vbExclamation, "Category Search"
End Sub

Private Function SafeCell(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) Like "[=+@-]" Then s = "'" & s
    End If
    SafeCell = s
End Function

Public Sub SaveResultsToFolder()
    Dim target As String, fullPath As String, baseName As String
    Dim r As Long, itm As Object, saved As Long, failed As Long, n As Long
    Dim nameMax As Long, msg As String
    If g_ResultCount = 0 Then
        MsgBox "No results to save.", vbInformation, "Category Search"
        Exit Sub
    End If
    On Error GoTo EH
    target = PickFolder("Choose the folder to save " & g_ResultCount & " email(s) into (.msg files)")
    If Len(target) = 0 Then Exit Sub
    m_LastFolder = target
    If Right$(target, 1) <> "\" Then target = target & "\"
    nameMax = 240 - Len(target) - 30
    If nameMax < 15 Then nameMax = 15
    For r = 0 To g_ResultCount - 1
        On Error Resume Next
        Set itm = Nothing
        Set itm = Application.Session.GetItemFromID(g_Results(r, RC_ENTRYID), g_Results(r, RC_STOREID))
        If itm Is Nothing Then
            failed = failed + 1
        Else
            baseName = Replace(Left$(CStr(g_Results(r, RC_RECEIVED)), 16), ":", "") & " - " & _
                       CleanFileName(CStr(g_Results(r, RC_SUBJECT)), nameMax)
            fullPath = target & baseName & ".msg"
            n = 1
            Do While Len(Dir$(fullPath)) > 0
                n = n + 1
                fullPath = target & baseName & " (" & n & ").msg"
            Loop
            Err.Clear
            itm.SaveAs fullPath, olMSGUnicode
            If Err.Number = 0 Then saved = saved + 1 Else failed = failed + 1
        End If
        Err.Clear
        On Error GoTo EH
        If r Mod 25 = 0 Then SetStatus "Saving " & (r + 1) & " of " & g_ResultCount & "..."
    Next r
    msg = saved & " email(s) saved to:" & vbCrLf & target
    If failed > 0 Then msg = msg & vbCrLf & vbCrLf & failed & " could not be saved (moved/deleted?)."
    MsgBox msg, vbInformation, "Category Search"
    CreateObject("WScript.Shell").Run "explorer.exe """ & target & """"
    Exit Sub
EH:
    MsgBox "Save failed: " & Err.Description, vbExclamation, "Category Search"
End Sub

Private Function PickFolder(ByVal prompt As String) As String
    Dim xl As Object, fd As Object, shl As Object, fol As Object
    Dim shown As Boolean, rc As Long
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    If Not xl Is Nothing Then
        Set fd = xl.FileDialog(4)
        If Not fd Is Nothing Then
            fd.Title = prompt
            fd.AllowMultiSelect = False
            If Len(m_LastFolder) > 0 Then fd.InitialFileName = m_LastFolder & "\"
            Err.Clear
            rc = fd.Show
            If Err.Number = 0 Then
                shown = True
                If rc = -1 Then PickFolder = CStr(fd.SelectedItems(1))
            End If
        End If
        xl.Quit
        Set xl = Nothing
    End If
    Err.Clear
    If shown Then Exit Function
    Set shl = CreateObject("Shell.Application")
    Set fol = shl.BrowseForFolder(0, prompt, 65)
    If Not fol Is Nothing Then PickFolder = fol.Self.Path
    If Left$(PickFolder, 2) = "::" Then PickFolder = ""
End Function

Private Function CleanFileName(ByVal s As String, ByVal maxLen As Long) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr("\/:*?""<>|", ch) > 0 Or AscW(ch) < 32 Then
            out = out & "-"
        Else
            out = out & ch
        End If
    Next i
    out = Trim$(out)
    Do While Len(out) > 0 And (Right$(out, 1) = "." Or Right$(out, 1) = " ")
        out = Left$(out, Len(out) - 1)
    Loop
    If Len(out) > maxLen Then out = Trim$(Left$(out, maxLen))
    If Len(out) = 0 Then out = "(no subject)"
    CleanFileName = out
End Function

Private Function BuildDasl(ByRef opt As SearchOptions) As String
    Dim kw As String
    If Len(opt.Keyword) = 0 Or Not opt.SearchBody Then Exit Function
    kw = Esc(opt.Keyword)
    BuildDasl = "@SQL=(" & Q(P_SUBJECT) & " LIKE '%" & kw & "%' OR " & _
                Q(P_BODY) & " LIKE '%" & kw & "%')"
End Function

Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function

Private Function Esc(ByVal s As String) As String
    Esc = Replace(s, "'", "''")
End Function

Private Function CategoriesMatch(ByVal itemCats As String, ByRef opt As SearchOptions) As Boolean
    Dim parts() As String, setStr As String, i As Long, hits As Long
    If opt.CatCount = 0 Then CategoriesMatch = True: Exit Function
    If Len(itemCats) = 0 Then Exit Function
    parts = Split(Replace(itemCats, ";", ","), ",")
    setStr = "|"
    For i = LBound(parts) To UBound(parts)
        setStr = setStr & LCase$(Trim$(parts(i))) & "|"
    Next i
    For i = 0 To opt.CatCount - 1
        If InStr(1, setStr, "|" & LCase$(Trim$(CStr(opt.Cats(i)))) & "|", vbBinaryCompare) > 0 Then
            hits = hits + 1
        End If
    Next i
    If opt.MatchAll Then
        CategoriesMatch = (hits = opt.CatCount)
    Else
        CategoriesMatch = (hits > 0)
    End If
End Function

Private Function DatePasses(ByVal recv As Date, ByRef opt As SearchOptions) As Boolean
    If Not opt.UseDate Then DatePasses = True: Exit Function
    DatePasses = (recv >= opt.DateFrom And recv < opt.DateTo + 1)
End Function

Private Function GetFoldersToSearch(ByRef opt As SearchOptions) As Collection
    Dim col As New Collection
    Dim st As Outlook.Store
    Dim defId As String
    m_Skip = "|"
    On Error Resume Next
    defId = Application.Session.DefaultStore.StoreID
    On Error GoTo 0
    For Each st In Application.Session.Stores
        If st.StoreID = defId Or opt.IncludeShared Then
            AddStoreFolders col, st, opt
        End If
    Next st
    Set GetFoldersToSearch = col
End Function

Private Sub AddStoreFolders(col As Collection, st As Outlook.Store, ByRef opt As SearchOptions)
    Dim f As Outlook.Folder
    Set f = SafeDefaultFolder(st, olFolderDeletedItems)
    If Not f Is Nothing Then m_Skip = m_Skip & f.EntryID & "|"
    Set f = SafeDefaultFolder(st, olFolderJunk)
    If Not f Is Nothing Then m_Skip = m_Skip & f.EntryID & "|"
    On Error Resume Next
    If opt.EntireMailbox Then
        Set f = st.GetRootFolder
        If Not f Is Nothing Then AddFolderTree col, f, True
    Else
        Set f = SafeDefaultFolder(st, olFolderInbox)
        If Not f Is Nothing Then AddFolderTree col, f, opt.IncludeSubfolders
    End If
    On Error GoTo 0
End Sub

Private Function SafeDefaultFolder(st As Outlook.Store, ByVal fid As OlDefaultFolders) As Outlook.Folder
    On Error Resume Next
    Set SafeDefaultFolder = st.GetDefaultFolder(fid)
End Function

Private Sub AddFolderTree(col As Collection, fld As Outlook.Folder, ByVal recurse As Boolean)
    Dim subFld As Outlook.Folder
    If InStr(m_Skip, "|" & fld.EntryID & "|") > 0 Then Exit Sub
    If fld.DefaultItemType = olMailItem Then col.Add fld
    If recurse Then
        For Each subFld In fld.Folders
            AddFolderTree col, subFld, True
        Next subFld
    End If
End Sub

Private Sub AddResult(ByVal subj As String, ByVal sndr As String, ByVal recv As Date, _
                      ByVal cats As String, ByVal folderPath As String, _
                      ByVal entryId As String, ByVal storeId As String)
    If m_Count > m_Cap - 1 Then
        m_Cap = m_Cap * 2
        ReDim Preserve m_Buf(0 To RC_COUNT - 1, 0 To m_Cap - 1)
    End If
    m_Buf(RC_SUBJECT, m_Count) = subj
    m_Buf(RC_SENDER, m_Count) = sndr
    m_Buf(RC_RECEIVED, m_Count) = Format$(recv, "yyyy-mm-dd hh:nn")
    m_Buf(RC_CATEGORIES, m_Count) = cats
    m_Buf(RC_FOLDER, m_Count) = folderPath
    m_Buf(RC_ENTRYID, m_Count) = entryId
    m_Buf(RC_STOREID, m_Count) = storeId
    m_Count = m_Count + 1
End Sub

Private Sub FinalizeResults()
    Dim res() As Variant
    Dim r As Long, c As Long
    g_ResultCount = m_Count
    If m_Count = 0 Then g_Results = Empty: Exit Sub
    ReDim res(0 To m_Count - 1, 0 To RC_COUNT - 1)
    For r = 0 To m_Count - 1
        For c = 0 To RC_COUNT - 1
            res(r, c) = m_Buf(c, r)
        Next c
    Next r
    Erase m_Buf
    If m_Count > 1 Then QuickSortRows res, 0, m_Count - 1, RC_RECEIVED
    g_Results = res
End Sub

Private Sub QuickSortRows(ByRef a As Variant, ByVal lo As Long, ByVal hi As Long, ByVal keyCol As Long)
    Dim i As Long, j As Long, pivot As String
    i = lo: j = hi
    pivot = a((lo + hi) \ 2, keyCol)
    Do While i <= j
        Do While a(i, keyCol) > pivot: i = i + 1: Loop
        Do While a(j, keyCol) < pivot: j = j - 1: Loop
        If i <= j Then
            SwapRows a, i, j
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortRows a, lo, j, keyCol
    If i < hi Then QuickSortRows a, i, hi, keyCol
End Sub

Private Sub SwapRows(ByRef a As Variant, ByVal r1 As Long, ByVal r2 As Long)
    Dim c As Long, t As Variant
    For c = 0 To RC_COUNT - 1
        t = a(r1, c): a(r1, c) = a(r2, c): a(r2, c) = t
    Next c
End Sub

Private Function SearchOneFolder(fld As Outlook.Folder, ByRef opt As SearchOptions, _
                                 ByVal dasl As String) As Boolean
    Dim tbl As Outlook.Table
    Dim rw As Outlook.Row
    Dim cFrom As String, cMail As String, cRecv As String, cCats As String, cAtt As String
    Dim cats As String, recv As Date, v As Variant
    Dim subj As String, sndr As String, mail As String
    Dim path As String, storeId As String
    On Error GoTo EH
    If Len(dasl) > 0 Then
        Set tbl = fld.GetTable(dasl)
    Else
        Set tbl = fld.GetTable
    End If
    tbl.Columns.RemoveAll
    tbl.Columns.Add "EntryID"
    tbl.Columns.Add "Subject"
    cFrom = AddCol(tbl, P_FROMNAME, "SenderName")
    cRecv = AddCol(tbl, P_RECEIVED, "ReceivedTime")
    cCats = AddCol(tbl, "Categories", P_KEYWORDS)
    If Len(opt.Sender) > 0 Then cMail = AddCol(tbl, P_FROMEMAIL, "SenderEmailAddress")
    If opt.OnlyAttachments Then cAtt = AddCol(tbl, P_HASATTACH, "HasAttachment")
    If Len(cCats) = 0 Then Err.Raise vbObjectError + 1, , "Categories column not available"
    path = fld.FolderPath
    storeId = fld.StoreID
    Do Until tbl.EndOfTable
        Set rw = tbl.GetNextRow
        m_Scanned = m_Scanned + 1
        cats = NzStr(rw.Item(cCats))
        If Len(cats) > 0 Then
            m_WithCats = m_WithCats + 1
            If Len(m_Sample) = 0 Then m_Sample = cats
        End If
        If CategoriesMatch(cats, opt) Then
            subj = NzStr(rw.Item("Subject"))
            If Len(opt.Keyword) > 0 And Not opt.SearchBody Then
                If InStr(1, subj, opt.Keyword, vbTextCompare) = 0 Then GoTo NextRow
            End If
            If opt.OnlyAttachments Then
                If Not IsTrue(rw.Item(cAtt)) Then GoTo NextRow
            End If
            If Len(cFrom) > 0 Then sndr = NzStr(rw.Item(cFrom)) Else sndr = ""
            If Len(opt.Sender) > 0 Then
                If Len(cMail) > 0 Then mail = NzStr(rw.Item(cMail)) Else mail = ""
                If InStr(1, sndr, opt.Sender, vbTextCompare) = 0 And _
                   InStr(1, mail, opt.Sender, vbTextCompare) = 0 Then GoTo NextRow
            End If
            recv = 0
            If Len(cRecv) > 0 Then
                v = rw.Item(cRecv)
                If IsDate(v) Then recv = ToLocal(rw, v)
            End If
            If Not DatePasses(recv, opt) Then GoTo NextRow
            AddResult subj, sndr, recv, cats, path, NzStr(rw.Item("EntryID")), storeId
        End If
NextRow:
    Loop
    SearchOneFolder = True
    Exit Function
EH:
    If Len(m_FirstErr) = 0 Then m_FirstErr = fld.FolderPath & " " & Err.Description
    SearchOneFolder = False
End Function

Private Function AddCol(tbl As Outlook.Table, ByVal primary As String, ByVal alt As String) As String
    On Error Resume Next
    tbl.Columns.Add primary
    If Err.Number = 0 Then AddCol = primary: Exit Function
    Err.Clear
    If Len(alt) > 0 Then
        tbl.Columns.Add alt
        If Err.Number = 0 Then AddCol = alt: Exit Function
        Err.Clear
    End If
    AddCol = ""
End Function

Private Function ToLocal(rw As Outlook.Row, ByVal v As Variant) As Date
    Dim d As Date
    d = CDate(v)
    On Error Resume Next
    ToLocal = rw.UTCToLocal(d)
    If Err.Number <> 0 Or ToLocal = 0 Then
        Err.Clear
        ToLocal = DateAdd("n", -TzBiasMinutes(), d)
    End If
End Function

Private Function TzBiasMinutes() As Long
    Static done As Boolean, bias As Long
    Dim tzi As TIME_ZONE_INFORMATION, r As Long
    If Not done Then
        r = GetTimeZoneInformation(tzi)
        If r = 2 Then bias = tzi.Bias + tzi.DaylightBias Else bias = tzi.Bias + tzi.StandardBias
        done = True
    End If
    TzBiasMinutes = bias
End Function

Private Function IsTrue(v As Variant) As Boolean
    On Error Resume Next
    IsTrue = CBool(v)
End Function

Private Function NzStr(v As Variant) As String
    If IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    ElseIf IsArray(v) Then
        NzStr = Join(v, "; ")
    Else
        NzStr = CStr(v)
    End If
End Function

Public Sub SetStatus(ByVal msg As String)
    On Error Resume Next
    If Not g_Status Is Nothing Then g_Status.Caption = msg
    DoEvents
End Sub
