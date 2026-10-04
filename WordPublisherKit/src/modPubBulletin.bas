Attribute VB_Name = "modPubBulletin"
Option Explicit
'==============================================================================
' Word Publisher Kit - church bulletin tools ("Bulletin" ribbon tab)
'
' A bulletin is a folded booklet: 7 x 8.5" pages on landscape legal paper
' (or 5.5 x 8.5" on letter). Word's "Book fold" page setup handles the page
' imposition, so volunteers just type pages 1, 2, 3... in order and Word prints
' them in the right order for folding and stapling.
'
' The order of worship is ordinary flowing text formatted with the bulletin
' styles below, so items can be added or removed and everything reflows.
'==============================================================================

' ---- Church defaults (edit before sending the add-in out, or use Settings) ---
Private Const DEF_CHURCH_NAME As String = "Park Road Presbyterian Church"
Private Const DEF_SHORT_NAME As String = "Park Road"
Private Const DEF_ADDRESS As String = "[Street address]  |  Hollywood, Florida"
Private Const DEF_WEBSITE As String = "parkroadpres.org"
Private Const DEF_SERVICES As String = "Lord's Day Worship  |  9:00 & 11:15 AM"
Private Const DEF_COPYRIGHT1 As String = "Scripture taken from the NEW AMERICAN STANDARD BIBLE"
Private Const DEF_COPYRIGHT2 As String = "Hymns and Praise Songs printed in our bulletin are used by permission through our Church Copyright License #337734"

Private Const SETTINGS_APP As String = "WordPublisherKit"
Private Const SETTINGS_SECTION As String = "Church"
Private Const BULLETIN_FONT As String = "Georgia"
Private Const DATE_TAG As String = "BulletinDate"
Private Const DATE_FORMAT As String = "mmmm d, yyyy"

'------------------------------------------------------------------------------
' Church settings (stored per Windows user)
'------------------------------------------------------------------------------
Public Function ChurchSetting(ByVal key As String) As String
    Dim defaultValue As String
    Select Case key
        Case "Name": defaultValue = DEF_CHURCH_NAME
        Case "ShortName": defaultValue = DEF_SHORT_NAME
        Case "Address": defaultValue = DEF_ADDRESS
        Case "Website": defaultValue = DEF_WEBSITE
        Case "Services": defaultValue = DEF_SERVICES
        Case "Copyright1": defaultValue = DEF_COPYRIGHT1
        Case "Copyright2": defaultValue = DEF_COPYRIGHT2
    End Select
    ChurchSetting = GetSetting(SETTINGS_APP, SETTINGS_SECTION, key, defaultValue)
End Function

Public Sub EditChurchSettings()
    AskSetting "Name", "Church name (cover):"
    AskSetting "ShortName", "Short name (e.g. ""THIS WEEK at Park Road""):"
    AskSetting "Address", "Address line (cover):"
    AskSetting "Website", "Website (cover):"
    AskSetting "Services", "Service times line (cover):"
    AskSetting "Copyright1", "Copyright line 1 (Bible translation):"
    AskSetting "Copyright2", "Copyright line 2 (music license):"
    MsgBox "Saved. New bulletins will use these settings.", vbInformation, "Church Settings"
End Sub

Private Sub AskSetting(ByVal key As String, ByVal prompt As String)
    Dim cur As String, ans As String
    cur = ChurchSetting(key)
    ans = InputBox(prompt & vbCr & vbCr & "(Cancel keeps the current value.)", "Church Settings", cur)
    If Len(ans) > 0 And ans <> cur Then SaveSetting SETTINGS_APP, SETTINGS_SECTION, key, ans
End Sub

'------------------------------------------------------------------------------
' Styles
'------------------------------------------------------------------------------
Public Sub EnsureBulletinStyles(doc As Document)
    Dim tw As Single, st As Style

    tw = TextWidth(doc)
    DefStyle doc, "Cover Title", 24, True, False, wdAlignParagraphCenter, 0, 36, 6
    DefStyle doc, "Cover Subtitle", 14, False, True, wdAlignParagraphCenter, 0, 0, 4
    DefStyle doc, "Cover Details", 11, False, False, wdAlignParagraphCenter, 0, 0, 2
    DefStyle doc, "Bulletin Date", 10, False, False, wdAlignParagraphLeft, 0, 0, 8

    Set st = DefStyle(doc, "Heading Bar", 11, False, False, wdAlignParagraphCenter, 0, 6, 8, True)
    If Not st Is Nothing Then
        st.Font.Color = wdColorWhite
        st.Font.Spacing = 1.5
        st.ParagraphFormat.Shading.BackgroundPatternColor = wdColorBlack
        st.NextParagraphStyle = doc.Styles(wdStyleNormal).NameLocal
    End If

    DefStyle doc, "Section Title", 12, True, False, wdAlignParagraphCenter, 0, 10, 6, True
    AddRightTab DefStyle(doc, "Worship Item", 11.5, True, False, wdAlignParagraphLeft, 0, 10, 2, True), tw
    AddRightTab DefStyle(doc, "Worship Detail", 10.5, False, False, wdAlignParagraphLeft, 0.25, 0, 2), tw
    AddRightTab DefStyle(doc, "Hymn Line", 10.5, False, False, wdAlignParagraphLeft, 0.25, 0, 4, True), tw
    DefStyle doc, "Lyrics", 10, False, False, wdAlignParagraphLeft, 0.5, 0, 0
    DefStyle doc, "Chorus", 10, False, True, wdAlignParagraphLeft, 0.75, 0, 0
    DefStyle doc, "Scripture Text", 10, False, True, wdAlignParagraphJustify, 0.25, 0, 4
    DefStyle doc, "Leader Line", 10.5, False, False, wdAlignParagraphLeft, 0.25, 0, 2
    DefStyle doc, "People Line", 10.5, True, False, wdAlignParagraphLeft, 0.25, 0, 2
    DefStyle doc, "Reflection Text", 10.5, False, False, wdAlignParagraphLeft, 0, 0, 6
    DefStyle doc, "Reflection Credit", 9.5, False, False, wdAlignParagraphRight, 0, 0, 14
    DefStyle doc, "Announcement Title", 10.5, True, False, wdAlignParagraphLeft, 0, 8, 1, True
    DefStyle doc, "Announcement Text", 10, False, False, wdAlignParagraphLeft, 0, 0, 2
    DefStyle doc, "Small Print", 7.5, False, False, wdAlignParagraphCenter, 0, 0, 0
End Sub

' Creates the style if it does not exist yet. Existing styles are left alone so
' a church's own tweaks survive. Returns Nothing when the style already existed.
Private Function DefStyle(doc As Document, ByVal styleName As String, ByVal fontSize As Single, _
        ByVal isBold As Boolean, ByVal isItalic As Boolean, ByVal align As Long, _
        ByVal indentIn As Single, ByVal spaceBefore As Single, ByVal spaceAfter As Single, _
        Optional ByVal keepNext As Boolean = False) As Style
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(styleName)
    On Error GoTo 0
    If Not st Is Nothing Then Exit Function

    Set st = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeParagraph)
    st.BaseStyle = doc.Styles(wdStyleNormal).NameLocal
    st.QuickStyle = True
    With st.Font
        .Name = BULLETIN_FONT
        .Size = fontSize
        .Bold = isBold
        .Italic = isItalic
    End With
    With st.ParagraphFormat
        .Alignment = align
        .LeftIndent = InchesToPoints(indentIn)
        .FirstLineIndent = 0
        .SpaceBefore = spaceBefore
        .SpaceAfter = spaceAfter
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = keepNext
    End With
    Set DefStyle = st
End Function

Private Sub AddRightTab(st As Style, ByVal pos As Single)
    If st Is Nothing Then Exit Sub
    st.ParagraphFormat.TabStops.ClearAll
    st.ParagraphFormat.TabStops.Add Position:=pos, Alignment:=wdAlignTabRight
End Sub

Private Function TextWidth(doc As Document) As Single
    Dim w As Single, h As Single
    GetPageSize w, h, doc
    With doc.Sections(1).PageSetup
        TextWidth = w - .LeftMargin - .RightMargin - .Gutter
    End With
End Function

Public Sub ApplyBulletinStyle(ByVal styleName As String)
    RequireDoc
    EnsureBulletinStyles ActiveDocument
    Selection.Style = ActiveDocument.Styles(styleName)
End Sub

'------------------------------------------------------------------------------
' New bulletin
'------------------------------------------------------------------------------
Public Sub NewBulletin(ByVal paper As String)
    Dim doc As Document, r As Range, w As Single, h As Single, shp As Shape
    Dim sunday As String

    sunday = Format$(NextSunday(Date), DATE_FORMAT)
    Set doc = Documents.Add
    With doc.PageSetup
        If paper = "Letter" Then
            .PageWidth = InchesToPoints(11)
        Else
            .PageWidth = InchesToPoints(14)
        End If
        .PageHeight = InchesToPoints(8.5)
        .BookFoldPrinting = True
        .TopMargin = InchesToPoints(0.5)
        .BottomMargin = InchesToPoints(0.5)
        .LeftMargin = InchesToPoints(0.5)
        .RightMargin = InchesToPoints(0.5)
        .Gutter = 0
        .HeaderDistance = InchesToPoints(0.3)
        .FooterDistance = InchesToPoints(0.25)
    End With
    With doc.Styles(wdStyleNormal)
        .Font.Name = BULLETIN_FONT
        .Font.Size = 10.5
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    End With
    EnsureBulletinStyles doc

    ' ---- Cover -------------------------------------------------------------
    AppendPara doc, ChurchSetting("Name"), "Cover Title"
    AppendPara doc, "The Lord's Day Worship", "Cover Subtitle"
    TagDate AppendPara(doc, sunday, "Cover Subtitle")
    AppendPara doc, ChurchSetting("Services"), "Cover Details"
    AppendPara doc, ChurchSetting("Address"), "Cover Details"
    AppendPara doc, ChurchSetting("Website"), "Cover Details"
    AppendPara doc, Chr$(12), wdStyleNormal

    ' ---- Inside cover: reflection -----------------------------------------
    TagDate AppendPara(doc, sunday, "Bulletin Date")
    AppendPara doc, "A Reflection Before the Service", "Heading Bar"
    AppendPara doc, "[Type or paste a reflection, poem or quotation here.]", "Reflection Text"
    AppendPara doc, ChrW$(8212) & " [Author, source]", "Reflection Credit"
    AppendPara doc, ChurchSetting("Copyright1"), "Small Print"
    AppendPara doc, ChurchSetting("Copyright2"), "Small Print"
    AppendPara doc, Chr$(12), wdStyleNormal

    ' ---- Order of worship -------------------------------------------------
    AppendPara doc, "The Prelude", "Worship Item"
    AppendPara doc, "The Welcome and Announcements", "Worship Item"
    AppendPara doc, "The Call to Worship: [Psalm]", "Worship Item"
    AppendPara doc, "Leader: [Leader's line]", "Leader Line"
    AppendPara doc, "People: [People's response]", "People Line"
    AppendPara doc, "The Hymn of Praise", "Worship Item"
    AppendPara doc, ChrW$(8220) & "[Hymn title]" & ChrW$(8221) & " (vv. 1-4)" & vbTab & "#[000]", "Hymn Line"
    AppendPara doc, "The Confession of Sin", "Worship Item"
    AppendPara doc, "The Assurance of Pardon", "Worship Item"
    AppendPara doc, "The Scripture Reading: [Reference]", "Worship Item"
    AppendPara doc, "[Paste the Scripture passage here.]", "Scripture Text"
    AppendPara doc, "Reader: The Word of the Lord", "Leader Line"
    AppendPara doc, "People: Thanks be to God!", "People Line"
    AppendPara doc, "The Sermon" & vbTab & ChrW$(8212) & " [Preacher]", "Worship Item"
    AppendPara doc, "[Sermon title]", "Worship Detail"
    AppendPara doc, "The Hymn of Response", "Worship Item"
    AppendPara doc, ChrW$(8220) & "[Hymn title]" & ChrW$(8221) & vbTab & "#[000]", "Hymn Line"
    AppendPara doc, "The Benediction", "Worship Item"
    AppendPara doc, "The Postlude", "Worship Item"
    AppendPara doc, Chr$(12), wdStyleNormal

    ' ---- Announcements ----------------------------------------------------
    AppendPara doc, "THIS WEEK at " & ChurchSetting("ShortName"), "Heading Bar"
    AppendPara doc, "Welcome! We are glad to have you with us today.", "Announcement Text"
    AppendScheduleTable doc
    AppendPara doc, "UPCOMING THIS MONTH at " & ChurchSetting("ShortName"), "Heading Bar"
    AppendPara doc, "[Event name] | [Day, Date] | [Time] | [Place]", "Announcement Title"
    AppendPara doc, "[Details about the event.]", "Announcement Text"
    AppendPara doc, "[Event name] | [Day, Date] | [Time] | [Place]", "Announcement Title"
    AppendPara doc, "[Details about the event.]", "Announcement Text"

    ' Page numbers (none on the cover)
    doc.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add _
        PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=False

    ' Cover picture: a placeholder that text flows above and below.
    GetPageSize w, h, doc
    Set shp = AddPicturePlaceholderOnPage(1, (w - InchesToPoints(4.5)) / 2, InchesToPoints(2.6), _
        InchesToPoints(4.5), InchesToPoints(3))
    shp.WrapFormat.Type = PUB_WRAP_TOPBOTTOM

    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.View.ShowAll = False
    Selection.HomeKey wdStory
    ActiveWindow.View.Zoom.PageFit = wdPageFitFullPage
    Application.StatusBar = "New bulletin created. Replace the [bracketed] text, then use Print Booklet."
End Sub

' Adds a paragraph at the end of the document and returns the range of its text.
Private Function AppendPara(doc As Document, ByVal txt As String, ByVal styleRef As Variant) As Range
    Dim r As Range, txtRange As Range
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertAfter txt
    If VarType(styleRef) = vbString Then
        r.Style = doc.Styles(styleRef)
    Else
        r.Style = styleRef
    End If
    Set txtRange = r.Duplicate
    r.InsertParagraphAfter
    Set AppendPara = txtRange
End Function

Private Sub AppendScheduleTable(doc As Document)
    Dim r As Range, t As Table, rowData As Variant, i As Long, tw As Single
    rowData = Array( _
        Array("Sunday", "9:00, 11:15 AM", "Worship Service"), _
        Array("", "10:10 AM", "Sunday School"), _
        Array("[Day]", "[Time]", "[Event] ([Room])"), _
        Array("[Day]", "[Time]", "[Event] ([Room])"))
    tw = TextWidth(doc)
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set t = doc.Tables.Add(Range:=r, NumRows:=UBound(rowData) + 1, NumColumns:=3)
    t.Borders.Enable = False
    t.Range.Style = doc.Styles("Announcement Text")
    t.Columns(1).Width = tw * 0.2
    t.Columns(2).Width = tw * 0.27
    t.Columns(3).Width = tw * 0.53
    For i = 0 To UBound(rowData)
        t.Cell(i + 1, 1).Range.Text = rowData(i)(0)
        t.Cell(i + 1, 2).Range.Text = rowData(i)(1)
        t.Cell(i + 1, 3).Range.Text = rowData(i)(2)
    Next i
    For i = 1 To t.Rows.Count
        t.Cell(i, 1).Range.Font.Bold = True
    Next i
End Sub

'------------------------------------------------------------------------------
' Dates
'------------------------------------------------------------------------------
Private Function NextSunday(ByVal d As Date) As Date
    NextSunday = d + (8 - Weekday(d, vbSunday)) Mod 7
End Function

Private Sub TagDate(r As Range)
    Dim cc As ContentControl
    Set cc = r.Document.ContentControls.Add(wdContentControlText, r)
    cc.Tag = DATE_TAG
    cc.Title = "Bulletin date"
End Sub

Private Function SetDateControls(ByVal txt As String) As Long
    Dim cc As ContentControl, n As Long
    For Each cc In ActiveDocument.ContentControls
        If cc.Tag = DATE_TAG Then
            cc.Range.Text = txt
            n = n + 1
        End If
    Next cc
    SetDateControls = n
End Function

Private Function AskSundayDate(ByVal title As String) As Date
    Dim ans As String
    ans = InputBox("Date of the service:", title, Format$(NextSunday(Date), DATE_FORMAT))
    If Len(ans) = 0 Then Exit Function
    If Not IsDate(ans) Then PubFail "'" & ans & "' is not a date. Try something like October 11, 2026."
    AskSundayDate = CDate(ans)
End Function

Public Sub SetBulletinDate()
    Dim d As Date
    RequireDoc
    d = AskSundayDate("Set Bulletin Date")
    If d = 0 Then Exit Sub
    If SetDateControls(Format$(d, DATE_FORMAT)) = 0 Then
        PubFail "This document has no bulletin date boxes. (They are created by New Bulletin.)"
    End If
End Sub

' Copies the open bulletin to a new file for the next service and updates the date.
Public Sub StartNextWeek()
    Dim d As Date, folder As String, target As String
    RequireDoc
    d = AskSundayDate("Start Next Week's Bulletin")
    If d = 0 Then Exit Sub
    folder = ActiveDocument.Path
    If Len(folder) = 0 Then folder = Options.DefaultFilePath(wdDocumentsPath)
    target = folder & Application.PathSeparator & "Bulletin " & Format$(d, "yyyy-mm-dd") & ".docx"
    If Len(Dir$(target)) > 0 Then
        If MsgBox(target & " already exists. Replace it?", vbYesNo + vbExclamation, "Next Week") = vbNo Then Exit Sub
    End If
    SetDateControls Format$(d, DATE_FORMAT)
    ActiveDocument.SaveAs2 FileName:=target, FileFormat:=wdFormatXMLDocument
    MsgBox "Saved as:" & vbCr & target & vbCr & vbCr & _
        "Last week's file was not changed. Now update the hymns, readings and announcements.", _
        vbInformation, "Next Week"
End Sub

'------------------------------------------------------------------------------
' Insert worship items at the cursor
'------------------------------------------------------------------------------
Private Sub BeginBlock()
    Dim r As Range
    RequireDoc
    If Selection.StoryType <> wdMainTextStory Then
        PubFail "Click in the bulletin text where the new item should go (not inside a text box or the master page)."
    End If
    If Selection.Information(wdWithInTable) Then PubFail "Click below the table first."
    EnsureBulletinStyles ActiveDocument
    Selection.Collapse wdCollapseEnd
    Set r = Selection.Paragraphs(1).Range
    If Len(r.Text) > 1 Then
        ' The current paragraph has text: start a new paragraph after it.
        Selection.SetRange r.End - 1, r.End - 1
        Selection.TypeParagraph
    End If
End Sub

Private Sub PutLine(ByVal txt As String, ByVal styleName As String, Optional ByVal isLast As Boolean = False)
    Selection.Style = ActiveDocument.Styles(styleName)
    Selection.TypeText txt
    If Not isLast Then Selection.TypeParagraph
End Sub

Private Function Quoted(ByVal s As String) As String
    Quoted = ChrW$(8220) & s & ChrW$(8221)
End Function

Public Sub InsertWorshipItem()
    Dim item As String, extra As String
    item = InputBox("Worship item, e.g. The Prelude, The Confession of Sin, The Offering:", _
        "Insert Worship Item", "The Prelude")
    If Len(item) = 0 Then Exit Sub
    extra = InputBox("Optional text at the right edge (a reference, a name...). Leave blank for none:", _
        "Insert Worship Item")
    BeginBlock
    If Len(extra) > 0 Then item = item & vbTab & extra
    PutLine item, "Worship Item", True
End Sub

Public Sub InsertHymn()
    Dim label As String, title As String, verses As String, num As String, ln As String
    Dim wantLyrics As Boolean
    label = InputBox("Heading for this song:", "Insert Hymn", "The Hymn of Praise")
    If Len(label) = 0 Then Exit Sub
    title = InputBox("Hymn or song title:", "Insert Hymn")
    If Len(title) = 0 Then Exit Sub
    verses = InputBox("Verses to sing (optional), e.g. vv. 1-3, 5:", "Insert Hymn")
    num = InputBox("Hymnal number, or the author for songs not in the hymnal (optional):", "Insert Hymn")
    wantLyrics = (MsgBox("Add space to paste the words (for songs not in the hymnal)?", _
        vbYesNo + vbQuestion, "Insert Hymn") = vbYes)

    ln = Quoted(title)
    If Len(verses) > 0 Then ln = ln & " (" & verses & ")"
    If Len(num) > 0 Then
        If IsNumeric(num) Then num = "#" & num
        ln = ln & vbTab & num
    End If
    BeginBlock
    PutLine label, "Worship Item"
    If wantLyrics Then
        PutLine ln, "Hymn Line"
        PutLine "[Paste verse 1 here]", "Lyrics"
        PutLine "[Paste the chorus here]", "Chorus", True
    Else
        PutLine ln, "Hymn Line", True
    End If
End Sub

Public Sub InsertScripture()
    Dim label As String, ref As String
    label = InputBox("Heading:", "Insert Scripture Reading", "The Scripture Reading")
    If Len(label) = 0 Then Exit Sub
    ref = InputBox("Passage, e.g. Nehemiah 2:11-3:5:", "Insert Scripture Reading")
    BeginBlock
    If Len(ref) > 0 Then label = label & ": " & ref
    PutLine label, "Worship Item"
    PutLine "[Paste the Scripture passage here.]", "Scripture Text"
    PutLine "Reader: The Word of the Lord", "Leader Line"
    PutLine "People: Thanks be to God!", "People Line", True
End Sub

Public Sub InsertResponsiveReading()
    Dim label As String, ref As String, pairs As String, i As Long, n As Long
    label = InputBox("Heading:", "Insert Responsive Reading", "The Call to Worship")
    If Len(label) = 0 Then Exit Sub
    ref = InputBox("Passage (optional), e.g. Psalm 150:", "Insert Responsive Reading")
    pairs = InputBox("How many Leader / People pairs?", "Insert Responsive Reading", "3")
    n = Val(pairs)
    If n < 1 Then n = 1
    If n > 20 Then n = 20
    BeginBlock
    If Len(ref) > 0 Then label = label & ": " & ref
    PutLine label, "Worship Item"
    For i = 1 To n
        PutLine "Leader: [line]", "Leader Line"
        PutLine "People: [response]", "People Line", (i = n)
    Next i
End Sub

Public Sub InsertSermon()
    Dim title As String, preacher As String, ref As String
    title = InputBox("Sermon title:", "Insert Sermon")
    If Len(title) = 0 Then Exit Sub
    preacher = InputBox("Preacher (optional):", "Insert Sermon")
    ref = InputBox("Sermon text (optional), e.g. Nehemiah 2:11-3:5:", "Insert Sermon")
    BeginBlock
    If Len(preacher) > 0 Then
        PutLine "The Sermon" & vbTab & ChrW$(8212) & " " & preacher, "Worship Item"
    Else
        PutLine "The Sermon", "Worship Item"
    End If
    If Len(ref) > 0 Then
        PutLine ref & "  |  " & title, "Worship Detail", True
    Else
        PutLine title, "Worship Detail", True
    End If
End Sub

Public Sub InsertHeadingBar()
    Dim txt As String
    txt = InputBox("Heading text:", "Insert Heading Bar", "A Reflection Before the Service")
    If Len(txt) = 0 Then Exit Sub
    BeginBlock
    PutLine txt, "Heading Bar", True
End Sub

Public Sub InsertAnnouncement()
    Dim title As String, whenWhere As String
    title = InputBox("Event name:", "Insert Announcement")
    If Len(title) = 0 Then Exit Sub
    whenWhere = InputBox("When and where (optional), e.g. Wed., Oct. 21 | 11:30 AM | Fellowship Hall:", _
        "Insert Announcement")
    BeginBlock
    If Len(whenWhere) > 0 Then title = title & " | " & whenWhere
    PutLine title, "Announcement Title"
    PutLine "[Details about the event.]", "Announcement Text", True
End Sub

Public Sub InsertScheduleTable()
    Dim t As Table, tw As Single, i As Long
    BeginBlock
    tw = TextWidth(ActiveDocument)
    Set t = ActiveDocument.Tables.Add(Range:=Selection.Range, NumRows:=4, NumColumns:=3)
    t.Borders.Enable = False
    t.Range.Style = ActiveDocument.Styles("Announcement Text")
    t.Columns(1).Width = tw * 0.2
    t.Columns(2).Width = tw * 0.27
    t.Columns(3).Width = tw * 0.53
    t.Cell(1, 1).Range.Text = "[Day]"
    t.Cell(1, 2).Range.Text = "[Time]"
    t.Cell(1, 3).Range.Text = "[Event] ([Room])"
    For i = 1 To t.Rows.Count
        t.Cell(i, 1).Range.Font.Bold = True
    Next i
    t.Cell(1, 1).Range.Select
End Sub

Public Sub InsertBulletinPageBreak()
    RequireDoc
    If Selection.StoryType <> wdMainTextStory Then PubFail "Click in the bulletin text first."
    Selection.Collapse wdCollapseEnd
    Selection.InsertBreak wdPageBreak
End Sub

'------------------------------------------------------------------------------
' Printing
'------------------------------------------------------------------------------
Public Sub PrintBooklet()
    RequireDoc
    If Not ActiveDocument.Sections(1).PageSetup.BookFoldPrinting Then
        If MsgBox("This document is not set up as a folded booklet. Print it as normal pages?", _
                  vbYesNo + vbQuestion, "Print Booklet") = vbNo Then Exit Sub
    Else
        MsgBox "In the Print window choose:" & vbCr & _
            "   " & ChrW$(8226) & " Print on Both Sides " & ChrW$(8212) & " flip pages on SHORT edge" & vbCr & _
            "   " & ChrW$(8226) & " Paper: " & PaperName() & ", Landscape" & vbCr & vbCr & _
            "Word puts the pages in booklet order for you. Fold the stack in half and staple.", _
            vbInformation, "Print Booklet"
    End If
    PrintPublication
End Sub

' Printer-ready PDF with the pages already arranged for folding (what a print
' shop or a copier needs). Uses the "Microsoft Print to PDF" printer.
Public Sub ExportBookletPdf()
    Dim target As String, oldPrinter As String, errMsg As String, was As Boolean
    RequireDoc
    target = AskPdfPath(" (booklet)")
    If Len(target) = 0 Then Exit Sub

    oldPrinter = Application.ActivePrinter
    was = GuidesVisible()
    SetGuidesVisible False
    On Error Resume Next
    Application.ActivePrinter = "Microsoft Print to PDF"
    If Err.Number = 0 Then
        ActiveDocument.PrintOut Background:=False, Copies:=1, OutputFileName:=target, PrintToFile:=True
    End If
    If Err.Number <> 0 Then errMsg = Err.Description
    Application.ActivePrinter = oldPrinter
    On Error GoTo 0
    SetGuidesVisible was

    If Len(errMsg) > 0 Then
        MsgBox "Word couldn't print to PDF automatically (" & errMsg & ")." & vbCr & vbCr & _
            "Use Print Booklet instead and pick ""Microsoft Print to PDF"" as the printer.", _
            vbExclamation, "Booklet PDF"
    Else
        MsgBox "Saved:" & vbCr & target, vbInformation, "Booklet PDF"
    End If
End Sub

Private Function AskPdfPath(ByVal suffix As String) As String
    Dim baseName As String, target As String
    baseName = ActiveDocument.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Save PDF"
        If Len(ActiveDocument.Path) > 0 Then
            .InitialFileName = ActiveDocument.Path & Application.PathSeparator & baseName & suffix & ".pdf"
        Else
            .InitialFileName = baseName & suffix & ".pdf"
        End If
        If .Show <> -1 Then Exit Function
        target = .SelectedItems(1)
    End With
    If InStrRev(target, ".") > InStrRev(target, Application.PathSeparator) Then
        target = Left$(target, InStrRev(target, ".") - 1)
    End If
    AskPdfPath = target & ".pdf"
End Function

Private Function PaperName() As String
    Dim longSide As Single
    With ActiveDocument.Sections(1).PageSetup
        longSide = .PageWidth
        If .PageHeight > longSide Then longSide = .PageHeight
        ' Word may report the folded page (7 x 8.5) instead of the sheet.
        If PointsToInches(longSide) < 9 Then
            longSide = 2 * IIf(.PageWidth < .PageHeight, .PageWidth, .PageHeight)
        End If
    End With
    If Abs(PointsToInches(longSide) - 14) < 0.2 Then
        PaperName = "Legal (8.5 x 14)"
    ElseIf Abs(PointsToInches(longSide) - 11) < 0.2 Then
        PaperName = "Letter (8.5 x 11)"
    Else
        PaperName = "the paper size in Page Setup"
    End If
End Function
