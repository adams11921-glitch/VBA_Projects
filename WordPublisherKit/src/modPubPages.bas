Attribute VB_Name = "modPubPages"
Option Explicit
'==============================================================================
' Word Publisher Kit - publications, pages, master page, guides, output
'
' How a "publication" is laid out in Word:
'   * Every page is ended by a manual page break, so each page starts with its
'     own paragraph. Objects are anchored to that paragraph and positioned
'     relative to the page, so they stay on "their" page like in Publisher.
'   * The header of section 1 acts as the master page: anything placed there
'     repeats on every page. Layout guides are drawn there too.
'==============================================================================

Private Const GUIDE_MARGIN_RGB As Long = 13107455   ' RGB(255, 0, 200) magenta
Private Const GUIDE_COLUMN_RGB As Long = 16752640   ' RGB(0, 160, 255) blue

'------------------------------------------------------------------------------
' New publication
'------------------------------------------------------------------------------
Public Sub NewPublication(ByVal kind As String)
    Dim w As Single, h As Single, m As Single, g As Single
    Dim cols As Long, rows As Long, pages As Long, i As Long
    Dim doc As Document, ans As String, parts() As String

    cols = 1: rows = 1: pages = 1: m = 0.5: g = 0.25
    Select Case kind
        Case "Letter":       w = 8.5: h = 11
        Case "A4":           w = 8.27: h = 11.69
        Case "Flyer":        w = 8.5: h = 11
        Case "Brochure":     w = 11: h = 8.5: m = 0.375: g = 0.75: cols = 3: pages = 2
        Case "Newsletter":   w = 8.5: h = 11: cols = 3: pages = 4
        Case "Postcard":     w = 6: h = 4: m = 0.25: pages = 2
        Case "BusinessCard": w = 3.5: h = 2: m = 0.15
        Case "GreetingCard": w = 11: h = 8.5: m = 0.5: g = 1: cols = 2: pages = 2
        Case "Poster":       w = 11: h = 17: m = 0.75
        Case "Custom"
            ans = InputBox("Page size in inches (width x height):", "Custom Publication", "8.5 x 11")
            If Len(ans) = 0 Then Exit Sub
            parts = Split(LCase$(Replace(ans, " ", "")), "x")
            If UBound(parts) <> 1 Then PubFail "Enter the size like 8.5 x 11"
            w = Val(parts(0)): h = Val(parts(1))
            If w < 1 Or h < 1 Or w > 22 Or h > 22 Then PubFail "Width and height must be between 1 and 22 inches."
            If w < 4 Or h < 4 Then m = 0.25
        Case Else
            PubFail "Unknown publication type: " & kind
    End Select

    Set doc = Documents.Add
    With doc.PageSetup
        .PageWidth = InchesToPoints(w)
        .PageHeight = InchesToPoints(h)
        .TopMargin = InchesToPoints(m)
        .BottomMargin = InchesToPoints(m)
        .LeftMargin = InchesToPoints(m)
        .RightMargin = InchesToPoints(m)
        .HeaderDistance = InchesToPoints(m / 2)
        .FooterDistance = InchesToPoints(m / 2)
        .DifferentFirstPageHeaderFooter = False
        .OddAndEvenPagesHeaderFooter = False
    End With
    ' Keep the (invisible) body paragraphs tiny so they never push pages around.
    With doc.Content
        .Font.Size = 2
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
    End With
    With doc.Sections(1).Headers(wdHeaderFooterPrimary).Range
        .Font.Size = 2
        .ParagraphFormat.SpaceAfter = 0
    End With

    For i = 2 To pages
        AddPageAtEnd doc
    Next i

    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.View.ShowAll = False
    DrawGuides cols, rows, InchesToPoints(g)

    On Error Resume Next
    doc.GridDistanceHorizontal = InchesToPoints(0.125)
    doc.GridDistanceVertical = InchesToPoints(0.125)
    doc.GridOriginFromMargin = False
    doc.SnapToGrid = True
    doc.SnapToShapes = True
    On Error GoTo 0

    AddStarterContent kind, w, h, m

    Selection.HomeKey wdStory
    FitWholePage
End Sub

' A few ready-made frames so common publication types are not a blank page.
Private Sub AddStarterContent(ByVal kind As String, ByVal w As Single, ByVal h As Single, ByVal m As Single)
    Dim shp As Shape
    Select Case kind
        Case "Flyer"
            Set shp = AddTextBoxOnPage(1, InchesToPoints(m), InchesToPoints(m), _
                InchesToPoints(w - 2 * m), InchesToPoints(1.5), "Event Title")
            FormatFrameText shp, 48, True, wdAlignParagraphCenter
            AddPicturePlaceholderOnPage 1, InchesToPoints(m), InchesToPoints(m + 1.75), _
                InchesToPoints(w - 2 * m), InchesToPoints(4.5)
            Set shp = AddTextBoxOnPage(1, InchesToPoints(m), InchesToPoints(m + 6.5), _
                InchesToPoints(w - 2 * m), InchesToPoints(h - 2 * m - 6.5), _
                "Date  |  Time  |  Location" & vbCr & vbCr & "Describe your event here.")
            FormatFrameText shp, 18, False, wdAlignParagraphCenter
        Case "Newsletter"
            Set shp = AddTextBoxOnPage(1, InchesToPoints(m), InchesToPoints(m), _
                InchesToPoints(w - 2 * m), InchesToPoints(1.25), "Newsletter Title")
            FormatFrameText shp, 40, True, wdAlignParagraphLeft
        Case "BusinessCard"
            Set shp = AddTextBoxOnPage(1, InchesToPoints(m), InchesToPoints(m), _
                InchesToPoints(w - 2 * m), InchesToPoints(h - 2 * m), _
                "Your Name" & vbCr & "Title" & vbCr & "Company" & vbCr & "Phone  |  Email")
            FormatFrameText shp, 10, False, wdAlignParagraphLeft
            shp.TextFrame.TextRange.Paragraphs(1).Range.Font.Size = 14
            shp.TextFrame.TextRange.Paragraphs(1).Range.Font.Bold = True
    End Select
End Sub

'------------------------------------------------------------------------------
' Pages
'------------------------------------------------------------------------------
' Appends a page. The page break goes in just before the final paragraph mark,
' after any anchors in the last paragraph, so existing objects stay put.
Public Sub AddPageAtEnd(Optional doc As Document)
    Dim r As Range
    If doc Is Nothing Then Set doc = ActiveDocument
    Set r = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    r.InsertAfter Chr$(12) & vbCr
End Sub

Public Sub AddPage()
    RequireDoc
    AddPageAtEnd
    GoToPage PageTotal()
End Sub

Public Sub InsertPageAfterCurrent()
    Dim n As Long, r As Range
    RequireDoc
    n = CurrentPageNumber()
    If n >= PageTotal() Then
        AddPage
        Exit Sub
    End If
    ' Find the page break that ends page n and add an empty page right after it.
    Set r = ActiveDocument.Range(PageAnchor(n).Start, PageAnchor(n + 1).Start)
    With r.Find
        .ClearFormatting
        .Text = "^m"
        .Forward = False
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = False
    End With
    If r.Find.Execute Then
        r.Collapse wdCollapseEnd
        r.InsertAfter vbCr & Chr$(12)
    Else
        Set r = PageAnchor(n + 1)
        r.InsertBreak wdPageBreak
    End If
    GoToPage n + 1
End Sub

Public Sub DeleteCurrentPage()
    Dim n As Long, total As Long, i As Long, r As Range
    RequireDoc
    If InMasterPage() Then PubFail "Close the master page first."
    n = CurrentPageNumber()
    total = PageTotal()
    If MsgBox("Delete page " & n & " and everything on it?", vbYesNo + vbQuestion, "Delete Page") = vbNo Then Exit Sub

    ' Remove the objects that live on this page.
    For i = ActiveDocument.Shapes.Count To 1 Step -1
        If ShapePage(ActiveDocument.Shapes(i)) = n Then ActiveDocument.Shapes(i).Delete
    Next i
    If total = 1 Then Exit Sub

    If n < total Then
        Set r = ActiveDocument.Range(PageAnchor(n).Start, PageAnchor(n + 1).Start)
    Else
        ' Last page: remove the page break that ends the previous page instead.
        Set r = ActiveDocument.Range(PageAnchor(n - 1).Start, ActiveDocument.Content.End - 1)
        With r.Find
            .ClearFormatting
            .Text = "^m"
            .Forward = False
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
        End With
        If Not r.Find.Execute Then PubFail "Could not find the page break before this page."
        r.End = ActiveDocument.Content.End - 1
    End If
    r.Delete
    GoToPage IIf(n > PageTotal(), PageTotal(), n)
End Sub

Public Function ShapePage(shp As Shape) As Long
    On Error Resume Next
    ShapePage = shp.Anchor.Information(wdActiveEndPageNumber)
End Function

Public Sub GoToPage(ByVal n As Long)
    If InMasterPage() Then CloseMasterPage
    Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=n
End Sub

Public Sub TogglePagesPanel()
    RequireDoc
    ActiveWindow.DocumentMap = Not ActiveWindow.DocumentMap
End Sub

'------------------------------------------------------------------------------
' Master page (= header of section 1, repeated on every page)
'------------------------------------------------------------------------------
Public Sub EditMasterPage()
    RequireDoc
    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Application.StatusBar = "Editing master page: objects added now appear on every page. Click 'Close Master' when done."
End Sub

Public Sub CloseMasterPage()
    RequireDoc
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    Application.StatusBar = False
End Sub

'------------------------------------------------------------------------------
' Layout guides (drawn on the master page, hidden when printing/exporting)
'------------------------------------------------------------------------------
Public Sub LayoutGuidesDialog()
    Dim ans As String, v() As String, cols As Long, rows As Long, g As Single
    RequireDoc
    ans = InputBox("Columns, Rows, Gutter (inches)" & vbCr & vbCr & _
        "Example: 3, 1, 0.25", "Layout Guides", "3, 1, 0.25")
    If Len(ans) = 0 Then Exit Sub
    v = Split(ans, ",")
    If UBound(v) < 1 Then PubFail "Enter at least columns and rows, e.g. 3, 1"
    cols = Val(v(0)): rows = Val(v(1))
    If UBound(v) >= 2 Then g = Val(v(2)) Else g = 0.25
    If cols < 1 Or cols > 12 Or rows < 1 Or rows > 12 Then PubFail "Columns and rows must be between 1 and 12."
    DrawGuides cols, rows, InchesToPoints(g)
End Sub

Public Sub DrawGuides(ByVal cols As Long, ByVal rows As Long, ByVal gutter As Single)
    Dim ps As PageSetup, hf As HeaderFooter
    Dim pw As Single, ph As Single, l As Single, t As Single, r As Single, b As Single
    Dim cw As Single, rh As Single, x As Single, y As Single, i As Long

    RemoveGuides
    Set ps = ActiveDocument.Sections(1).PageSetup
    Set hf = MasterHeader()
    GetPageSize pw, ph
    l = ps.LeftMargin: t = ps.TopMargin
    r = pw - ps.RightMargin: b = ph - ps.BottomMargin

    ' Margin guides
    AddGuide hf, l, 0, l, ph, GUIDE_MARGIN_RGB, "ML"
    AddGuide hf, r, 0, r, ph, GUIDE_MARGIN_RGB, "MR"
    AddGuide hf, 0, t, pw, t, GUIDE_MARGIN_RGB, "MT"
    AddGuide hf, 0, b, pw, b, GUIDE_MARGIN_RGB, "MB"

    ' Column guides
    If cols > 1 Then
        cw = ((r - l) - gutter * (cols - 1)) / cols
        For i = 1 To cols - 1
            x = l + i * cw + (i - 1) * gutter
            AddGuide hf, x, 0, x, ph, GUIDE_COLUMN_RGB, "C" & i & "a"
            If gutter > 0 Then AddGuide hf, x + gutter, 0, x + gutter, ph, GUIDE_COLUMN_RGB, "C" & i & "b"
        Next i
    End If

    ' Row guides
    If rows > 1 Then
        rh = ((b - t) - gutter * (rows - 1)) / rows
        For i = 1 To rows - 1
            y = t + i * rh + (i - 1) * gutter
            AddGuide hf, 0, y, pw, y, GUIDE_COLUMN_RGB, "R" & i & "a"
            If gutter > 0 Then AddGuide hf, 0, y + gutter, pw, y + gutter, GUIDE_COLUMN_RGB, "R" & i & "b"
        Next i
    End If
End Sub

Private Sub AddGuide(hf As HeaderFooter, ByVal x1 As Single, ByVal y1 As Single, _
                     ByVal x2 As Single, ByVal y2 As Single, ByVal clr As Long, ByVal tag As String)
    Dim shp As Shape
    Set shp = hf.Shapes.AddLine(x1, y1, x2, y2, hf.Range)
    With shp
        .Name = PUB_GUIDE_PREFIX & tag
        .Line.ForeColor.RGB = clr
        .Line.Weight = 0.5
        .Line.DashStyle = msoLineDash
        .WrapFormat.Type = PUB_WRAP_BEHIND
        .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
        .RelativeVerticalPosition = wdRelativeVerticalPositionPage
        .Left = x1
        .Top = y1
        .LockAnchor = True
    End With
End Sub

Public Sub RemoveGuides()
    Dim i As Long, shps As Shapes
    Set shps = MasterHeader().Shapes
    For i = shps.Count To 1 Step -1
        If shps(i).Name Like PUB_GUIDE_PREFIX & "*" Then shps(i).Delete
    Next i
End Sub

Public Function GuidesVisible() As Boolean
    Dim shp As Shape
    For Each shp In MasterHeader().Shapes
        If shp.Name Like PUB_GUIDE_PREFIX & "*" Then
            GuidesVisible = (shp.Visible = msoTrue)
            Exit Function
        End If
    Next shp
End Function

Public Sub SetGuidesVisible(ByVal show As Boolean)
    Dim shp As Shape
    For Each shp In MasterHeader().Shapes
        If shp.Name Like PUB_GUIDE_PREFIX & "*" Then shp.Visible = IIf(show, msoTrue, msoFalse)
    Next shp
End Sub

Public Sub ToggleGuides()
    RequireDoc
    SetGuidesVisible Not GuidesVisible()
End Sub

Public Sub ToggleGrid()
    Dim opts As Object   ' late bound: DisplayGridLines is missing from some type libraries
    RequireDoc
    Set opts = Options
    opts.DisplayGridLines = Not opts.DisplayGridLines
End Sub

Public Sub ToggleSnap()
    RequireDoc
    ActiveDocument.SnapToGrid = Not ActiveDocument.SnapToGrid
    Application.StatusBar = "Snap to grid: " & IIf(ActiveDocument.SnapToGrid, "ON", "OFF")
End Sub

Public Sub FitWholePage()
    RequireDoc
    ActiveWindow.View.Type = wdPrintView
    ActiveWindow.View.Zoom.PageFit = wdPageFitFullPage
End Sub

'------------------------------------------------------------------------------
' Output (guides are hidden while printing / exporting)
'------------------------------------------------------------------------------
Public Sub PrintPublication()
    Dim was As Boolean
    RequireDoc
    was = GuidesVisible()
    SetGuidesVisible False
    On Error Resume Next
    Dialogs(wdDialogFilePrint).Show
    On Error GoTo 0
    SetGuidesVisible was
End Sub

Public Sub ExportPdf()
    Dim was As Boolean, target As String, baseName As String, errMsg As String
    RequireDoc
    baseName = ActiveDocument.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "Export publication as PDF"
        If Len(ActiveDocument.Path) > 0 Then
            .InitialFileName = ActiveDocument.Path & Application.PathSeparator & baseName & ".pdf"
        Else
            .InitialFileName = baseName & ".pdf"
        End If
        If .Show <> -1 Then Exit Sub
        target = .SelectedItems(1)
    End With
    If InStrRev(target, ".") > InStrRev(target, Application.PathSeparator) Then
        target = Left$(target, InStrRev(target, ".") - 1)
    End If
    target = target & ".pdf"

    was = GuidesVisible()
    SetGuidesVisible False
    On Error Resume Next
    ActiveDocument.ExportAsFixedFormat OutputFileName:=target, _
        ExportFormat:=wdExportFormatPDF, OpenAfterExport:=True, _
        OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, _
        Item:=wdExportDocumentContent, IncludeDocProps:=True, _
        CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, _
        BitmapMissingFonts:=True, UseISO19005_1:=False
    If Err.Number <> 0 Then errMsg = Err.Description
    On Error GoTo 0
    SetGuidesVisible was
    If Len(errMsg) > 0 Then PubFail "PDF export failed: " & errMsg
End Sub
