Attribute VB_Name = "modPubCore"
Option Explicit
'==============================================================================
' Word Publisher Kit - core
'   * Ribbon dispatcher: every button on the "Publisher" tab calls Pub_OnAction
'   * Shared helpers (current page, page anchors, selected shapes, naming)
'==============================================================================

Public Const PUB_GUIDE_PREFIX As String = "PubGuide_"
Public Const PUB_PICHOLDER_PREFIX As String = "PubPicHolder_"
Public Const PUB_TEXTBOX_PREFIX As String = "PubText_"

' WdWrapType values, declared literally so the code compiles on every Word build
Public Const PUB_WRAP_SQUARE As Long = 0
Public Const PUB_WRAP_TIGHT As Long = 1
Public Const PUB_WRAP_FRONT As Long = 3
Public Const PUB_WRAP_TOPBOTTOM As Long = 4
Public Const PUB_WRAP_BEHIND As Long = 5

Private mNameCounter As Long

'------------------------------------------------------------------------------
' Ribbon entry point
'------------------------------------------------------------------------------
Public Sub Pub_OnAction(control As IRibbonControl)
    RunCommand control.ID, control.Tag
End Sub

' For buttons that appear on both tabs (ribbon IDs must be unique): the
' command name is in the button's tag.
Public Sub Pub_OnTagAction(control As IRibbonControl)
    RunCommand control.Tag, ""
End Sub

Public Sub RunCommand(ByVal cmd As String, ByVal tag As String)
    On Error GoTo ErrHandler
    Select Case cmd
        ' --- Publication ---------------------------------------------------
        Case "New_Letter": NewPublication "Letter"
        Case "New_A4": NewPublication "A4"
        Case "New_Flyer": NewPublication "Flyer"
        Case "New_Brochure": NewPublication "Brochure"
        Case "New_Newsletter": NewPublication "Newsletter"
        Case "New_Postcard": NewPublication "Postcard"
        Case "New_BusinessCard": NewPublication "BusinessCard"
        Case "New_GreetingCard": NewPublication "GreetingCard"
        Case "New_Poster": NewPublication "Poster"
        Case "New_Custom": NewPublication "Custom"
        Case "PageSetup": RequireDoc: Dialogs(wdDialogFilePageSetup).Show

        ' --- Bulletin --------------------------------------------------------
        Case "Bul_NewLegal": NewBulletin "Legal"
        Case "Bul_NewLetter": NewBulletin "Letter"
        Case "Bul_NextWeek": StartNextWeek
        Case "Bul_SetDate": SetBulletinDate
        Case "Bul_Settings": EditChurchSettings
        Case "Bul_Item": InsertWorshipItem
        Case "Bul_Hymn": InsertHymn
        Case "Bul_Scripture": InsertScripture
        Case "Bul_Responsive": InsertResponsiveReading
        Case "Bul_Sermon": InsertSermon
        Case "Bul_HeadingBar": InsertHeadingBar
        Case "Bul_Announcement": InsertAnnouncement
        Case "Bul_Schedule": InsertScheduleTable
        Case "Bul_PageBreak": InsertBulletinPageBreak
        Case "Bul_PrintBooklet": PrintBooklet
        Case "Bul_BookletView": ToggleBookletView
        Case "Bul_CheckPages": CheckBookletPages
        Case "Bul_BookletPdf": ExportBookletPdf

        ' --- Pages ---------------------------------------------------------
        Case "Page_Add": AddPage
        Case "Page_InsertAfter": InsertPageAfterCurrent
        Case "Page_Delete": DeleteCurrentPage
        Case "Page_Panel": TogglePagesPanel
        Case "Master_Edit": EditMasterPage
        Case "Master_Close": CloseMasterPage

        ' --- Objects -------------------------------------------------------
        Case "Obj_TextBox": InsertTextBox
        Case "Obj_Picture": InsertPicture
        Case "Obj_PicHolder": InsertPicturePlaceholder

        ' --- Linked text ---------------------------------------------------
        Case "Link_Create": LinkTextBoxes
        Case "Link_Break": BreakTextBoxLink
        Case "Link_Continue": ContinueOnNextPage
        Case "Link_Overflow": CheckOverflow

        ' --- Arrange -------------------------------------------------------
        Case "Arr_Front": ArrangeZ msoBringToFront
        Case "Arr_Back": ArrangeZ msoSendToBack
        Case "Arr_Forward": ArrangeZ msoBringForward
        Case "Arr_Backward": ArrangeZ msoSendBackward
        Case "Align_Left": AlignShapes msoAlignLefts
        Case "Align_Center": AlignShapes msoAlignCenters
        Case "Align_Right": AlignShapes msoAlignRights
        Case "Align_Top": AlignShapes msoAlignTops
        Case "Align_Middle": AlignShapes msoAlignMiddles
        Case "Align_Bottom": AlignShapes msoAlignBottoms
        Case "Dist_Horiz": DistributeShapes msoDistributeHorizontally
        Case "Dist_Vert": DistributeShapes msoDistributeVertically
        Case "Wrap_Front": SetWrap PUB_WRAP_FRONT
        Case "Wrap_Behind": SetWrap PUB_WRAP_BEHIND
        Case "Wrap_Square": SetWrap PUB_WRAP_SQUARE
        Case "Wrap_Tight": SetWrap PUB_WRAP_TIGHT
        Case "Wrap_TopBottom": SetWrap PUB_WRAP_TOPBOTTOM
        Case "Obj_Group": GroupShapes
        Case "Obj_Ungroup": UngroupShapes
        Case "Obj_SelPane": ShowSelectionPane

        ' --- Layout --------------------------------------------------------
        Case "Guides_Setup": LayoutGuidesDialog
        Case "Guides_Toggle": ToggleGuides
        Case "Grid_Toggle": ToggleGrid
        Case "Snap_Toggle": ToggleSnap
        Case "Obj_Position": SetPositionSize
        Case "View_FitPage": FitWholePage

        ' --- Output --------------------------------------------------------
        Case "Out_Print": PrintPublication
        Case "Out_Pdf": ExportPdf

        Case Else
            If Left$(cmd, 4) = "Sty_" Then
                ApplyBulletinStyle tag
            Else
                MsgBox "Unknown command: " & cmd, vbExclamation, "Publisher Kit"
            End If
    End Select
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Publisher Kit"
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Public Sub RequireDoc()
    If Documents.Count = 0 Then
        Err.Raise vbObjectError + 513, "Publisher Kit", "Open or create a publication first."
    End If
End Sub

Public Sub PubFail(ByVal msg As String)
    Err.Raise vbObjectError + 514, "Publisher Kit", msg
End Sub

' Size of one page as it appears on screen. With "Book fold" printing Word may
' report the whole sheet (e.g. 14 x 8.5), so halve it to get the folded page.
Public Sub GetPageSize(ByRef w As Single, ByRef h As Single, Optional doc As Document)
    If doc Is Nothing Then Set doc = ActiveDocument
    With doc.Sections(1).PageSetup
        w = .PageWidth
        h = .PageHeight
        If .BookFoldPrinting And w > h Then w = w / 2
    End With
End Sub

Public Function PageTotal() As Long
    PageTotal = ActiveDocument.ComputeStatistics(wdStatisticPages)
End Function

' Page the cursor / selected object is on.
Public Function CurrentPageNumber() As Long
    Dim n As Long
    On Error Resume Next
    n = Selection.Information(wdActiveEndPageNumber)
    On Error GoTo 0
    If n < 1 Then n = 1
    If n > PageTotal() Then n = PageTotal()
    CurrentPageNumber = n
End Function

' Collapsed range at the first paragraph that starts on the given page.
' Floating objects anchored here are drawn on that page.
Public Function PageAnchor(ByVal pageNum As Long) As Range
    Dim r As Range, para As Paragraph
    If pageNum < 1 Then pageNum = 1
    If pageNum > PageTotal() Then pageNum = PageTotal()
    Set r = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=pageNum)
    r.Collapse wdCollapseStart
    ' The anchor must be the start of a paragraph, otherwise the object would
    ' belong to a paragraph that began on the previous page.
    Set para = r.Paragraphs(1)
    If para.Range.Start <> r.Start Then
        If Not para.Next Is Nothing Then
            Set r = para.Next.Range
            r.Collapse wdCollapseStart
        End If
    End If
    Set PageAnchor = r
End Function

' True when the cursor is in a header/footer (the "master page").
Public Function InMasterPage() As Boolean
    Select Case Selection.StoryType
        Case wdPrimaryHeaderStory, wdFirstPageHeaderStory, wdEvenPagesHeaderStory, _
             wdPrimaryFooterStory, wdFirstPageFooterStory, wdEvenPagesFooterStory
            InMasterPage = True
    End Select
End Function

Public Function MasterHeader() As HeaderFooter
    Set MasterHeader = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
End Function

' Shapes the user has selected. If the cursor is inside a text box, that box.
' Returns Nothing when no object is selected.
Public Function SelectedShapes() As ShapeRange
    Dim shp As Shape
    On Error Resume Next
    If Selection.Type = wdSelectionShape Then
        Set SelectedShapes = Selection.ShapeRange
        Exit Function
    End If
    If Selection.StoryType = wdTextFrameStory Then
        For Each shp In ActiveDocument.Shapes
            If FrameHoldsSelection(shp) Then
                Set SelectedShapes = ActiveDocument.Shapes.Range(Array(shp.Name))
                Exit Function
            End If
        Next shp
        For Each shp In MasterHeader().Shapes
            If FrameHoldsSelection(shp) Then
                Set SelectedShapes = MasterHeader().Shapes.Range(Array(shp.Name))
                Exit Function
            End If
        Next shp
    End If
End Function

Private Function FrameHoldsSelection(shp As Shape) As Boolean
    On Error Resume Next
    FrameHoldsSelection = Selection.Range.InRange(shp.TextFrame.TextRange)
End Function

Public Function RequireShapes(ByVal minCount As Long) As ShapeRange
    Dim sr As ShapeRange
    RequireDoc
    Set sr = SelectedShapes()
    If sr Is Nothing Then PubFail "Select an object first."
    If sr.Count < minCount Then PubFail "Select at least " & minCount & " objects (hold Shift and click)."
    Set RequireShapes = sr
End Function

Public Function FindShape(ByVal shapeName As String) As Shape
    Dim shp As Shape
    For Each shp In ActiveDocument.Shapes
        If shp.Name = shapeName Then Set FindShape = shp: Exit Function
    Next shp
    For Each shp In MasterHeader().Shapes
        If shp.Name = shapeName Then Set FindShape = shp: Exit Function
    Next shp
End Function

Public Function UniqueName(ByVal prefix As String) As String
    mNameCounter = mNameCounter + 1
    UniqueName = prefix & Format(Now, "yymmddhhnnss") & "_" & mNameCounter
End Function

' Inches as text with a "." decimal separator regardless of regional settings.
Public Function InchText(ByVal points As Single) As String
    InchText = Trim$(Str$(Round(PointsToInches(points), 2)))
End Function
