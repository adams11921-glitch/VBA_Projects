Attribute VB_Name = "modPubObjects"
Option Explicit
'==============================================================================
' Word Publisher Kit - frames and objects
'   Text boxes, picture frames/placeholders, linked text (story flow),
'   arrange / align / distribute / wrap, exact size & position.
'
' Every object is created "in front of text", positioned relative to the page
' and anchored to the page it was created on, so it behaves like a Publisher
' object instead of moving with the (empty) body text.
'==============================================================================

Private mCascade As Long            ' offsets successive new objects
Private mPendingLink As String      ' first box picked by the Link command
Private mOverflowUnsupported As Boolean

'------------------------------------------------------------------------------
' Creation helpers (also used by modPubPages for starter content)
'------------------------------------------------------------------------------
Public Function AddTextBoxOnPage(ByVal pageNum As Long, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single, Optional ByVal txt As String = "") As Shape
    Dim shp As Shape
    Set shp = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h, PageAnchor(pageNum))
    StyleTextBox shp
    PlaceOnPage shp, l, t
    If Len(txt) > 0 Then shp.TextFrame.TextRange.Text = txt
    Set AddTextBoxOnPage = shp
End Function

Public Function AddPicturePlaceholderOnPage(ByVal pageNum As Long, ByVal l As Single, ByVal t As Single, _
        ByVal w As Single, ByVal h As Single) As Shape
    Dim shp As Shape
    Set shp = ActiveDocument.Shapes.AddShape(msoShapeRectangle, l, t, w, h, PageAnchor(pageNum))
    StylePlaceholder shp
    PlaceOnPage shp, l, t
    Set AddPicturePlaceholderOnPage = shp
End Function

Public Sub FormatFrameText(shp As Shape, ByVal fontSize As Single, ByVal isBold As Boolean, ByVal alignment As Long)
    With shp.TextFrame.TextRange
        .Font.Size = fontSize
        .Font.Bold = isBold
        .ParagraphFormat.Alignment = alignment
        .ParagraphFormat.SpaceAfter = 0
    End With
End Sub

Public Sub PlaceOnPage(shp As Shape, ByVal l As Single, ByVal t As Single)
    shp.WrapFormat.Type = PUB_WRAP_FRONT
    shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shp.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    shp.Left = l
    shp.Top = t
    shp.LockAnchor = True
End Sub

Private Sub StyleTextBox(shp As Shape)
    shp.Name = UniqueName(PUB_TEXTBOX_PREFIX)
    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse
    With shp.TextFrame
        .MarginLeft = 3.6
        .MarginRight = 3.6
        .MarginTop = 3.6
        .MarginBottom = 3.6
    End With
End Sub

Private Sub StylePlaceholder(shp As Shape)
    Dim tf As Object
    shp.Name = UniqueName(PUB_PICHOLDER_PREFIX)
    shp.Fill.Visible = msoTrue
    shp.Fill.Solid
    shp.Fill.ForeColor.RGB = RGB(235, 235, 235)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(160, 160, 160)
    shp.Line.DashStyle = msoLineDash
    shp.Line.Weight = 1
    With shp.TextFrame.TextRange
        .Text = "Picture placeholder" & vbCr & "Select me, then click Picture"
        .Font.Size = 10
        .Font.Color = RGB(110, 110, 110)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.SpaceAfter = 0
    End With
    On Error Resume Next
    Set tf = shp.TextFrame
    tf.VerticalAnchor = msoAnchorMiddle
End Sub

' Where a new object goes: the master page if it is open, else the current page.
Private Function TargetShapes(ByRef anchor As Range) As Shapes
    If InMasterPage() Then
        Set anchor = MasterHeader().Range
        Set TargetShapes = MasterHeader().Shapes
    Else
        Set anchor = PageAnchor(CurrentPageNumber())
        Set TargetShapes = ActiveDocument.Shapes
    End If
End Function

Private Sub NextDefaultSpot(ByRef l As Single, ByRef t As Single)
    With ActiveDocument.Sections(1).PageSetup
        l = .LeftMargin + mCascade * 18
        t = .TopMargin + mCascade * 18
    End With
    mCascade = (mCascade + 1) Mod 6
End Sub

Private Function DefaultWidth() As Single
    Dim avail As Single, pw As Single, ph As Single
    GetPageSize pw, ph
    With ActiveDocument.Sections(1).PageSetup
        avail = pw - .LeftMargin - .RightMargin
    End With
    If avail < InchesToPoints(3) Then
        DefaultWidth = avail
    Else
        DefaultWidth = InchesToPoints(3)
    End If
End Function

'------------------------------------------------------------------------------
' Ribbon: Objects
'------------------------------------------------------------------------------
Public Sub InsertTextBox()
    Dim anchor As Range, shps As Shapes, shp As Shape, l As Single, t As Single
    RequireDoc
    Set shps = TargetShapes(anchor)
    NextDefaultSpot l, t
    Set shp = shps.AddTextbox(msoTextOrientationHorizontal, l, t, DefaultWidth(), InchesToPoints(1.5), anchor)
    StyleTextBox shp
    PlaceOnPage shp, l, t
    shp.TextFrame.TextRange.Text = "Type your text here"
    shp.TextFrame.TextRange.Select     ' start typing to replace it
End Sub

Public Sub InsertPicturePlaceholder()
    Dim anchor As Range, shps As Shapes, shp As Shape, l As Single, t As Single
    RequireDoc
    Set shps = TargetShapes(anchor)
    NextDefaultSpot l, t
    Set shp = shps.AddShape(msoShapeRectangle, l, t, DefaultWidth(), InchesToPoints(2), anchor)
    StylePlaceholder shp
    PlaceOnPage shp, l, t
    shp.Select
End Sub

' Inserts a floating picture. If a picture placeholder is selected, the picture
' is scaled to fit inside it, centred, and replaces it.
Public Sub InsertPicture()
    Dim f As String, holder As Shape, shp As Shape, anchor As Range, shps As Shapes
    Dim l As Single, t As Single
    RequireDoc
    Set holder = SelectedPlaceholder()
    f = PickPictureFile()
    If Len(f) = 0 Then Exit Sub

    If holder Is Nothing Then
        Set shps = TargetShapes(anchor)
        NextDefaultSpot l, t
        Set shp = shps.AddPicture(FileName:=f, LinkToFile:=False, SaveWithDocument:=True, Anchor:=anchor)
        FitInto shp, DefaultWidth(), InchesToPoints(4)
        PlaceOnPage shp, l, t
    Else
        If holder.Anchor.StoryType = wdMainTextStory Then
            Set shps = ActiveDocument.Shapes
        Else
            Set shps = MasterHeader().Shapes
        End If
        Set shp = shps.AddPicture(FileName:=f, LinkToFile:=False, SaveWithDocument:=True, Anchor:=holder.Anchor)
        shp.WrapFormat.Type = holder.WrapFormat.Type
        shp.RelativeHorizontalPosition = holder.RelativeHorizontalPosition
        shp.RelativeVerticalPosition = holder.RelativeVerticalPosition
        FitInto shp, holder.Width, holder.Height
        shp.Left = holder.Left + (holder.Width - shp.Width) / 2
        shp.Top = holder.Top + (holder.Height - shp.Height) / 2
        shp.LockAnchor = True
        holder.Delete
    End If
    shp.Select
End Sub

Private Function SelectedPlaceholder() As Shape
    Dim sr As ShapeRange
    Set sr = SelectedShapes()
    If sr Is Nothing Then Exit Function
    If sr.Count <> 1 Then Exit Function
    If sr(1).Name Like PUB_PICHOLDER_PREFIX & "*" Then Set SelectedPlaceholder = sr(1)
End Function

Private Function PickPictureFile() As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Insert Picture"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Pictures", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tif;*.tiff;*.svg;*.emf;*.wmf"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then PickPictureFile = .SelectedItems(1)
    End With
End Function

Private Sub FitInto(shp As Shape, ByVal maxW As Single, ByVal maxH As Single)
    Dim k As Single, nw As Single, nh As Single
    If shp.Width <= 0 Or shp.Height <= 0 Then Exit Sub
    k = maxW / shp.Width
    If maxH / shp.Height < k Then k = maxH / shp.Height
    nw = shp.Width * k
    nh = shp.Height * k
    shp.LockAspectRatio = msoFalse
    shp.Width = nw
    shp.Height = nh
    shp.LockAspectRatio = msoTrue
End Sub

'------------------------------------------------------------------------------
' Ribbon: Linked text (story flow between text boxes)
'------------------------------------------------------------------------------
' Select the box with the story, click Link, select the empty box, click Link.
' (Or select both boxes at once and click Link.)
Public Sub LinkTextBoxes()
    Dim sr As ShapeRange, src As Shape, dst As Shape
    Set sr = RequireShapes(1)
    If sr.Count >= 2 Then
        Set src = sr(1): Set dst = sr(2)
        If HasRealText(dst) And Not HasRealText(src) Then Set src = sr(2): Set dst = sr(1)
        mPendingLink = ""
    ElseIf Len(mPendingLink) > 0 Then
        Set src = FindShape(mPendingLink)
        mPendingLink = ""
        If src Is Nothing Then PubFail "The first text box no longer exists. Start the link again."
        Set dst = sr(1)
    Else
        If Not IsTextBox(sr(1)) Then PubFail "Select a text box."
        mPendingLink = sr(1).Name
        MsgBox "Now select the EMPTY text box the story should continue in, " & _
               "then click 'Link' again.", vbInformation, "Link Text Boxes"
        Exit Sub
    End If
    ConnectFrames src, dst
End Sub

Public Sub BreakTextBoxLink()
    Dim sr As ShapeRange, i As Long
    Set sr = RequireShapes(1)
    mPendingLink = ""
    For i = 1 To sr.Count
        If IsTextBox(sr(i)) Then
            If HasForwardLink(sr(i)) Then sr(i).TextFrame.BreakForwardLink
        End If
    Next i
End Sub

' Creates an identical, linked text box on the next page (adding a page if
' needed) so an overflowing story continues there.
Public Sub ContinueOnNextPage()
    Dim sr As ShapeRange, src As Shape, dst As Shape, pg As Long
    Set sr = RequireShapes(1)
    Set src = sr(1)
    If Not IsTextBox(src) Then PubFail "Select the text box whose story should continue."
    If src.Anchor.StoryType <> wdMainTextStory Then PubFail "This works for text boxes on pages, not on the master page."
    If HasForwardLink(src) Then PubFail "This text box already continues in another box. Break the link first."

    pg = ShapePage(src) + 1
    If pg > PageTotal() Then AddPageAtEnd
    Set dst = ActiveDocument.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        src.Left, src.Top, src.Width, src.Height, PageAnchor(pg))
    StyleTextBox dst
    dst.WrapFormat.Type = src.WrapFormat.Type
    dst.RelativeHorizontalPosition = src.RelativeHorizontalPosition
    dst.RelativeVerticalPosition = src.RelativeVerticalPosition
    dst.Left = src.Left
    dst.Top = src.Top
    dst.LockAnchor = True
    ConnectFrames src, dst
    dst.Select
End Sub

Public Sub CheckOverflow()
    Dim shp As Shape, list As String, firstHit As Shape, n As Long
    RequireDoc
    mOverflowUnsupported = False
    For Each shp In ActiveDocument.Shapes
        If IsOverflowing(shp) Then
            n = n + 1
            If firstHit Is Nothing Then Set firstHit = shp
            list = list & vbCr & "   - " & shp.Name & "  (page " & ShapePage(shp) & ")"
        End If
    Next shp
    If mOverflowUnsupported Then
        MsgBox "This version of Word can't report text overflow.", vbInformation, "Text Overflow"
    ElseIf n = 0 Then
        MsgBox "No text is overflowing.", vbInformation, "Text Overflow"
    Else
        firstHit.Select
        MsgBox n & " text box(es) have hidden overflow text:" & list & vbCr & vbCr & _
            "Enlarge the box, make the text smaller, or use 'Continue on Next Page'.", _
            vbExclamation, "Text Overflow"
    End If
End Sub

Private Sub ConnectFrames(src As Shape, dst As Shape)
    If src.Name = dst.Name Then PubFail "Pick two different text boxes."
    If Not IsTextBox(src) Or Not IsTextBox(dst) Then PubFail "Both objects must be text boxes."
    If HasForwardLink(src) Then PubFail "'" & src.Name & "' already continues in another box. Break the link first."
    If HasRealText(dst) Then
        If MsgBox("The second text box already has text. Clear it so the story can flow into it?", _
                  vbYesNo + vbQuestion, "Link Text Boxes") = vbNo Then Exit Sub
        dst.TextFrame.DeleteText
    End If
    If Not src.TextFrame.ValidLinkTarget(dst) Then
        PubFail "Word can't link these boxes. The second box must be empty, not already " & _
                "linked, and in the same place (both on pages, or both on the master page)."
    End If
    src.TextFrame.Next = dst.TextFrame
    Application.StatusBar = "Linked " & src.Name & "  ->  " & dst.Name
End Sub

Private Function IsTextBox(shp As Shape) As Boolean
    IsTextBox = (shp.Type = msoTextBox)
End Function

Private Function HasRealText(shp As Shape) As Boolean
    On Error Resume Next
    HasRealText = shp.TextFrame.HasText
End Function

Private Function HasForwardLink(shp As Shape) As Boolean
    On Error Resume Next
    HasForwardLink = Not (shp.TextFrame.Next Is Nothing)
End Function

Private Function IsOverflowing(shp As Shape) As Boolean
    Dim tf As Object   ' late bound: Overflowing is missing from older type libraries
    If Not IsTextBox(shp) Then Exit Function
    If HasForwardLink(shp) Then Exit Function
    On Error Resume Next
    Set tf = shp.TextFrame
    IsOverflowing = tf.Overflowing
    If Err.Number = 438 Then mOverflowUnsupported = True
End Function

'------------------------------------------------------------------------------
' Ribbon: Arrange
'------------------------------------------------------------------------------
Public Sub ArrangeZ(ByVal cmd As MsoZOrderCmd)
    RequireShapes(1).ZOrder cmd
End Sub

' One object: aligns to the page. Several: aligns them to each other.
Public Sub AlignShapes(ByVal cmd As MsoAlignCmd)
    Dim sr As ShapeRange
    Set sr = RequireShapes(1)
    If sr.Count = 1 Then
        sr.Align cmd, msoTrue
    Else
        sr.Align cmd, msoFalse
    End If
End Sub

' Three or more objects: spaced evenly between the outer two. Fewer: across the page.
Public Sub DistributeShapes(ByVal cmd As MsoDistributeCmd)
    Dim sr As ShapeRange
    Set sr = RequireShapes(1)
    If sr.Count >= 3 Then
        sr.Distribute cmd, msoFalse
    Else
        sr.Distribute cmd, msoTrue
    End If
End Sub

Public Sub SetWrap(ByVal wrapType As Long)
    RequireShapes(1).WrapFormat.Type = wrapType
End Sub

Public Sub GroupShapes()
    RequireShapes(2).Group.Select
End Sub

Public Sub UngroupShapes()
    Dim sr As ShapeRange
    Set sr = RequireShapes(1)
    If sr(1).Type <> msoGroup Then PubFail "Select a group first."
    sr.Ungroup.Select
End Sub

Public Sub ShowSelectionPane()
    RequireDoc
    On Error Resume Next
    CommandBars.ExecuteMso "SelectionPane"
    If Err.Number <> 0 Then
        On Error GoTo 0
        PubFail "Open it from Home > Select > Selection Pane."
    End If
End Sub

'------------------------------------------------------------------------------
' Ribbon: Size & Position (exact measurements, like Publisher's measurement bar)
'------------------------------------------------------------------------------
Public Sub SetPositionSize()
    Dim sr As ShapeRange, shp As Shape, ans As String, v() As String, lockState As Long
    Set sr = RequireShapes(1)
    If sr.Count > 1 Then PubFail "Select just one object."
    Set shp = sr(1)
    ans = InputBox("Left, Top, Width, Height in inches." & vbCr & _
        "Left and Top are measured from the top-left corner of the page.", "Size & Position", _
        InchText(shp.Left) & ", " & InchText(shp.Top) & ", " & InchText(shp.Width) & ", " & InchText(shp.Height))
    If Len(ans) = 0 Then Exit Sub
    v = Split(ans, ",")
    If UBound(v) <> 3 Then PubFail "Enter four numbers separated by commas."

    lockState = shp.LockAspectRatio
    shp.LockAspectRatio = msoFalse
    shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
    shp.RelativeVerticalPosition = wdRelativeVerticalPositionPage
    shp.Width = InchesToPoints(Val(v(2)))
    shp.Height = InchesToPoints(Val(v(3)))
    shp.Left = InchesToPoints(Val(v(0)))
    shp.Top = InchesToPoints(Val(v(1)))
    shp.LockAspectRatio = lockState
End Sub
