VERSION 5.00
Begin VB.UserControl ucCoolList 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2010
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   95
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   134
   Begin VB.VScrollBar sbVert 
      Height          =   1125
      Left            =   1215
      Max             =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'Kein
      Height          =   510
      Left            =   1620
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   105
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.PictureBox picList 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      FillStyle       =   0  'Ausgefüllt
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   78
      TabIndex        =   2
      Top             =   45
      Width           =   1170
   End
End
Attribute VB_Name = "ucCoolList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' UserControl:   ucCoolList 1.2
' Author:        Carles P.V.
' Dependencies:
' First release: 2002
' Last revision: 2005.09.19
'================================================

Option Explicit

'-- API:

Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT2) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT2, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT2, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT2) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT2, ByVal dx As Long, ByVal dy As Long) As Long

Private Type TRIVERTEX
    x     As Long
    y     As Long
    R     As Integer
    G     As Integer
    B     As Integer
    Alpha As Integer
End Type

Private Type RGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private Type GRADIENT_RECT
    UpperLeft  As Long
    LowerRight As Long
End Type

Private Type RECT2
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const PS_SOLID As Long = 0

Private Const GRADIENT_FILL_RECT_H As Long = &H0
Private Const GRADIENT_FILL_RECT_V As Long = &H1

Private Const DT_LEFT       As Long = &H0
Private Const DT_CENTER     As Long = &H1
Private Const DT_RIGHT      As Long = &H2
Private Const DT_VCENTER    As Long = &H4
Private Const DT_WORDBREAK  As Long = &H10
Private Const DT_SINGLELINE As Long = &H20

'-- Public enums.:

Public Enum AlignmentCts
    [AlignLeft] = 0
    [AlignCenter]
    [AlignRight]
End Enum

Public Enum AppearanceCts
    [Flat] = 0
    [3D]
End Enum

Public Enum BorderStyleCts
    [None] = 0
    [Fixed Single]
End Enum

Public Enum OrderTypeCts
    [Ascendent] = 0
    [Descendent]
End Enum

Public Enum SelectModeCts
    [Single] = 0
    [Multiple]
End Enum

Public Enum SelectModeStyleCts
    [Standard] = 0
    [Dither]
    [Gradient_V]
    [Gradient_H]
    [Box]
    [Underline]
    [ByPicture]
End Enum

'-- Private types:

Private Type uItem
    Text         As String
    Icon         As Integer
    IconSelected As Integer
End Type

'-- Private variables:

Private m_uList()          As uItem    ' List array of items (Text, icons)
Private m_bSelected()      As Boolean  ' List array of items (Selected/Unselected)
Private m_nItems           As Integer  ' Number of Items

Private m_nLastBar         As Integer  ' Last scroll bar value
Private m_nLastItem        As Integer  ' Last Selected item
Private m_snLastY          As Single   ' Last Y value [pixels] (prevents item repaint)
Private m_bAnchorSelected  As Boolean  ' Anchor item value (multiple selection).
                                       '  Case extended selection: all selected items
                                       '  will be set to Anchor selection state.
                                        
Private m_bEnsureVisible   As Boolean  ' Ensure visible last m_bSelected item (ListIndex)

Private m_uRctItem()       As RECT2    ' Item rectangle
Private m_uRctText()       As RECT2    ' Item text rectangle
Private m_uPtIcon()        As POINTAPI ' Item icon position

Private m_nTmpItemHeight   As Integer  ' Item height [pixels]
Private m_nVisibleRows     As Integer  ' Visible rows in control area
Private m_bScrolling       As Boolean  ' Scrolling by mouse
Private m_lScrollingY      As Long     ' Y Scrolling coordinate flag (scroll speed = f(Y))
Private m_bHasFocus        As Boolean  ' Control has focus
Private m_bResizing        As Boolean  ' Prevent repaints when Resizing

Private m_lpoIL            As Object   ' Will point to ImageList control
Private m_nILScale         As Integer  ' ImageList parent scale mode

Private m_lClrBack         As Long     ' Back color [Normal]
Private m_lClrBackSel      As Long     ' Back color [Selected]
Private m_lClrFont         As Long     ' Font color [Normal]
Private m_lClrFontSel      As Long     ' Font color [Selected]
Private m_lClrGradient1    As RGB      ' Gradient color from [Selected]
Private m_lClrGradient2    As RGB      ' Gradient color  to  [Selected]
Private m_lClrBox          As Long     ' Box border color

Private WithEvents m_oFont As StdFont  ' Font object
Attribute m_oFont.VB_VarHelpID = -1

'-- Property variables:

Private m_Alignment        As AlignmentCts
Private m_Apeareance       As AppearanceCts
Private m_BackNormal       As OLE_COLOR
Private m_BackSelected     As OLE_COLOR
Private m_BackSelectedG1   As OLE_COLOR
Private m_BackSelectedG2   As OLE_COLOR
Private m_BoxBorder        As OLE_COLOR
Private m_BoxOffset        As Integer
Private m_BoxRadius        As Integer
Private m_Focus            As Boolean
Private m_FontNormal       As OLE_COLOR
Private m_FontSelected     As OLE_COLOR
Private m_HoverSelection   As Boolean
Private m_ItemHeight       As Integer
Private m_ItemHeightAuto   As Boolean
Private m_ItemOffset       As Integer
Private m_ItemTextLeft     As Integer
Private m_ListIndex        As Integer
Private m_OrderType        As OrderTypeCts
Private m_ScrollBarWidth   As Integer
Private m_SelectionPicture As Picture
Private m_SelectMode       As SelectModeCts
Private m_SelectModeStyle  As SelectModeStyleCts
Private m_TopIndex         As Integer
Private m_WordWrap         As Boolean

'-- Default property values:

Private Const m_def_Appearance      As Long = [3D]
Private Const m_def_Alignment       As Long = [AlignLeft]
Private Const m_def_BackNormal      As Long = vbWindowBackground
Private Const m_def_BackSelected    As Long = vbHighlight
Private Const m_def_BackSelectedG1  As Long = vbHighlight
Private Const m_def_BackSelectedG2  As Long = vbWindowBackground
Private Const m_def_BorderStyle     As Long = [Fixed Single]
Private Const m_def_BoxBorder       As Long = vbHighlightText
Private Const m_def_BoxOffset       As Long = 1
Private Const m_def_BoxRadius       As Long = 0
Private Const m_def_Focus           As Boolean = True
Private Const m_def_FontNormal      As Long = vbWindowText
Private Const m_def_FontSelected    As Long = vbHighlightText
Private Const m_def_HoverSelection  As Boolean = False
Private Const m_def_ItemHeightAuto  As Boolean = True
Private Const m_def_ItemOffset      As Long = 0
Private Const m_def_ItemTextLeft    As Long = 2
Private Const m_def_OrderType       As Long = [Ascendent]
Private Const m_def_ScrollBarWidth  As Long = 13
Private Const m_def_SelectMode      As Long = [Single]
Private Const m_def_SelectModeStyle As Long = [Standard]
Private Const m_def_WordWrap        As Boolean = True

'-- Events:

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event ListIndexChange()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event Scroll()
Public Event TopIndexChange()



'========================================================================================
' UserControl initialitation, focus, size, refresh, termination
'========================================================================================

Private Sub UserControl_Initialize()
    
    '-- Initialize arrays
    ReDim m_uList(0)
    ReDim m_bSelected(0)
    
    '-- Initialize position flags
    m_bEnsureVisible = True ' Ensure visible last selected
    m_nLastItem = -1        ' Last selected
    m_snLastY = -1          ' Last Y coordinate
    
    '-- Initialize font object
    Set m_oFont = New StdFont
End Sub

Private Sub UserControl_EnterFocus()
    m_bHasFocus = True
    Call pvDrawFocus(m_ListIndex)
End Sub

Private Sub UserControl_ExitFocus()
    m_bHasFocus = False
    Call pvDrawItem(m_ListIndex)
End Sub

Private Sub UserControl_Resize()
    
    '-- Set item height
    If (m_ItemHeightAuto) Then
        m_nTmpItemHeight = picList.TextHeight(vbNullString)
      Else
        If (m_ItemHeight < picList.TextHeight(vbNullString)) Then
            m_nTmpItemHeight = picList.TextHeight(vbNullString)
          Else
            m_nTmpItemHeight = m_ItemHeight
        End If
    End If
    
    '-- Get visible rows and re-adjust control height
    m_nVisibleRows = ScaleHeight \ m_nTmpItemHeight
    Height = (m_nVisibleRows) * m_nTmpItemHeight * Screen.TwipsPerPixelX + (Height - ScaleHeight * Screen.TwipsPerPixelY)
    
    '-- Locate and resize drawing area, calc. rects and readjust scroll bar
    m_bResizing = True
    Call picList.Move(0, 0, ScaleWidth - IIf(sbVert.Visible, sbVert.Width, 0), ScaleHeight)
    With sbVert
        Call .Move(ScaleWidth - .Width, 0, .Width, ScaleHeight)
        .Visible = False
    End With
    ReDim m_uRctItem(m_nVisibleRows - 1)
    ReDim m_uRctText(m_nVisibleRows - 1)
    ReDim m_uPtIcon(m_nVisibleRows - 1)
    Call pvCalculateRects
    Call pvReadjustScrollBar
    m_bResizing = False
End Sub

Private Sub picList_Paint()
    
  Dim uRct As RECT2
  
    If (Not Ambient.UserMode) Then
        
        Call picList.Cls

        Select Case m_Alignment
            Case 0: picList.CurrentX = m_ItemTextLeft + m_ItemOffset
            Case 1: picList.CurrentX = (ScaleWidth - picList.TextWidth(Ambient.DisplayName)) \ 2
            Case 2: picList.CurrentX = (ScaleWidth - picList.TextWidth(Ambient.DisplayName)) - m_ItemOffset
        End Select
        picList.CurrentY = m_ItemOffset
                    
        Call SetTextColor(picList.hdc, m_lClrFont)
        picList.Print Ambient.DisplayName
       
        Call SetRect(uRct, 0, 0, ScaleWidth, m_nTmpItemHeight)
        Call DrawFocusRect(picList.hdc, uRct)

      Else
        If (Not m_bResizing) Then
            Call pvDrawList
        End If
    End If
End Sub

Private Sub UserControl_Terminate()

    Erase m_uList()
    Erase m_bSelected()
    Set m_lpoIL = Nothing
    m_bScrolling = False
End Sub

'========================================================================================
' ScrollBar
'========================================================================================

Private Sub sbVert_Change()

    If (m_nLastBar <> sbVert.Value) Then
        m_nLastBar = sbVert.Value
        m_snLastY = -1
        If (txtEdit.Visible) Then
            Call txtEdit_LostFocus
        End If
        If (m_ListIndex = m_nLastItem) Then
            Call pvDrawList
        End If
        RaiseEvent Scroll
        RaiseEvent TopIndexChange
    End If
End Sub

Private Sub sbVert_Scroll()
    Call sbVert_Change
    RaiseEvent Scroll
End Sub

'========================================================================================
' Scrolling / Events
'========================================================================================

'-- Click()
Private Sub picList_Click()
    If (m_ListIndex > -1) Then RaiseEvent Click
End Sub

'-- DblClick()
Private Sub picList_DblClick()
    If (m_ListIndex > -1) Then RaiseEvent DblClick
End Sub

'-- KeyDown(KeyCode, Shift)
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (m_nItems = 0 Or m_ListIndex = -1) Then
        RaiseEvent KeyDown(KeyCode, Shift)
        Exit Sub
    End If
    
    Select Case KeyCode
        
        Case vbKeyUp
            If (m_ListIndex > 0) Then ListIndex = ListIndex - 1
        
        Case vbKeyDown
            If (m_ListIndex < m_nItems - 1) Then ListIndex = ListIndex + 1
        
        Case vbKeyPageUp
            If (m_ListIndex > m_nVisibleRows) Then
                ListIndex = ListIndex - m_nVisibleRows
              Else
                ListIndex = 0
            End If
        
        Case vbKeyPageDown
            If (m_ListIndex < m_nItems - m_nVisibleRows - 1) Then
                ListIndex = ListIndex + m_nVisibleRows
              Else
                ListIndex = m_nItems - 1
            End If
        
        Case vbKeyHome
            ListIndex = 0
        
        Case vbKeyEnd
            ListIndex = m_nItems - 1
        
        Case vbKeySpace
            If (m_SelectMode <> [Single] And m_ListIndex > -1) Then
                m_bSelected(m_ListIndex) = Not m_bSelected(m_ListIndex)
                Call pvDrawItem(m_ListIndex)
                Call pvDrawFocus(m_ListIndex)
            End If
            RaiseEvent Click
    End Select
    
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

'-- KeyPress(KeyAscii)
Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'-- KeyPress(KeyCode, Shift)
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'-- MouseDown(Button, Shift, x, y)
Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
  Dim nIndex As Integer
  
    If (Button = vbRightButton) Then
        RaiseEvent MouseDown(Button, Shift, x, y)
        Exit Sub
    End If
   
    nIndex = sbVert.Value + Int(y / m_nTmpItemHeight)
    
    If (nIndex >= 0 And nIndex < m_nItems) Then
        Select Case m_SelectMode
            Case [Single]
                m_bSelected(nIndex) = True
            Case [Multiple]
                m_bSelected(nIndex) = Not m_bSelected(nIndex)
                m_bAnchorSelected = m_bSelected(nIndex)
        End Select
        
        m_snLastY = y
        ListIndex = nIndex
    End If
    
    m_bScrolling = True
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'-- MouseMove(Button, Shift, x, y)
Private Sub picList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
  Dim nIndex As Integer
  
    m_lScrollingY = y
    
    If (y < 0) Then
        Call pvScrollUp
        RaiseEvent MouseMove(Button, Shift, x, y)
        Exit Sub
    End If
    If (y > ScaleHeight) Then
        Call pvScrollDown
        RaiseEvent MouseMove(Button, Shift, x, y)
        Exit Sub
    End If
                
    If (m_HoverSelection Or Button) And (y \ m_nTmpItemHeight <> m_snLastY \ m_nTmpItemHeight) Then
     
        nIndex = sbVert.Value + (y \ m_nTmpItemHeight)
        
        If (nIndex >= 0 And nIndex < m_nItems) Then
            m_bSelected(nIndex) = m_bAnchorSelected
            ListIndex = nIndex
            m_snLastY = y
        End If
    End If
    
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'-- MouseUp(Button, Shift, x, y)
Private Sub picList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_bScrolling = False
    m_bAnchorSelected = True
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'========================================================================================
' Methods
'========================================================================================

'-- SetImageList
Public Sub SetImageList(ImageListControl)
    
    Set m_lpoIL = ImageListControl
    
    On Error Resume Next
        m_nILScale = m_lpoIL.Parent.ScaleMode
    On Error GoTo 0
    
    Call picList_Paint
End Sub

'-- AddItem
'-- 0 , ... , n-1 [n = ListCount]

Public Sub AddItem(ByVal Text As Variant, _
                   Optional ByVal Icon As Integer, _
                   Optional ByVal IconSelected As Integer)
                
    With m_uList(m_nItems)
        .Text = CStr(Text)
        .Icon = Icon
        .IconSelected = IconSelected
    End With
    m_nItems = m_nItems + 1
    
    ReDim Preserve m_uList(m_nItems)
    ReDim Preserve m_bSelected(m_nItems)
    
    Call pvReadjustScrollBar
    If (m_nItems < m_nVisibleRows + 1) Then
        Call pvDrawItem(m_nItems - 1)
    End If
End Sub

'-- InsertItem
Public Sub InsertItem(ByVal index As Integer, _
                      ByVal Text As Variant, _
                      Optional ByVal Icon As Integer, _
                      Optional ByVal IconSelected As Integer)
     
  Dim i As Long
  
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
        
    m_nItems = m_nItems + 1
    ReDim Preserve m_uList(m_nItems)
    ReDim Preserve m_bSelected(m_nItems)

    For i = m_nItems - 1 To index Step -1
        m_uList(i + 1) = m_uList(i)
        m_bSelected(i + 1) = m_bSelected(i)
    Next i
      
    With m_uList(index)
        .Text = CStr(Text)
        .Icon = Icon
        .IconSelected = IconSelected
    End With
    m_bSelected(index) = False
        
    Call pvReadjustScrollBar
    m_bEnsureVisible = False
    If (m_ListIndex > -1 And index <= m_ListIndex) Then
        ListIndex = ListIndex + 1
    End If
    Call picList_Paint
End Sub

'-- ModifyItem
Public Sub ModifyItem(ByVal index As Integer, _
                      Optional ByVal Text As Variant = vbEmpty, _
                      Optional ByVal Icon As Integer = -1, _
                      Optional ByVal IconSelected As Integer = -1)
    
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    
    If (Text <> vbEmpty) Then
        m_uList(index).Text = CStr(Text)
    End If
    If (Icon > -1) Then
        m_uList(index).Icon = Icon
    End If
    If (IconSelected > -1) Then
        m_uList(index).IconSelected = IconSelected
    End If
    
    Call pvDrawItem(index)
    Call pvDrawFocus(m_ListIndex)
End Sub

'-- RemoveItem
Public Sub RemoveItem(ByVal index As Integer)

  Dim i As Long
  
    If (m_nItems = 0 Or index > m_nItems - 1) Then Err.Raise 381
        
    If (index < m_nItems) Then
        For i = index To m_nItems - 1
            m_uList(i) = m_uList(i + 1)
            m_bSelected(i) = m_bSelected(i + 1)
        Next i
    End If
    
    m_nItems = m_nItems - 1
    ReDim Preserve m_uList(m_nItems)
    ReDim Preserve m_bSelected(m_nItems)
        
    Call pvReadjustScrollBar
    m_bEnsureVisible = False
    
    If (index < m_ListIndex) Then
        If (m_ListIndex > -1) Then
            ListIndex = ListIndex - 1
        End If
      ElseIf (index = m_ListIndex) Then
        ListIndex = -1
    End If
    
    If (m_nItems < m_nVisibleRows) Then
        Call picList.Cls
    End If
    Call picList_Paint
End Sub

'-- FindFirst
Public Function FindFirst(ByVal FindString As String, _
                          Optional ByVal StartIndex As Integer = 0, _
                          Optional ByVal StartWith As Boolean = 0) As Integer
  Dim i As Long
    
    If (m_nItems = 0) Then Err.Raise 381
    
    For i = StartIndex To m_nItems
        If (StartWith) Then
            If (InStr(1, LCase(m_uList(i).Text), LCase(FindString)) = 1) Then FindFirst = i: Exit Function
          Else
            If (InStr(1, LCase(m_uList(i).Text), LCase(FindString)) > 1) Then FindFirst = i: Exit Function
        End If
    Next i
    
    '-- FindString not found
    FindFirst = -1
End Function

'-- Clear
Public Sub Clear()
    
    '-- Hide scroll bar
    sbVert.Visible = False
    sbVert.Max = 0
    
    '-- Clear and resize drawing area
    Call picList.Cls
    Call picList.Move(0, 0, ScaleWidth, ScaleHeight)
    
    '-- Reset Item arrays
    ReDim m_uList(0)
    ReDim m_bSelected(0)
    m_nItems = 0
    m_nLastItem = -1
    m_ListIndex = -1
    m_TopIndex = -1
End Sub

'-- Order
Public Sub Order()

  Dim i0     As Long
  Dim i1     As Long
  Dim i2     As Long
  Dim d      As Long
  Dim xItem  As uItem
  Dim bDesc  As Boolean
  
    If (m_nItems > 1) Then
    
        i0 = 0
        bDesc = (m_OrderType = [Descendent])
        
        If (m_SelectMode = [Single]) Then
            If (m_ListIndex > -1) Then m_bSelected(m_ListIndex) = False
        End If
        
        Do
            d = d * 3 + 1
        Loop Until d > m_nItems
        
        Do
            d = d \ 3
            For i1 = d + i0 To m_nItems + i0 - 1
            
                xItem = m_uList(i1)
                i2 = i1
                
                Do While (m_uList(i2 - d).Text > xItem.Text) Xor bDesc
                    m_uList(i2) = m_uList(i2 - d)
                    i2 = i2 - d
                    If (i2 - d < i0) Then Exit Do
                Loop
                m_uList(i2) = xItem
            Next i1
        Loop Until d = 1
        
        ListIndex = -1
        sbVert.Value = 0
        
        '-- Unselect all and refresh
        ReDim m_bSelected(0 To m_nItems)
        Call picList_Paint
    End If
End Sub

'========================================================================================
' Private
'========================================================================================

'-- pvDrawList
Private Sub pvDrawList()
    
  Dim i As Long
    
    If (Extender.Visible And UBound(m_uList())) Then
        '-- Draw visible rows
        For i = sbVert.Value To sbVert.Value + m_nVisibleRows - 1
            Call pvDrawItem(i)
        Next i
        '-- Draw focus
        Call pvDrawFocus(m_ListIndex)
    End If
End Sub

'-- pvDrawItem
Private Sub pvDrawItem(ByVal index As Integer)
   
  Dim nRctIndex As Integer
   
    '-- Item out of area?
    If (index < sbVert.Value Or index > sbVert.Value + m_nVisibleRows - 1) Then Exit Sub
    If (index > UBound(m_uList) - 1) Then Exit Sub
    
    picList.FontUnderline = False
     
    nRctIndex = index - sbVert.Value
            
    '-- Draw m_bSelected Item
    If (m_bSelected(index)) Then
    
       '-- Draw back area
        Select Case m_SelectModeStyle
            
            Case [Standard]
                Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBackSel)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                
            Case [Dither] ' Effect will be applied after drawing icon
                Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBack)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                        
            Case [Gradient_V]
                Call pvDrawBackGrad(picList.hdc, m_uRctItem(nRctIndex), m_lClrGradient1, m_lClrGradient2, GRADIENT_FILL_RECT_V)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                        
            Case [Gradient_H]
                Call pvDrawBackGrad(picList.hdc, m_uRctItem(nRctIndex), m_lClrGradient1, m_lClrGradient2, GRADIENT_FILL_RECT_H)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                                
            Case [Box]
                Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBack)
                Call pvDrawBox(picList.hdc, m_uRctItem(nRctIndex), m_BoxOffset, m_BoxRadius, m_lClrBackSel, m_lClrBox)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                                                
            Case [Underline]
                Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBack)
                Call SetTextColor(picList.hdc, m_lClrFontSel)
                picList.FontUnderline = True
                                        
            Case [ByPicture]
                If (Not SelectionPicture Is Nothing) Then
                    Call picList.PaintPicture(SelectionPicture, 0, m_uRctItem(nRctIndex).y1, m_uRctItem(nRctIndex).x2, m_nTmpItemHeight)
                  Else
                    Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBackSel)
                End If
                Call SetTextColor(picList.hdc, m_lClrFontSel)
        End Select
       
        '-- Draw icon
        If (Not m_lpoIL Is Nothing) Then
            On Error Resume Next 'Image list icon # out of bounds
            If (m_WordWrap) Then
                Call m_lpoIL.ListImages(m_uList(index).IconSelected).Draw(picList.hdc, ScaleX(m_ItemOffset, vbPixels, m_nILScale), ScaleY(m_uRctItem(nRctIndex).y1 + m_ItemOffset, vbPixels, m_nILScale), 1)
              Else
                Call m_lpoIL.ListImages(m_uList(index).IconSelected).Draw(picList.hdc, ScaleX(m_ItemOffset, vbPixels, m_nILScale), ScaleY(m_uRctItem(nRctIndex).y1 + (m_nTmpItemHeight - m_lpoIL.ImageHeight) \ 2, vbPixels, m_nILScale), 1)
            End If
            On Error GoTo 0
        End If
        
        '-- Apply dither effect (*)
        If (m_SelectModeStyle = 1) Then
            Call pvDrawDither(picList.hdc, m_uRctItem(nRctIndex), m_lClrBackSel)
        End If
     
     Else
     
        '-- Draw back area
        Call pvDrawBack(picList.hdc, m_uRctItem(nRctIndex), m_lClrBack)
        Call SetTextColor(picList.hdc, m_lClrFont)
        
        '-- Draw icon
        If (Not m_lpoIL Is Nothing) Then
            On Error Resume Next 'Image list icon # out of bounds
            If (m_WordWrap) Then
                Call m_lpoIL.ListImages(m_uList(index).Icon).Draw(picList.hdc, ScaleX(m_ItemOffset, vbPixels, m_nILScale), ScaleY(m_uRctItem(nRctIndex).y1 + m_ItemOffset, vbPixels, m_nILScale), 1)
              Else
                Call m_lpoIL.ListImages(m_uList(index).Icon).Draw(picList.hdc, ScaleX(m_ItemOffset, vbPixels, m_nILScale), ScaleY(m_uRctItem(nRctIndex).y1 + (m_nTmpItemHeight - m_lpoIL.ImageHeight) \ 2, vbPixels, m_nILScale), 1)
            End If
            On Error GoTo 0
        End If
    End If
    
    '-- Draw text...
    If (m_WordWrap) Then
        Call DrawText(picList.hdc, m_uList(index).Text, -1, m_uRctText(nRctIndex), m_Alignment Or DT_WORDBREAK)
      Else
        Call DrawText(picList.hdc, m_uList(index).Text, -1, m_uRctText(nRctIndex), DT_SINGLELINE Or DT_VCENTER)
    End If
End Sub

'-- pvDrawFocus
Private Sub pvDrawFocus(ByVal index As Integer)
    
    If (Not m_Focus Or Not m_bHasFocus) Then Exit Sub
    
    '-- Item out of area ?
    If (index < sbVert.Value Or index > sbVert.Value + m_nVisibleRows - 1) Then Exit Sub
       
    '-- Draw it
    Call SetTextColor(picList.hdc, m_lClrFont)
    Call DrawFocusRect(picList.hdc, m_uRctItem(index - sbVert.Value))
End Sub

Private Sub pvDrawBack(ByVal hdc As Long, pRect As RECT2, ByVal Color As Long)

  Dim hBrush As Long
    
    hBrush = CreateSolidBrush(Color)
    Call FillRect(hdc, pRect, hBrush)
    Call DeleteObject(hBrush)
End Sub

Private Sub pvDrawDither(ByVal hdc As Long, pRect As RECT2, ByVal Color As Long)

  Dim hBrush As Long
    
    hBrush = SelectObject(hdc, CreateSolidBrush(Color))
    Call PatBlt(hdc, pRect.x1, pRect.y1, pRect.x2 - pRect.x1, pRect.y2 - pRect.y1, &HA000C9)
    Call DeleteObject(SelectObject(hdc, hBrush))
End Sub

Private Sub pvDrawBackGrad(ByVal hdc As Long, pRect As RECT2, Color1 As RGB, Color2 As RGB, ByVal Direction As Long)

  Dim uTV(1) As TRIVERTEX
  Dim uGR    As GRADIENT_RECT
    
    '-- from
    With uTV(0)
        .x = pRect.x1
        .y = pRect.y1
        .R = Color1.R
        .G = Color1.G
        .B = Color1.B
        .Alpha = 0
    End With
    '-- to
    With uTV(1)
        .x = pRect.x2
        .y = pRect.y2
        .R = Color2.R
        .G = Color2.G
        .B = Color2.B
        .Alpha = 0
    End With
    
    uGR.UpperLeft = 0
    uGR.LowerRight = 1

    Call GradientFillRect(hdc, uTV(0), 2, uGR, 1, Direction)
End Sub

Private Sub pvDrawBox(ByVal hdc As Long, pRect As RECT2, ByVal Offset As Long, ByVal Radius As Long, ByVal ColorFill As Long, ByVal ColorBorder As Long)

  Dim hPen   As Long
  Dim hBrush As Long

    hPen = SelectObject(hdc, CreatePen(PS_SOLID, 1, ColorBorder))
    hBrush = SelectObject(hdc, CreateSolidBrush(ColorFill))
    Call InflateRect(pRect, -Offset, -Offset)
    Call RoundRect(hdc, pRect.x1, pRect.y1, pRect.x2, pRect.y2, Radius, Radius)
    Call InflateRect(pRect, Offset, Offset)
    Call DeleteObject(SelectObject(hdc, hPen))
    Call DeleteObject(SelectObject(hdc, hBrush))
End Sub

Private Sub pvReadjustScrollBar()
     
    If (m_nItems > m_nVisibleRows) Then
    
        If (Not sbVert.Visible) Then
            '-- Show scroll bar
            sbVert.Visible = True
            Call sbVert.Refresh
            sbVert.LargeChange = m_nVisibleRows
            '-- Update item rects. right margin
            Call pvRigthOffsetRects(sbVert.Width)
            '-- Repaint control area
            Call picList_Paint
        End If
      
      Else
        '-- Hide scroll bar
        sbVert.Visible = False
        '-- Update item rects. right margin
        Call pvRigthOffsetRects(0)
    End If
    
    '-- Update sbVert max value
    sbVert.Max = m_nItems - m_nVisibleRows
End Sub

Private Sub pvCalculateRects()
    
  Dim i As Long
  
    For i = 0 To m_nVisibleRows - 1
        Call SetRect(m_uRctItem(i), 0, i * m_nTmpItemHeight, ScaleWidth, i * m_nTmpItemHeight + m_nTmpItemHeight)
        Call SetRect(m_uRctText(i), m_ItemOffset + m_ItemTextLeft, i * m_nTmpItemHeight + m_ItemOffset, ScaleWidth - m_ItemOffset, i * m_nTmpItemHeight + m_nTmpItemHeight - m_ItemOffset)
        m_uPtIcon(i).x = m_ItemOffset
        m_uPtIcon(i).y = m_ItemOffset
    Next i
End Sub

Private Sub pvRigthOffsetRects(ByVal Offset As Long)
    
  Dim i As Long
  
    For i = 0 To m_nVisibleRows - 1
        m_uRctItem(i).x2 = ScaleWidth - Offset
        m_uRctText(i).x2 = ScaleWidth - m_ItemOffset - Offset
    Next i
End Sub

'-- pvScrollUp
Private Sub pvScrollUp()

  Dim t As Long ' Timer counter
  Dim d As Long ' Scrolling delay
    
    d = 500 + 20 * m_lScrollingY
    If (d < 40) Then d = 40
    
    '-- Scroll while MouseDown and mouse pos. < "Control top"
    Do While m_bScrolling And m_lScrollingY < 0
       If (GetTickCount() - t > d) Then
           t = GetTickCount()
           If (m_ListIndex > 0) Then
               If (m_SelectMode = [Multiple]) Then
                   m_bSelected(m_ListIndex - 1) = m_bAnchorSelected
               End If
               ListIndex = ListIndex - 1
           End If
       End If
       Call VBA.DoEvents
    Loop
End Sub

'-- pvScrollDown
Private Sub pvScrollDown()

  Dim t As Long ' Timer counter
  Dim d As Long ' Scrolling delay
    
    d = 500 - 20 * (m_lScrollingY - ScaleHeight - 1)
    If (d < 40) Then d = 40
    
    '-- Scroll while MouseDown and mouse pos. > "Control bottom"
    Do While m_bScrolling And m_lScrollingY > ScaleHeight - 1
       If (GetTickCount() - t > d) Then
           t = GetTickCount()
           If (m_ListIndex < m_nItems - 1) Then
               If (m_SelectMode = [Multiple]) Then
                   m_bSelected(m_ListIndex + 1) = m_bAnchorSelected
               End If
               ListIndex = ListIndex + 1
           End If
       End If
       Call VBA.DoEvents
    Loop
End Sub

'-- pvSetColors
Private Sub pvSetColors()
    
    m_lClrBack = pvGetColorLong(m_BackNormal)
    m_lClrBackSel = pvGetColorLong(m_BackSelected)
    m_lClrGradient1 = pvGetColorRGB(pvGetColorLong(m_BackSelectedG1))
    m_lClrGradient2 = pvGetColorRGB(pvGetColorLong(m_BackSelectedG2))
    m_lClrBox = pvGetColorLong(m_BoxBorder)
    m_lClrFont = pvGetColorLong(m_FontNormal)
    m_lClrFontSel = pvGetColorLong(m_FontSelected)
End Sub

Private Function pvGetColorLong(Color As Long) As Long
    
    If (Color And &H80000000) Then
        pvGetColorLong = GetSysColor(Color And &H7FFFFFFF)
      Else
        pvGetColorLong = Color
    End If
End Function

Private Function pvGetColorRGB(Color As Long) As RGB

  Dim HexColor As String
        
    HexColor = String(6 - Len(Hex(Color)), "0") & Hex(Color)
    pvGetColorRGB.R = "&H" & Mid(HexColor, 5, 2) & "00"
    pvGetColorRGB.G = "&H" & Mid(HexColor, 3, 2) & "00"
    pvGetColorRGB.B = "&H" & Mid(HexColor, 1, 2) & "00"
End Function

'========================================================================================
' Properties
'========================================================================================

'-- Alignment
Public Property Get Alignment() As AlignmentCts
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As AlignmentCts)
    m_Alignment = New_Alignment
    Call picList_Paint
End Property

'-- Appearance
Public Property Get Appearance() As AppearanceCts
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)
    UserControl.Appearance() = New_Appearance
End Property

'-- BackNormal
Public Property Get BackNormal() As OLE_COLOR
    BackNormal = m_BackNormal
End Property

Public Property Let BackNormal(ByVal New_BackNormal As OLE_COLOR)
    m_BackNormal = New_BackNormal
    m_lClrBack = pvGetColorLong(m_BackNormal)
    picList.BackColor = m_lClrBack
    Call picList_Paint
End Property

'-- BackSelected
Public Property Get BackSelected() As OLE_COLOR
    BackSelected = m_BackSelected
End Property

Public Property Let BackSelected(ByVal New_BackSelected As OLE_COLOR)
    m_BackSelected = New_BackSelected
    m_lClrBackSel = pvGetColorLong(m_BackSelected)
    Call picList_Paint
End Property

'-- BackSelectedG1
Public Property Get BackSelectedG1() As OLE_COLOR
    BackSelectedG1 = m_BackSelectedG1
End Property

Public Property Let BackSelectedG1(ByVal New_BackSelectedG1 As OLE_COLOR)
    m_BackSelectedG1 = New_BackSelectedG1
    m_lClrGradient1 = pvGetColorRGB(pvGetColorLong(m_BackSelectedG1))
    Call picList_Paint
End Property

'-- BackSelectedG2
Public Property Get BackSelectedG2() As OLE_COLOR
    BackSelectedG2 = m_BackSelectedG2
End Property

Public Property Let BackSelectedG2(ByVal New_BackSelectedG2 As OLE_COLOR)
    m_BackSelectedG2 = New_BackSelectedG2
    m_lClrGradient2 = pvGetColorRGB(pvGetColorLong(m_BackSelectedG2))
    Call picList_Paint
End Property

'-- BorderStyle
Public Property Get BorderStyle() As BorderStyleCts
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)
    UserControl.BorderStyle() = New_BorderStyle
End Property

'-- BoxBorder
Public Property Get BoxBorder() As OLE_COLOR
    BoxBorder = m_BoxBorder
End Property

Public Property Let BoxBorder(ByVal New_BoxBorder As OLE_COLOR)
    m_BoxBorder = New_BoxBorder
    m_lClrBox = pvGetColorLong(m_BoxBorder)
    Call picList_Paint
End Property

'-- BoxOffset
Public Property Get BoxOffset() As Integer
    BoxOffset = m_BoxOffset
End Property

Public Property Let BoxOffset(ByVal New_BoxOffset As Integer)
    If (New_BoxOffset <= m_nTmpItemHeight \ 2) Then
        m_BoxOffset = New_BoxOffset
    End If
    Call picList_Paint
End Property

'-- BoxRadius
Public Property Get BoxRadius() As Integer
    BoxRadius = m_BoxRadius
End Property

Public Property Let BoxRadius(ByVal New_BoxRadius As Integer)
    m_BoxRadius = New_BoxRadius
    Call picList_Paint
End Property

'-- Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    sbVert.Enabled = New_Enabled
End Property

'-- Focus
Public Property Get Focus() As Boolean
    Focus = m_Focus
End Property

Public Property Let Focus(ByVal New_Focus As Boolean)
    m_Focus = New_Focus
    If (New_Focus) Then
        Call pvDrawFocus(m_ListIndex)
      Else
        Call pvDrawItem(m_ListIndex)
    End If
End Property

'-- Font
Public Property Get Font() As Font
    Set Font = m_oFont
End Property

Public Property Set Font(ByVal New_Font As Font)
    With m_oFont
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
    End With
    Call picList_Paint
End Property

Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set picList.Font = m_oFont
    Call UserControl_Resize
End Sub

'-- FontNormal
Public Property Get FontNormal() As OLE_COLOR
    FontNormal = m_FontNormal
End Property

Public Property Let FontNormal(ByVal New_FontNormal As OLE_COLOR)
    m_FontNormal = New_FontNormal
    m_lClrFont = pvGetColorLong(m_FontNormal)
    Call SetTextColor(picList.hdc, m_lClrFont)
    Call picList_Paint
End Property

'-- FontSelected
Public Property Get FontSelected() As OLE_COLOR
    FontSelected = m_FontSelected
End Property

Public Property Let FontSelected(ByVal New_FontSelected As OLE_COLOR)
    m_FontSelected = New_FontSelected
    m_lClrFontSel = pvGetColorLong(m_FontSelected)
    Call picList_Paint
End Property

'-- HoverSelection
Public Property Get HoverSelection() As Boolean
    HoverSelection = m_HoverSelection
End Property

Public Property Let HoverSelection(ByVal New_HoverSelection As Boolean)
    m_HoverSelection = New_HoverSelection
    Call pvDrawItem(m_ListIndex)
    Call pvDrawFocus(m_ListIndex)
End Property

'-- ItemHeight
Public Property Get ItemHeight() As Integer
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)
    m_ItemHeight = New_ItemHeight
    Call UserControl_Resize
    Call picList_Paint
End Property

'-- ItemHeightAuto
Public Property Get ItemHeightAuto() As Boolean
    ItemHeightAuto = m_ItemHeightAuto
End Property

Public Property Let ItemHeightAuto(ByVal New_ItemHeightAuto As Boolean)
    m_ItemHeightAuto = New_ItemHeightAuto
    Call UserControl_Resize
    Call picList_Paint
End Property

'-- ItemOffset
Public Property Get ItemOffset() As Integer
    ItemOffset = m_ItemOffset
End Property

Public Property Let ItemOffset(ByVal New_ItemOffset As Integer)
    If (New_ItemOffset <= m_nTmpItemHeight) Then
        m_ItemOffset = New_ItemOffset
    End If
    Call pvCalculateRects
    If (sbVert.Visible) Then
        Call pvRigthOffsetRects(sbVert.Width)
    End If
    Call picList_Paint
End Property

'-- ItemTextLeft
Public Property Get ItemTextLeft() As Integer
    ItemTextLeft = m_ItemTextLeft
End Property

Public Property Let ItemTextLeft(ByVal New_ItemTextLeft As Integer)
    m_ItemTextLeft = New_ItemTextLeft
    Call pvCalculateRects
    If (sbVert.Visible) Then
        Call pvRigthOffsetRects(sbVert.Width)
    End If
    Call picList_Paint
End Property

'-- <ListCount>
Public Property Get ListCount() As Integer
    ListCount = m_nItems
End Property

'-- ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_MemberFlags = "400"
    ListIndex = m_ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    
    If (New_ListIndex < -1 Or New_ListIndex > m_nItems - 1) Then Err.Raise 380
    
    If (txtEdit.Visible) Then
        Call txtEdit_LostFocus
    End If
    
    If (New_ListIndex < 0 Or m_nItems = 0) Then
        m_ListIndex = -1
        m_snLastY = -1
      Else
        m_ListIndex = New_ListIndex
    End If
    
    '-- Unselect last / Select actual [Single selection mode]
    If (m_SelectMode = [Single]) Then
        If (m_nLastItem > -1) Then m_bSelected(m_nLastItem) = False
        If (m_ListIndex > -1) Then m_bSelected(m_ListIndex) = True
    End If

    '-- Draw last (delete Focus) ...
    Call pvDrawItem(m_nLastItem)
    m_nLastItem = m_ListIndex
    
    '-- ... and draw actual (draw Focus)
    Call pvDrawItem(m_ListIndex)
    Call pvDrawFocus(m_ListIndex)

    '-- Ensure visible actual Selected item
    If (m_bEnsureVisible) Then
        If (m_ListIndex < sbVert.Value And m_ListIndex > -1) Then
            sbVert.Value = m_ListIndex
          ElseIf (m_ListIndex > sbVert.Value + m_nVisibleRows - 1) Then
            sbVert.Value = m_ListIndex - m_nVisibleRows + 1
        End If
      Else
        m_bEnsureVisible = True
    End If
    
    RaiseEvent ListIndexChange
End Property

'-- MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = picList.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set picList.MouseIcon = New_MouseIcon
End Property

'-- MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = picList.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    picList.MousePointer() = New_MousePointer
End Property

'-- OrderType
Public Property Get OrderType() As OrderTypeCts
    OrderType = m_OrderType
End Property

Public Property Let OrderType(ByVal New_OrderType As OrderTypeCts)
    m_OrderType = New_OrderType
End Property

'-- ScrollBarWidth
Public Property Get ScrollBarWidth() As Integer
    ScrollBarWidth = m_ScrollBarWidth
End Property

Public Property Let ScrollBarWidth(ByVal New_ScrollBarWidth As Integer)
    
    '-- Check Min value width...
    If (New_ScrollBarWidth < 9) Then
        m_ScrollBarWidth = 9
        sbVert.Width = 9
    '-- Check Max value width...
      ElseIf (New_ScrollBarWidth > ScaleWidth \ 2) Then
        m_ScrollBarWidth = ScaleWidth \ 2
        sbVert.Width = ScaleWidth \ 2
    '-- Set new value...
      Else
        m_ScrollBarWidth = New_ScrollBarWidth
        sbVert.Width = New_ScrollBarWidth
    End If
    
    sbVert.Visible = False
    Call pvReadjustScrollBar
    Call UserControl_Resize
End Property

'-- <SelectedCount>
Public Property Get SelectedCount() As Integer
    
  Dim i As Long
    
    SelectedCount = 0
    For i = 0 To m_nItems
        If (m_bSelected(i)) Then SelectedCount = SelectedCount + 1
    Next i
End Property

'-- SelectionPicture
Public Property Get SelectionPicture() As Picture
    Set SelectionPicture = m_SelectionPicture
End Property

Public Property Set SelectionPicture(ByVal New_SelectionPicture As Picture)
    Set m_SelectionPicture = New_SelectionPicture
    Call picList_Paint
End Property

'-- SelectMode
Public Property Get SelectMode() As SelectModeCts
    SelectMode = m_SelectMode
End Property

Public Property Let SelectMode(ByVal New_SelectMode As SelectModeCts)
    
  Dim i As Long
  
    m_SelectMode = New_SelectMode
    
    If (Ambient.UserMode) Then
        If (New_SelectMode = [Single]) Then
            '-- Unselect all and select actual
            If (m_ListIndex > -1) Then
                For i = LBound(m_uList()) To m_nItems
                    If (i <> m_ListIndex) Then m_bSelected(i) = False
                Next i
                m_bSelected(m_ListIndex) = True
                Call pvDrawItem(m_ListIndex)
                Call pvDrawFocus(m_ListIndex)
            End If
        End If
   End If
   
   Call pvReadjustScrollBar
   Call picList_Paint
End Property

'-- SelectModeStyle
Public Property Get SelectModeStyle() As SelectModeStyleCts
    SelectModeStyle = m_SelectModeStyle
End Property

Public Property Let SelectModeStyle(ByVal New_SelectModeStyle As SelectModeStyleCts)
    m_SelectModeStyle = New_SelectModeStyle
    Call picList_Paint
End Property

'-- TopIndex
Public Property Get TopIndex() As Integer
Attribute TopIndex.VB_MemberFlags = "400"
    TopIndex = sbVert.Value
End Property

Public Property Let TopIndex(ByVal New_TopIndex As Integer)
    
    If (New_TopIndex < 0 Or New_TopIndex > m_nItems - m_nVisibleRows) Then Err.Raise 380

    m_TopIndex = New_TopIndex
    sbVert.Value = New_TopIndex
    
    RaiseEvent TopIndexChange
End Property

'-- WordWrap
Public Property Get WordWrap() As Boolean
    WordWrap = m_WordWrap
End Property

Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    m_WordWrap = New_WordWrap
    Call picList_Paint
End Property

Private Sub UserControl_InitProperties()

    UserControl.Appearance = m_def_Appearance
    UserControl.BorderStyle = m_def_BorderStyle
    m_ScrollBarWidth = m_def_ScrollBarWidth

    Set picList.Font = Ambient.Font
    Set m_oFont = Ambient.Font
    
    m_FontNormal = m_def_FontNormal
    m_FontSelected = m_def_FontSelected
    m_BackNormal = m_def_BackNormal
    m_BackSelected = m_def_BackSelected
    m_BackSelectedG1 = m_def_BackSelectedG1
    m_BackSelectedG2 = m_def_BackSelectedG2
    
    m_BoxBorder = m_def_BoxBorder
    m_BoxOffset = m_def_BoxOffset
    m_BoxRadius = m_def_BoxRadius
    
    m_Alignment = m_def_Alignment
    m_Focus = m_def_Focus
    m_HoverSelection = m_def_HoverSelection
    m_WordWrap = m_def_WordWrap
    
    m_ItemHeight = picList.TextHeight(vbNullString)
    m_ItemHeightAuto = m_def_ItemHeightAuto
    m_ItemOffset = m_def_ItemOffset
    m_ItemTextLeft = m_def_ItemTextLeft
    
    m_OrderType = m_def_OrderType
    Set m_SelectionPicture = Nothing
    m_SelectMode = m_def_SelectMode
    m_SelectModeStyle = m_def_SelectModeStyle
    
    m_ListIndex = -1
    m_TopIndex = -1
    
    Call pvSetColors
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  Dim sTmp As String
    
    With PropBag
    
        UserControl.Appearance = .ReadProperty("Appearance", m_def_Appearance)
        UserControl.BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        UserControl.Enabled = .ReadProperty("Enabled", -1)
        m_ScrollBarWidth = .ReadProperty("ScrollBarWidth", m_def_ScrollBarWidth)
        sbVert.Width = .ReadProperty("ScrollBarWidth", m_def_ScrollBarWidth)
        
        Set picList.Font = .ReadProperty("Font", Ambient.Font)
        Set m_oFont = .ReadProperty("Font", Ambient.Font)
        
        m_FontNormal = .ReadProperty("FontNormal", m_def_FontNormal)
        m_FontSelected = .ReadProperty("FontSelected", m_def_FontSelected)
        m_BackNormal = .ReadProperty("BackNormal", m_def_BackNormal)
        picList.BackColor = .ReadProperty("BackNormal", m_def_BackNormal)
        m_BackSelected = .ReadProperty("BackSelected", m_def_BackSelected)
        m_BackSelectedG1 = .ReadProperty("BackSelectedG1", m_def_BackSelectedG1)
        m_BackSelectedG2 = .ReadProperty("BackSelectedG2", m_def_BackSelectedG2)
        
        m_BoxBorder = .ReadProperty("BoxBorder", m_def_BoxBorder)
        m_BoxOffset = .ReadProperty("BoxOffset", m_def_BoxOffset)
        m_BoxRadius = .ReadProperty("BoxRadius", m_def_BoxRadius)
        
        m_Alignment = .ReadProperty("Alignment", m_def_Alignment)
        m_Focus = .ReadProperty("Focus", m_def_Focus)
        m_HoverSelection = .ReadProperty("HoverSelection", m_def_HoverSelection)
        m_WordWrap = .ReadProperty("WordWrap", m_def_WordWrap)
    
        m_ItemOffset = .ReadProperty("ItemOffset", m_def_ItemOffset)
        m_ItemHeightAuto = .ReadProperty("ItemHeightAuto", m_def_ItemHeightAuto)
        m_ItemTextLeft = .ReadProperty("ItemTextLeft", m_def_ItemTextLeft)
    
        m_OrderType = .ReadProperty("OrderType", m_def_OrderType)
        Set m_SelectionPicture = .ReadProperty("SelectionPicture", Nothing)
        m_SelectMode = .ReadProperty("SelectMode", m_def_SelectMode)
        m_SelectModeStyle = .ReadProperty("SelectModeStyle", m_def_SelectModeStyle)
        
        picList.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set picList.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        
        sTmp = .ReadProperty("ItemHeight", 0)
        If (sTmp < picList.TextHeight(vbNullString)) Then
            m_ItemHeight = picList.TextHeight(vbNullString)
          Else
            m_ItemHeight = sTmp
        End If
    End With
    
    m_ListIndex = -1
    m_TopIndex = -1

    Call pvSetColors
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
    
        Call .WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
        Call .WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("ScrollBarWidth", m_ScrollBarWidth, m_def_ScrollBarWidth)
        
        Call .WriteProperty("Font", picList.Font, Ambient.Font)
        Call .WriteProperty("FontNormal", m_FontNormal, m_def_FontNormal)
        Call .WriteProperty("FontSelected", m_FontSelected, m_def_FontSelected)
        Call .WriteProperty("BackNormal", m_BackNormal, m_def_BackNormal)
        Call .WriteProperty("BackSelected", m_BackSelected, m_def_BackSelected)
        Call .WriteProperty("BackSelectedG1", m_BackSelectedG1, m_def_BackSelectedG1)
        Call .WriteProperty("BackSelectedG2", m_BackSelectedG2, m_def_BackSelectedG2)
        
        Call .WriteProperty("BoxBorder", m_BoxBorder, m_def_BoxBorder)
        Call .WriteProperty("BoxOffset", m_BoxOffset, m_def_BoxOffset)
        Call .WriteProperty("BoxRadius", m_BoxRadius, m_def_BoxRadius)
        
        Call .WriteProperty("Alignment", m_Alignment, m_def_Alignment)
        Call .WriteProperty("Focus", m_Focus, m_def_Focus)
        Call .WriteProperty("HoverSelection", m_HoverSelection, m_def_HoverSelection)
        Call .WriteProperty("WordWrap", m_WordWrap, m_def_WordWrap)
        
        Call .WriteProperty("ItemHeight", m_ItemHeight, 0)
        Call .WriteProperty("ItemHeightAuto", m_ItemHeightAuto, m_def_ItemHeightAuto)
        Call .WriteProperty("ItemOffset", m_ItemOffset, m_def_ItemOffset)
        Call .WriteProperty("ItemTextLeft", m_ItemTextLeft, m_def_ItemTextLeft)
        
        Call .WriteProperty("OrderType", m_OrderType, m_def_OrderType)
        Call .WriteProperty("SelectionPicture", m_SelectionPicture, Nothing)
        Call .WriteProperty("SelectMode", m_SelectMode, m_def_SelectMode)
        Call .WriteProperty("SelectModeStyle", m_SelectModeStyle, m_def_SelectModeStyle)
    
        Call .WriteProperty("MousePointer", picList.MousePointer, vbDefault)
        Call .WriteProperty("MouseIcon", picList.MouseIcon, Nothing)
    End With
End Sub



'Last revised: 02/07/02
'-------------------------------------------------------------------------------------------
' Some methods passed to R/W properties:
'
' GetItem i    GetIcon i    GetIconSelected i    IsSelected i
' to           to           to                   to
' ItemText(i)  ItemIcon(i)  ItemIconSelected(i)  ItemSelected(i)
'
' Or use ModifyItem to change all item parameters at time


'-- ItemText
Public Property Get ItemText(ByVal index As Integer) As String
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    ItemText = m_uList(index).Text
End Property

Public Property Let ItemText(ByVal index As Integer, ByVal Data As String)
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    m_uList(index).Text = CStr(Data)
    Call pvDrawItem(index)
    Call pvDrawFocus(m_ListIndex)
End Property

'-- ItemIcon
Public Property Get ItemIcon(ByVal index As Integer) As Integer
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    ItemIcon = m_uList(index).Icon
End Property

Public Property Let ItemIcon(ByVal index As Integer, ByVal Data As Integer)
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    m_uList(index).Icon = Data
    Call pvDrawItem(index)
    Call pvDrawFocus(m_ListIndex)
End Property

'-- ItemIconSelected
Public Property Get ItemIconSelected(ByVal index As Integer) As Integer
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    ItemIconSelected = m_uList(index).IconSelected
End Property

Public Property Let ItemIconSelected(ByVal index As Integer, ByVal Data As Integer)
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    m_uList(index).IconSelected = Data
    Call pvDrawItem(index)
    Call pvDrawFocus(m_ListIndex)
End Property

'-- ItemSelected
Public Property Get ItemSelected(ByVal index As Integer) As Boolean
    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    ItemSelected = m_bSelected(index)
End Property

Public Property Let ItemSelected(ByVal index As Integer, ByVal Data As Boolean)

    If (m_nItems = 0 Or index > m_nItems) Then Err.Raise 381
    
    If (Data) Then
        If (m_SelectMode = [Single]) Then
            ListIndex = index
          Else
            m_bSelected(index) = True
            Call pvDrawItem(index)
            If (index = m_ListIndex) Then
                Call pvDrawFocus(index)
            End If
        End If
      Else
        If (m_SelectMode = [Single]) Then
          Else
            m_bSelected(index) = False
            Call pvDrawItem(index)
            If (index = m_ListIndex) Then
                Call pvDrawFocus(index)
            End If
        End If
    End If
End Property

'Editing item...
'-------------------------------------------------------------------------------------------
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        
    ' WordWrap mode enabled:
    ' [Control]+[Return] = new line
    ' [Return]           = update text
    ' WordWrap mode disabled:
    ' [Return]           = update text
    
    '-- Enabled new line in WordWrap mode
    If (m_WordWrap) Then
        If (KeyAscii = vbKeyReturn) Then
            m_uList(m_ListIndex).Text = txtEdit.Text
            Call txtEdit_LostFocus
        End If
    '-- Don't allow new line in disabled WordWrap mode
      Else
        If (KeyAscii = vbKeyReturn Or KeyAscii = 10) Then
            m_uList(m_ListIndex).Text = txtEdit.Text
            Call txtEdit_LostFocus
        End If
    End If
    '-- Cancel edition
    If (KeyAscii = vbKeyEscape) Then
        Call txtEdit_LostFocus
    End If
End Sub

Private Sub txtEdit_LostFocus()
       
    '-- Hide edit TextBox and let ListBox keyboard control
    txtEdit.Visible = False
    UserControl.KeyPreview = True
End Sub

Public Sub StartEdit()
        
    '-- Item is selected...
    If (m_ListIndex > -1) Then
     
        '-- Let TextBox keyboard control
        UserControl.KeyPreview = False
        
        With txtEdit
            '-- Get TextBox item font properties
            Set .Font = m_oFont
            If (m_bSelected(m_ListIndex) And m_SelectModeStyle <> [Underline]) Then
                .BackColor = m_lClrBackSel
                .ForeColor = m_lClrFontSel
              Else
                .BackColor = m_lClrBack
                .ForeColor = m_lClrFont
            End If
                
            '-- Set alignment. Locate and resize TextBox
            If (m_WordWrap) Then
                .Alignment = Choose(m_Alignment + 1, [AlignLeft], [AlignRight], [AlignCenter])
                Call .Move(m_ItemTextLeft + m_ItemOffset, (m_ListIndex - sbVert.Value) * m_nTmpItemHeight + m_ItemOffset, m_uRctItem(m_ListIndex - sbVert.Value).x2 - m_ItemTextLeft - 2 * m_ItemOffset, m_nTmpItemHeight - 2 * m_ItemOffset)
              Else
                .Alignment = [AlignLeft]
                Call .Move(m_ItemTextLeft + m_ItemOffset, (m_ListIndex - sbVert.Value) * m_nTmpItemHeight + 0.5 * (m_nTmpItemHeight - picList.TextHeight(vbNullString)), m_uRctItem(m_ListIndex - sbVert.Value).x2 - m_ItemTextLeft - 2 * m_ItemOffset, 1)
            End If
             
            '-- Get item text and turn TextBox to visible
            .Text = m_uList(m_ListIndex).Text
            .SelStart = 0
            .SelLength = Len(txtEdit.Text)
            .Visible = True
            Call .SetFocus
        End With
    End If
End Sub

Public Sub EndEdit(Optional ByVal Modify As Boolean = False)
    If (Modify) Then
        Call txtEdit_KeyPress(vbKeyReturn)
      Else
        Call txtEdit_LostFocus
    End If
End Sub
