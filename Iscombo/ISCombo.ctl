VERSION 5.00
Begin VB.UserControl ISCombo 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3660
   ForwardFocus    =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   3660
   ToolboxBitmap   =   "ISCombo.ctx":0000
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   2760
      Width           =   1875
   End
   Begin VB.Image picbutton 
      Height          =   540
      Left            =   -210
      Picture         =   "ISCombo.ctx":0312
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   360
   End
   Begin VB.Image imgItem 
      Height          =   195
      Left            =   240
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   375
   End
End
Attribute VB_Name = "ISCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''      Enhanced by:    Jarry Claessen
''      email:          jacsoft@wishmail.net

''
''      Description:
''
''      I've Got a lot of problemas with the VB' combo, I couldn't detect
''      when the user changes the selection, and, those combos are relly ugly :(
''      so, I decided made one better.
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''

''      Changes made by Jarry Claessen: The old version used to much
''      of data, it created as much labels as it would have items. I have changed
''      it and now it only has at maximum 8 labels. Furthermore I made function
''      to let the user now wich item is selected (with selecteditem) and let
''      the user get the data back from the combo box, (pictures to). I made the
''      height of the control better and made it able to use keys in the iscombo
''      Next to that I also speeded it up a little.
''      And now when the users presses it twice the iscombo will be closed.
''      In other words, much improvements.


''      The combobox is for autocompleting the text so don't erase


Option Explicit

Private HowManyPicturesHaveWeGot As Integer
Private ShowPictureNow As Boolean
Private MayChangeDo As Boolean

Private hdown As Form

Private dontshow As Boolean
' Type Declarations
Private Type PointAPI
        x As Long
        Y As Long
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Enum State
    Normal
    Hover
    pushed
End Enum

Private selitem As Integer

Private InOut As Boolean
Private iState As State
Private OnClicking As Boolean
Private OnFocus As Boolean

Private gScaleX As Single '= Screen.TwipsPerPixelX
Private gScaleY As Single '= Screen.TwipsPerPixelY

Private WithEvents cdown As wndDown
Attribute cdown.VB_VarHelpID = -1

'Default Property Values:
Public Enum ShowPictures
    Auto
    Always
    Never
End Enum

Public Enum AlignMent
    Align_left
    Align_right
    Align_center
End Enum

Const m_def_Enabled = True
Const m_def_FontColor = 0
Const m_def_FontHighlightColor = 0
Const m_def_IconAlign = 0
Const m_def_IconSize = 0
Const m_def_TextAlign = 4
Const m_def_BackColor = &HE0E0E0
Const m_def_HoverColor = &HFFF0B8

'Property Variables:
Dim m_Enabled As Boolean
Dim m_FontColor As OLE_COLOR
Dim m_FontHighlightColor As OLE_COLOR
Dim m_IconSize As Integer

Dim m_TextAlign As AlignMent

Dim m_HoverIcon As Picture
Dim m_BackColor As OLE_COLOR
Dim m_HoverColor As OLE_COLOR
Dim m_Default As Boolean
Dim m_Focused As Boolean
Dim m_ImageSize As Integer
Dim m_Items As New Collection
Dim m_Images As New Collection
Dim m_ItemsCount As Integer

Dim m_AutoComplete As Boolean
Dim m_Sorted As Boolean
Dim m_Showpictures As ShowPictures
Dim m_Locked As Boolean
Dim m_picture As New StdPicture

Dim m_CaseSensitiveAcceptDoubles As Boolean

Dim m_max As Integer
Dim m_AcceptDoubles

Public Enum MaxExceeded
    OverWriteFirst  'remove first
    OverWriteLast   'remove last
    OverWriteRandom 'remove random
    DontAdd         'dontadd last
End Enum

Dim m_MaxExceeded As MaxExceeded

Dim sh_Images As New Collection
Dim sh_Items As New Collection
'Event Declarations:

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseHover(Index As Integer)
Event KeyPress(KeyAscii As Integer)
Event ButtonClick()
Event ItemClick(iItem As Integer)
Event Change()
Const pBorderColor = &HC08080
' API Declarations

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long



Private Sub DrawFlat()
    If OnFocus Then
        DrawFace 4
    Else
        DrawFace 0
    End If
End Sub

Private Sub DrawRaised()
    DrawFace 3
End Sub

Private Sub DrawPushed()
    DrawFace 1
End Sub

Public Function SelectedItem() As Integer
SelectedItem = selitem
End Function

Public Function ListItem(nr As Integer) As String
On Error Resume Next
ListItem = m_Items(nr + 1)
End Function
Public Function ListImage(nr As Integer) As Picture
Set ListImage = m_Images(nr + 1)
End Function

Private Sub cDown_ItemClick(iItem As Integer, sText As String, DontSay As Boolean)
    On Error Resume Next
    imgItem.Picture = m_Images(iItem + 1)
    Set m_picture = Nothing
    txtText.text = sText
    txtText.SelStart = 0
    txtText.SelLength = Len(sText)
    selitem = iItem
    If Not DontSay Then RaiseEvent ItemClick(iItem + 1)
End Sub
Private Sub cDown_MouseHover(Index As Integer)
RaiseEvent MouseHover(Index)
End Sub

Private Sub cDown_unloaded()
On Error Resume Next
txtText.SetFocus
End Sub

Private Sub imgItem_Click()
    If m_Enabled Then RaiseEvent Click
End Sub

Private Sub imgItem_DblClick()
If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    OnClicking = True
    Dim showing As Boolean
    On Error Resume Next
    showing = hdown.Visible
    If showing Then hdown.Visible = False: dontshow = True: Exit Sub
    'If Button = vbLeftButton Then
        'picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vbWindowBackground
        'picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vbWindowBackground
        'picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DShadow
        'picButton.Line (1, 0)-(picButton.ScaleWidth - 1, 0), vb3DShadow
    'End If
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Not OnClicking Then
        UserControl_MouseMove Button, Shift, x, Y
    End If
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    OnClicking = False
End Sub

Private Sub txtText_Change()
    
    RaiseEvent Change
        
    If MayChangeDo Then

    If m_AutoComplete Then
            'try to find something suitable
            Dim strTemp As String
            'Figure out the string prefix to search
            '     for
            'If txtText.SelStart = 0 Then
                strTemp = txtText.text '& Chr(KeyAscii)
            'Else
            '    strTemp = Left$(txtText.text, txtText.SelStart) & Chr(KeyAscii)
            'End If
                
                'This could really be speeded up, but it hadn't got time
                'enough to think good about it,so.....
                'maybe in a next version, but until now it works though
                
                Dim x As Integer
                
                For x = 1 To m_Items.Count
                    If Left$(UCase$(m_Items(x)), Len(strTemp)) = UCase$(strTemp) Then
                    'found one add it
                    cDown_ItemClick x - 1, m_Items(x), False
                    txtText.SelStart = Len(strTemp)
                    txtText.SelLength = Len(txtText.text) - UserControl.txtText.SelStart
                    Exit For
                    End If
                Next
            
            
        End If
    End If
    MayChangeDo = True
End Sub

Private Sub txtText_Click()
If m_Enabled Then RaiseEvent Click
End Sub

Private Sub txtText_DblClick()
If m_Enabled Then RaiseEvent DblClick
End Sub

Private Sub txtText_GotFocus()
    txtText.SelStart = 0
    txtText.SelLength = Len(txtText.text)
End Sub

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
picButton_Click
End If

If m_Enabled Then
    RaiseEvent KeyPress(KeyCode)
End If
'Debug.Print KeyCode
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    
If m_Enabled Then
    RaiseEvent KeyPress(KeyAscii)
End If
    
    If m_Locked Then
        
        KeyAscii = 0
    
    Else
        
        If KeyAscii = vbKeyDelete Or KeyAscii = 8 Then MayChangeDo = False
    
    End If
End Sub

Private Sub txtText_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UserControl_MouseMove Button, Shift, x, Y
End Sub

Private Sub UserControl_Click()
    If m_Enabled Then RaiseEvent Click
End Sub

Public Function ItemCount() As Integer
ItemCount = -1
ItemCount = m_Items.Count
End Function


Private Sub UserControl_Initialize()
    gScaleX = Screen.TwipsPerPixelX
    gScaleY = Screen.TwipsPerPixelY
    m_ImageSize = 16 * gScaleX
    'Set ImgIcon.Picture = LoadPicture()
    MayChangeDo = True
    UserControl_Resize
    AutoComplete = True
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then DrawFace 1
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If m_Enabled Then
        RaiseEvent Click
        RaiseEvent KeyPress(KeyAscii)
    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If m_Enabled Then DrawFace 0
    If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
'    picButton_MouseDown vbLeftButton, Shift
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If m_Enabled Then
        iState = pushed
        UserControl_Paint
        OnClicking = True
        RaiseEvent MouseDown(Button, Shift, x, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent MouseMove(Button, Shift, x, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If m_Enabled Then
        iState = Hover
        UserControl_Paint
        RaiseEvent MouseUp(Button, Shift, x, Y)
    End If
End Sub

Private Sub UserControl_Resize()
    '   Text Position
    On Error Resume Next
    UserControl.ScaleMode = 1
    
    If Width < 700 Then Width = 700

    ChangeShowImages True

    If ShowPictureNow Then
    
        imgItem.Visible = True
        imgItem.Move 75, (Height - m_ImageSize) / 2, m_ImageSize, m_ImageSize
        txtText.Move 105 + m_ImageSize, (Height - m_ImageSize) / 2, Width - m_ImageSize - 150
    
    Else
        imgItem.Visible = False
        txtText.Move 105, (Height - m_ImageSize) / 2, Width - m_ImageSize - 150
    End If
    
    Select Case m_TextAlign
        Case Align_left  '   Left
            txtText.AlignMent = 0
        Case Align_right '   Right
            txtText.AlignMent = 1
        Case Align_center  '   Center
            txtText.AlignMent = 2
    End Select
    'Locate Button
    picbutton.Move Width - 300, 15, 230, Height - 60
    txtText.Width = picbutton.Left - txtText.Left
    
'    imgItem.Refresh
    If Not (m_picture Is Nothing) Then
    Set imgItem.Picture = m_picture
    End If
    
    
    'imgDown.Move (picButton.ScaleWidth - imgDown.Width) / 2, (picButton.ScaleHeight - imgDown.Height) / 2
End Sub

Private Sub UserControl_Paint()
    '
    
    If m_Enabled Then
        Select Case iState
            Case Hover
                DrawFace 4
            Case pushed
                DrawFace 1
            Case Normal
                DrawFace 0
        End Select
    Else
        DrawFace 2
    End If
    
'    imgItem.Refresh
'    If Not (m_picture Is Nothing) Then
'    Set imgItem.Picture = m_picture
'    End If

End Sub

Private Sub UserControl_DblClick()
    If m_Enabled Then RaiseEvent DblClick
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txttext,txttext,-1,Caption
Public Property Get Caption() As String
    Caption = UserControl.txtText.text
End Property

Public Property Let Caption(ByVal New_Caption As String)
    txtText.text() = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get MaximumItems() As Integer
    MaximumItems = m_max
End Property

Public Property Let MaximumItems(ByVal New_max As Integer)
    m_max = New_max
    PropertyChanged "MaximumItems"
    changingnumbers
End Property

Public Property Get AcceptDoubles() As Boolean
    AcceptDoubles = m_AcceptDoubles
End Property

Public Property Let AcceptDoubles(ByVal New_ac As Boolean)
    m_AcceptDoubles = New_ac
    PropertyChanged "AcceptDoubles"
End Property

Public Property Get MaxExceeded() As MaxExceeded
    MaxExceeded = m_MaxExceeded
End Property

Public Property Let MaxExceeded(ByVal New_MaxExceeded As MaxExceeded)
    m_MaxExceeded = New_MaxExceeded
    PropertyChanged "MaxExceeded"
End Property

Private Sub changingnumbers()
Dim r As Integer, p As Integer, z As Integer, x As Integer
If m_max < m_Items.Count Then


    Select Case m_MaxExceeded
        
        Case OverWriteFirst  'remove first
        
        For x = 1 To m_Items.Count - m_max
        p = 1
        GoSub removefromsh
        m_Items.Remove 1
        m_Images.Remove 1
        Next
        
        Case OverWriteLast   'remove last
               
        For x = m_Items.Count To m_max + 1
        p = x
        GoSub removefromsh
        m_Items.Remove x
        m_Images.Remove x
        Next
               
               
        Case OverWriteRandom 'remove random
         
        Randomize Timer
        For x = 1 To m_Items.Count - m_max
        r = (Rnd * (m_Items.Count - 1)) + 1
        
        p = r
        GoSub removefromsh
        
        m_Items.Remove r
        m_Images.Remove r
        Next
        
        
        Case DontAdd    'do nothing so remove the lasts
        
        For x = m_Items.Count To m_max + 1
        m_Items.Remove x
        m_Images.Remove x
        Next
        
        
    End Select

End If

Exit Sub

removefromsh:

z = -1
        Do
        z = SearchNumberByTextsh(m_Items(p), z + 1)
            
            If z <> -1 Then
                
                Dim str1 As String, str2 As String
                On Error Resume Next
                str1 = Str(m_Images(p))
                str2 = Str(sh_Images(z))
                
                If sh_Items(z) = m_Items(p) And Trim(str1) = Trim(str2) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                Exit Do
                End If
            End If
        
        Loop Until z = -1
        
Return
End Sub

Public Property Get CaseSensitiveAcceptDoubles() As Boolean
    CaseSensitiveAcceptDoubles = m_CaseSensitiveAcceptDoubles
End Property

Public Property Let CaseSensitiveAcceptDoubles(ByVal New_CCD As Boolean)
    m_CaseSensitiveAcceptDoubles = New_CCD
    PropertyChanged "CaseSensitiveAcceptDoubles"
End Property


Public Property Get Picture() As StdPicture
    Set Picture = m_picture
End Property

Public Property Set Picture(ByVal New_pic As StdPicture)
    Set m_picture = New_pic
    Set imgItem.Picture = New_pic
    ChangeShowImages
    PropertyChanged "Picture"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txttext,txttext,-1,Font
Public Property Get Font() As Font
    Set Font = txtText.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtText.Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=picCmd,picCmd,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = txtText.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtText.ToolTipText() = New_ToolTipText
    'ImgIcon.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property


Public Property Get Sorted() As Boolean
    Sorted = m_Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    If m_Sorted <> New_Sorted Then
    m_Sorted = New_Sorted
    'ImgIcon.ToolTipText() = New_ToolTipText
    PropertyChanged "Sorted"
    startsorting
    End If
End Property

Public Property Get ShowImages() As ShowPictures
    ShowImages = m_Showpictures
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    m_Locked = New_Locked
    PropertyChanged "Locked"
    
    txtText.Locked = m_Locked

End Property

Public Property Get Locked() As Boolean
    Locked = m_Locked
End Property

Public Property Let ShowImages(ByVal New_showpictures As ShowPictures)
    m_Showpictures = New_showpictures
    'ImgIcon.ToolTipText() = New_ToolTipText
    PropertyChanged "ShowImages"
    
    ChangeShowImages
    
End Property
Private Sub ChangeShowImages(Optional notresize As Boolean = False)
Select Case m_Showpictures
Case Auto
    
    
    Dim data As String
    On Error Resume Next
    data = Trim(Str(m_picture))
    
    If HowManyPicturesHaveWeGot > 0 Or (Trim(data) <> "" And Trim(data) <> "0") Then
    GoTo yep
    Else
    GoTo nope
    End If

Case Never
nope:
ShowPictureNow = False
Case Always
yep:
ShowPictureNow = True
End Select

If Not notresize Then UserControl_Resize
End Sub


'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    'Global Constants Initialization.
    Set m_picture = LoadPicture("")
    'm_AutoComplete = True
    m_TextAlign = Align_left
    m_FontColor = m_def_FontColor
    m_FontHighlightColor = m_def_FontHighlightColor
    m_Enabled = m_def_Enabled
    txtText.text = "" 'Extender.Name
    m_Sorted = False
    m_Locked = False
    m_Showpictures = Auto
    m_max = -1
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    Dim picNormal As Picture
    With PropBag
        Set picNormal = PropBag.ReadProperty("Picture", Nothing)
        
        If Not (picNormal Is Nothing) Then
        Set m_picture = picNormal
        Else
        Set m_picture = LoadPicture()
        End If
    
    End With

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    txtText.text = PropBag.ReadProperty("Caption", "Caption")
    Set txtText.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtText.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    m_HoverColor = PropBag.ReadProperty("HoverColor", m_def_HoverColor)
    m_TextAlign = PropBag.ReadProperty("TextAlign", Align_left)
    m_AutoComplete = PropBag.ReadProperty("AutoComplete", m_AutoComplete)
    m_Sorted = PropBag.ReadProperty("Sorted", m_Sorted)
    m_IconSize = PropBag.ReadProperty("IconSize", m_def_IconSize)
    m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Locked = PropBag.ReadProperty("Locked", False)
    m_Showpictures = PropBag.ReadProperty("Showpictures", 0)
       
    m_max = PropBag.ReadProperty("MaximumItems", -1)
    m_AcceptDoubles = PropBag.ReadProperty("AcceptDoubles", True)
    m_MaxExceeded = PropBag.ReadProperty("MaxExceeded", OverWriteFirst)
    m_CaseSensitiveAcceptDoubles = PropBag.ReadProperty("CaseSensitiveAcceptDoubles", False)

End Sub

Private Sub UserControl_Show()
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If Not cdown Is Nothing Then
        Unload cdown
        Set cdown = Nothing
    End If
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Sorted", m_Sorted, False)
    Call PropBag.WriteProperty("Caption", txtText.text)
    Call PropBag.WriteProperty("Font", txtText.Font, Ambient.Font)
    Call PropBag.WriteProperty("ToolTipText", txtText.ToolTipText, "")
    Call PropBag.WriteProperty("TextAlign", m_TextAlign, Align_left)
    Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Locked", m_Locked, False)
    Call PropBag.WriteProperty("Showpictures", m_Showpictures, Auto)
    Call PropBag.WriteProperty("Picture", m_picture, Nothing)
    Call PropBag.WriteProperty("AutoComplete", m_AutoComplete, True)
    Call PropBag.WriteProperty("MaximumItems", m_max, -1)
    Call PropBag.WriteProperty("AcceptDoubles", m_AcceptDoubles, True)
    Call PropBag.WriteProperty("MaxExceeded", m_MaxExceeded, OverWriteFirst)
    Call PropBag.WriteProperty("CaseSensitiveAcceptDoubles", m_CaseSensitiveAcceptDoubles, False)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,4
Public Property Get TextAlign() As AlignMent
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal New_AlignMent As AlignMent)
    m_TextAlign = New_AlignMent
    UserControl_Resize
    PropertyChanged "TextAlign"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=10,0,0,0
Public Property Get FontColor() As OLE_COLOR
    FontColor = m_FontColor
End Property

Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
    m_FontColor = New_FontColor
    txtText.ForeColor = New_FontColor
    PropertyChanged "FontColor"
End Property

Public Property Get AutoComplete() As Boolean
    AutoComplete = m_AutoComplete
End Property

Public Property Let AutoComplete(ByVal New_AutoComplete As Boolean)
    m_AutoComplete = New_AutoComplete
    PropertyChanged "AutoComplete"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    UserControl.Enabled = New_Enabled
    If New_Enabled Then
        txtText.BackColor = vbWindowBackground
        UserControl.BackColor = vbWindowBackground
        txtText.ForeColor = vbButtonText
        txtText.Locked = False
    Else
        txtText.BackColor = vb3DFace
        UserControl.BackColor = vb3DFace
        txtText.ForeColor = vbGrayText
        txtText.Locked = True
    End If
    UserControl_Paint
    PropertyChanged "Enabled"
End Property


Private Sub picButton_Click()
    ' Show de Auxiliar Window
    'On Error GoTo NoItemsToShow
    Dim ni As Integer

    If dontshow Then dontshow = False: Exit Sub
    
    wndDown.lblhover.Top = wndDown.lblCaption(0).Top
    wndDown.lblhover.Left = wndDown.lblCaption(0).Left
    On Error GoTo doorgaan
    wndDown.lblhover.Caption = m_Items(1)
    wndDown.lblhover.Visible = True
    'wndDown.lblhover.BackColor = vbHighlight
    'wndDown.lblhover.ForeColor = vbHighlightText
    
latenzien:
   
    If m_Enabled Then
        
        Set cdown = New wndDown
        Set hdown = cdown
        
        ChangeShowImages
        
        cdown.ShowPictureNow = ShowPictureNow
        
        For ni = 1 To cdown.m_Items.Count
        cdown.m_Images.Remove cdown.m_Items.Count
        cdown.m_Items.Remove cdown.m_Items.Count
        Next ni
        
        For ni = 1 To m_Items.Count
            cdown.m_Items.Add m_Items(ni)
            cdown.m_Images.Add m_Images(ni)
        Next ni
        
        RaiseEvent ButtonClick
        Dim rt As RECT
        
        ChangeShowImages
        
        GetWindowRect UserControl.hwnd, rt
        'cDown.Show , UserControl.Extender.parent
        cdown.PopUp rt.Left * gScaleX, rt.Bottom * gScaleY, UserControl.Width, UserControl.Extender.parent
        
        
    End If
Exit Sub
doorgaan:
wndDown.lblhover.Visible = False
GoTo latenzien
End Sub

Private Sub DrawFace(iState As Integer)
    '' This is the drawing code, I know there are better ways to do this,
    '' but I writte this 6 months and, and I don't want to work on this :)
    If iState = 4 Or iState = 0 Or iState = 1 Then iState = 3
    UserControl.ScaleMode = 3
    Select Case iState
    '    Case 0: 'Normal
    '        UserControl.Cls
    '        UserControl.DrawWidth = 2
    '        UserControl.ForeColor = vb3DFace
    '        UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
    '        UserControl.Line (1, 1)-(1, ScaleHeight + 1)
    '        UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
    '        UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
    '        picButton.Line (0, 0)-(picButton.ScaleWidth - 1, picButton.ScaleHeight - 1), vbWindowBackground, B
    '    Case 1: 'Pushed
    '        UserControl.ForeColor = vb3DLight
    '        UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
    '        UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
    '        UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
    '        UserControl.Line (1, 1)-(1, ScaleHeight + 1)
    '        UserControl.ForeColor = vb3DShadow
    '        UserControl.Line (0, 0)-(ScaleWidth, 0)
    '        UserControl.Line (0, 0)-(0, ScaleHeight)
    '        UserControl.ForeColor = vbWindowBackground
    '        UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
    '        UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
    '        picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vb3DShadow
    '        picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vb3DShadow
    '        picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DFace
    '        picButton.Line (1, 0)-(1, picButton.ScaleHeight - 1), vbWindowBackground
        
        Case 2: 'Disabled
            'picbutton.Cls
            'picButton_Paint
            txtText.BackColor = vb3DFace
            UserControl.BackColor = vb3DFace
            UserControl.Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), vbWindowBackground, B
        Case 3: 'Highlight
            UserControl.DrawWidth = 1
            UserControl.ForeColor = vb3DLight
            UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
            UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
            UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
            UserControl.Line (1, 1)-(1, ScaleHeight + 1)
            UserControl.ForeColor = vb3DShadow
            UserControl.Line (0, 0)-(ScaleWidth, 0)
            UserControl.Line (0, 0)-(0, ScaleHeight)
            UserControl.ForeColor = vbWindowBackground
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
            UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
            'picbutton.Line (picbutton.ScaleWidth - 1, picbutton.ScaleHeight - 1)-(picbutton.ScaleWidth - 1, 0), vb3DShadow
            'picbutton.Line (picbutton.ScaleWidth - 1, picbutton.ScaleHeight - 1)-(0, picbutton.ScaleHeight - 1), vb3DShadow
            'picbutton.Line (0, 0)-(0, picbutton.ScaleHeight - 1), vb3DFace
            'picbutton.Line (1, 0)-(1, picbutton.ScaleHeight - 1), vbWindowBackground
       ' Case 4: 'Focused
       '     UserControl.ForeColor = vb3DLight
       '     UserControl.Line (ScaleWidth - 2, 2)-(ScaleWidth - 2, ScaleHeight - 2)
       '     UserControl.Line (ScaleWidth - 2, ScaleHeight - 2)-(2, ScaleHeight - 2)
       '     UserControl.Line (1, 1)-(ScaleWidth + 1, 1)
       '     UserControl.Line (1, 1)-(1, ScaleHeight + 1)
       '     UserControl.ForeColor = vb3DShadow
       '     UserControl.Line (0, 0)-(ScaleWidth, 0)
       '     UserControl.Line (0, 0)-(0, ScaleHeight)
       '     UserControl.ForeColor = vbWindowBackground
       '     UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(ScaleWidth - 1, -1)
       '     UserControl.Line (ScaleWidth - 1, ScaleHeight - 1)-(-1, ScaleHeight - 1)
       '     picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(picButton.ScaleWidth - 1, 0), vb3DShadow
       '     picButton.Line (picButton.ScaleWidth - 1, picButton.ScaleHeight - 1)-(0, picButton.ScaleHeight - 1), vb3DShadow
       '     picButton.Line (0, 0)-(0, picButton.ScaleHeight - 1), vb3DFace
       '     picButton.Line (1, 0)-(1, picButton.ScaleHeight - 1), vbWindowBackground
    End Select
    UserControl.ScaleMode = 1
End Sub

Sub About()
Attribute About.VB_UserMemId = -552
Form2.Show 1
End Sub

'' Add a new Item to the Combo List
Public Sub AddItem(text As String, Optional Index As Integer = -1, Optional iImage As Picture)
    
On Error Resume Next
If m_max = 0 Then Exit Sub

If m_Items.Count = m_max And m_Items.Count > 0 Then

Dim z As Integer

        Select Case m_MaxExceeded
        
        Case OverWriteFirst  'remove first
        
        z = -1
        Do
        z = SearchNumberByTextsh(m_Items(1), z + 1)
            
            If z <> -1 Then
                
                Dim str1 As String, str2 As String
                On Error Resume Next
                str1 = Str(m_Images(1))
                str2 = Str(sh_Images(z))
                
                If sh_Items(z) = m_Items(1) And Trim(str1) = Trim(str2) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                Exit Do
                End If
            End If
        
        Loop Until z = -1
        m_Items.Remove 1
        m_Images.Remove 1
        
        Case OverWriteLast   'remove last
        
        z = -1
        Do
        z = SearchNumberByTextsh(m_Items(m_Items.Count), z + 1)
            
            If z <> -1 Then
                
                On Error Resume Next
                str1 = Str(m_Images(m_Items.Count))
                str2 = Str(sh_Images(z))
                
                If sh_Items(z) = m_Items(m_Items.Count) And Trim(str1) = Trim(str2) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                Exit Do
                End If
            End If
        
        Loop Until z = -1
        
        m_Images.Remove m_Items.Count
        m_Items.Remove m_Items.Count
        
        Case OverWriteRandom 'remove random
        
        Randomize Timer
        Dim r As Integer
        r = (Rnd * m_Items.Count - 1) + 1
        
        Do
        z = SearchNumberByTextsh(m_Items(r), z + 1)
            
            If z <> -1 Then
                
                On Error Resume Next
                str1 = Str(m_Images(r))
                str2 = Str(sh_Images(z))
            
                If sh_Items(z) = m_Items(r) And Trim(str1) = Trim(str2) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                Exit Do
                End If
            End If
        
        Loop Until z = -1
        
        m_Images.Remove r
        m_Items.Remove r
        
        Case DontAdd    'do nothing
        Exit Sub
        
        End Select

End If

If Not m_AcceptDoubles Then

z = Me.SearchNumberByText(text, , m_CaseSensitiveAcceptDoubles)
    
    If z <> -1 Then  'we found a match what to do
        
        Exit Sub 'dont add doubles so...
                
    End If

End If


If Index < -1 Then Exit Sub 'no good index

    Dim ImageTemp As Picture
    
    
    If IsMissing(iImage) Then
        Set ImageTemp = LoadPicture()
    Else
        Set ImageTemp = iImage
        
        On Error Resume Next
        Dim wat As String
        wat = Trim(Str(ImageTemp))   'count if we have a picture
        If Trim(wat) <> "" Then
        HowManyPicturesHaveWeGot = HowManyPicturesHaveWeGot + 1
            If HowManyPicturesHaveWeGot = 1 Then
            ChangeShowImages
            UserControl_Resize
            End If
        End If
    End If

If Index > -1 Then
    
    If Not m_Sorted Then
    PutItemOnPlaceInCollection text, Index, iImage
    addsorted text, ImageTemp, True
    Else
    PutItemOnPlaceInCollection text, Index, iImage, True
    addsorted text, ImageTemp
    End If
    
    GoTo laatste

End If
    
    If Not m_Sorted Then
    'add sorted to sh_items
    m_Items.Add text
    m_Images.Add ImageTemp
    addsorted text, ImageTemp, True
    Else
    'add sorted to m_items
    addsorted text, ImageTemp
    sh_Items.Add text
    sh_Images.Add ImageTemp
    End If
    
laatste:
    
    If Trim(txtText.text) = "" Then
    'set the caption if it's nothing to the first entry
    On Error Resume Next
    imgItem.Picture = m_Images(1)
    txtText.text = m_Items(1)
    txtText.SelStart = 0
    txtText.SelLength = Len(m_Items(1))
    selitem = 1
    End If

End Sub

Public Sub Clear()
On Error Resume Next
For x = 1 To m_Items.Count

    m_Items.Remove 1
    sh_Items.Remove 1
    m_Images.Remove 1
    sh_Images.Remove 1
    txtText.text = ""
    Set imgItem.Picture = LoadPicture()
Next

End Sub

Private Sub addsorted(text As String, Optional pic As StdPicture, Optional sh As Boolean = False)
Dim eind As Boolean, ubnd As Integer, lbnd As Integer, g As Integer

Dim NewPic As Picture
    
    If IsMissing(pic) Then
        Set NewPic = LoadPicture()
    Else
        Set NewPic = pic
    End If

If Not sh Then
'add to m_items
lbnd = 1
ubnd = m_Items.Count

Do
g = (lbnd + ubnd) / 2
'Debug.Print sh_Items(g)
If g > 0 Then
    Select Case StrComp(text, m_Items(g))
    Case -1
    'text is smaller so it must be left from g
    If ubnd = g And ubnd <> lbnd Then
    ubnd = ubnd - 1
    Else
    ubnd = g
    End If
    
    If lbnd = ubnd Then
    
        If StrComp(text, m_Items(ubnd)) = 1 Then
        PutItemOnPlaceInCollection text, g - 1, pic, False
        Else
        PutItemOnPlaceInCollection text, lbnd - 1, pic, False
        End If
        
    eind = True
    
    End If
    
    Case 0
    'put it here we found as same word
    PutItemOnPlaceInCollection text, g, pic, False
    eind = True
    
    Case 1
    'text is bigger so it must be right from g
    
    If lbnd = g And ubnd <> lbnd Then
    lbnd = lbnd + 1
    Else
    lbnd = g
    End If
    
    If lbnd = ubnd Then
    
    If StrComp(text, m_Items(ubnd)) = -1 And ubnd > 0 Then
    PutItemOnPlaceInCollection text, g, pic, False
    Else
    
        If ubnd = m_Items.Count Then
        m_Items.Add text
        m_Images.Add NewPic
        Else
        PutItemOnPlaceInCollection text, g, pic, False
        End If
        
    End If
    eind = True
    End If
    
    End Select

Else
m_Items.Add text
m_Images.Add NewPic
eind = True
End If

Loop Until eind

Else

'add to sh_items

lbnd = 1
ubnd = sh_Items.Count

Do
g = (lbnd + ubnd) / 2
If g > 0 Then
    Select Case StrComp(text, sh_Items(g))
    Case -1
    'text is smaller so it must be left from g
    If ubnd = g And ubnd <> lbnd Then
    ubnd = ubnd - 1
    Else
    ubnd = g
    End If
    
    If lbnd = ubnd Then
  
        If StrComp(text, sh_Items(ubnd)) = 1 Then
        PutItemOnPlaceInCollection text, g - 1, pic, True
        Else
        PutItemOnPlaceInCollection text, lbnd - 1, pic, True
        End If
    
    eind = True
    
    End If
    
    Case 0
    'put it here we found as same word
    PutItemOnPlaceInCollection text, g, pic, True
    eind = True
    
    Case 1
    'text is bigger so it must be right from g
    
    If lbnd = g And ubnd <> lbnd Then
    lbnd = lbnd + 1
    Else
    lbnd = g
    End If
    
    If lbnd = ubnd Then
    
        If StrComp(text, sh_Items(ubnd)) = -1 And ubnd > 0 Then
        PutItemOnPlaceInCollection text, g, pic, True
        Else
    
            If ubnd = sh_Items.Count Then
            sh_Items.Add text
            sh_Images.Add NewPic
            Else
            PutItemOnPlaceInCollection text, g, pic, True
            End If
        
        End If
    eind = True
    End If
    
    End Select

Else
sh_Items.Add text
sh_Images.Add NewPic
eind = True
End If

Loop Until eind

End If
End Sub

Public Function RemoveItemByNumber(nr As Integer) As Boolean
Dim z As Integer

RemoveItemByNumber = False
If nr > -1 Then
    'wrong input
    RemoveItemByNumber = False
    On Error GoTo endje
    
    
    'first remove from sh
        Do
        z = SearchNumberByTextsh(m_Items(nr))
            
            If z <> -1 Then
                
                Dim str1 As String, str2 As String
                On Error Resume Next
                str1 = Str(m_Images(nr))
                str2 = Str(sh_Images(z))
                                
                If sh_Items(z) = m_Items(nr) And Trim(str1) = Trim(str2) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                End If
            End If
        
        Loop Until z = -1
    
    
    m_Items.Remove nr
    
    
    On Error Resume Next
    Dim wat As String
    wat = Str(m_Images(nr))
        If Trim(wat) <> "" Then
        HowManyPicturesHaveWeGot = HowManyPicturesHaveWeGot - 1
        If HowManyPicturesHaveWeGot < 0 Then HowManyPicturesHaveWeGot = 0
        
            If HowManyPicturesHaveWeGot = 0 Then
            ChangeShowImages
            UserControl_Resize
            End If
        
        End If
    
    On Error GoTo endje
    m_Images.Remove nr
    RemoveItemByNumber = True
End If
endje:
End Function
Public Function RemoveItemByText(text As String, Optional CaseSensitive As Boolean = True, Optional All As Boolean = True, Optional start As Integer = 1) As Integer
Dim x As Integer, p As Integer, z As Integer, str1 As String, str2 As String
RemoveItemByText = -1
If start < 1 Then Exit Function
For x = start To m_Items.Count

    If (m_Items(x - p) = text And CaseSensitive) Or (UCase$(m_Items(x - p)) = UCase$(text) And Not CaseSensitive) Then
    
        'first remove from sh
        Do
        z = SearchNumberByTextsh(m_Items(x - p))
            
            If z <> -1 Then
                
                On Error Resume Next
                str1 = Str(m_Images(x - p))
                str2 = Str(sh_Images(z))
                
                If sh_Items(z) = m_Items(x - p) And Trim(str1) = Trim(str2) Then
                'm_Images(x - p) = sh_Images(z) Then
                'remove this one
                sh_Items.Remove z
                sh_Images.Remove z
                End If
            End If
        
        Loop Until z = -1
    
    m_Items.Remove x - p
    
    On Error Resume Next
        Dim wat As String
        wat = Str(m_Images(x - p))
        If Trim(wat) <> "" Then
        HowManyPicturesHaveWeGot = HowManyPicturesHaveWeGot - 1
        If HowManyPicturesHaveWeGot < 0 Then HowManyPicturesHaveWeGot = 0
        
            If HowManyPicturesHaveWeGot = 0 Then
            ChangeShowImages
            UserControl_Resize
            End If
        
        End If
    
    
    m_Images.Remove x - p
    p = p + 1
    
    If RemoveItemByText = -1 Then RemoveItemByText = 0
    RemoveItemByText = RemoveItemByText + 1
    End If
    
    If Not All Then Exit Function

Next

End Function
Public Function CountDoubles(text As String, Optional start As Integer = 0, Optional MatchCase As Boolean = True, Optional WholeText As Boolean = True) As Integer
Dim x As Integer

CountDoubles = 0

If Trim(text) = "" Then Exit Function

Do
x = SearchNumberByText(text, start, MatchCase, WholeText)
If x > -1 Then CountDoubles = CountDoubles + 1
start = x + 1
Loop Until x = -1

End Function

Public Function SearchNumberByText(text As String, Optional start As Integer = 0, Optional MatchCase As Boolean = False, Optional WholeText As Boolean = True) As Integer
Dim items As String
SearchNumberByText = -1

If Trim(text) = "" Then Exit Function

On Error GoTo endje

    If Not MatchCase Then text = UCase$(text)
    
Dim x As Integer

start = start + 1

For x = start To m_Items.Count

items = m_Items(x)

    If Not MatchCase Then items = UCase$(items)
    
    If StrComp(items, text, vbBinaryCompare) = 0 Then  'we found a match
    SearchNumberByText = x - 1
    Exit Function
    End If
    
    If Not WholeText Then   'search in the item
        If InStr(1, items, text, vbBinaryCompare) Then 'we found a match
        SearchNumberByText = x - 1
        Exit Function
        End If
    End If
Next
endje:
End Function

Private Function SearchNumberByTextsh(text As String, Optional start As Integer = 0) As Integer
Dim items As String, x As Integer
SearchNumberByTextsh = -1

If Trim(text) = "" Then Exit Function

start = start + 1

For x = start To sh_Items.Count

items = sh_Items(x)
    
    If StrComp(items, text, vbBinaryCompare) = 0 Then  'we found a match
    SearchNumberByTextsh = x
    Exit Function
    End If
    
Next
endje:
End Function


Sub ChangeItemText(nr As Integer, NewText As String)
'On Error GoTo endje
Dim pic As StdPicture

nr = nr + 1

If nr > m_Items.Count Then Exit Sub

Set pic = m_Images(nr)

m_Images.Remove nr
m_Items.Remove nr

nr = nr - 1

PutItemOnPlaceInCollection NewText, nr, pic

endje:
End Sub

Sub ChangeItemPicture(nr As Integer, NewPic As StdPicture)
'On Error GoTo endje
Dim text As String

If nr + 1 > m_Items.Count Then Exit Sub
nr = nr + 1

text = m_Items(nr)

        On Error Resume Next
        Dim wat As String
        wat = Trim(Str(m_Images(nr)))
        If Trim(wat) <> "" Then
        HowManyPicturesHaveWeGot = HowManyPicturesHaveWeGot - 1
            If HowManyPicturesHaveWeGot = 0 Then
            ChangeShowImages
            UserControl_Resize
            End If
        End If



m_Images.Remove nr
m_Items.Remove nr

nr = nr - 1


        wat = Trim(Str(NewPic))
        If Trim(wat) <> "" Then
        HowManyPicturesHaveWeGot = HowManyPicturesHaveWeGot + 1
            If HowManyPicturesHaveWeGot = 1 Then
            ChangeShowImages
            UserControl_Resize
            End If
        End If


PutItemOnPlaceInCollection text, nr, NewPic

endje:
End Sub

Private Sub PutItemOnPlaceInCollection(ItemText As String, NewPlace As Integer, Optional newpicture As StdPicture, Optional sh As Boolean = False)

Dim ImageTemp As Picture
    If IsMissing(newpicture) Then
        Set ImageTemp = LoadPicture()
    Else
        Set ImageTemp = newpicture
    End If
    

If Not sh Then
    If NewPlace + 1 > m_Items.Count Then
        'the index isn't used so just add
        m_Items.Add ItemText
        m_Images.Add ImageTemp
    Else
        'put it on the right place
        m_Items.Add ItemText, , NewPlace + 1
        m_Images.Add ImageTemp, , NewPlace + 1
    End If
Else
    If NewPlace + 1 > sh_Items.Count Then
        'the index isn't used so just add
        sh_Items.Add ItemText
        sh_Images.Add ImageTemp
    Else
        'put it on the right place
        sh_Items.Add ItemText, , NewPlace + 1
        sh_Images.Add ImageTemp, , NewPlace + 1
    End If
End If
End Sub
Private Sub startsorting()

Dim tempcollection As New Collection    'to copy m_items to
Dim temp2collection As New Collection   'to copy sh_items to

Dim x As Integer

For x = m_Items.Count To 1 Step -1

tempcollection.Add m_Items(x)
temp2collection.Add sh_Items(x)
m_Items.Remove x
sh_Items.Remove x

Next

For x = tempcollection.Count To 1 Step -1

sh_Items.Add tempcollection(x)
m_Items.Add temp2collection(x)
tempcollection.Remove x
temp2collection.Remove x

Next

For x = m_Items.Count To 1 Step -1

tempcollection.Add m_Images(x)
temp2collection.Add sh_Images(x)
m_Images.Remove x
sh_Images.Remove x

Next

For x = tempcollection.Count To 1 Step -1

sh_Images.Add tempcollection(x)
m_Images.Add temp2collection(x)
tempcollection.Remove x
temp2collection.Remove x

Next


End Sub

