VERSION 5.00
Begin VB.Form wndDown 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form2"
   ScaleHeight     =   3270
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pScroller 
      BackColor       =   &H80000005&
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.VScrollBar vsb 
         Height          =   1575
         LargeChange     =   10
         Left            =   2160
         Max             =   115
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.PictureBox picGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   60
         ScaleHeight     =   1935
         ScaleWidth      =   1875
         TabIndex        =   1
         Top             =   30
         Width           =   1875
         Begin VB.Label lblhover 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            Caption         =   "lblhover"
            Height          =   195
            Left            =   390
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Image ImgItem 
            Height          =   240
            Index           =   0
            Left            =   60
            Picture         =   "wndDown.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblCaption 
            BackColor       =   &H80000005&
            Caption         =   "Item-0"
            Height          =   205
            Index           =   0
            Left            =   420
            TabIndex        =   2
            Top             =   0
            Visible         =   0   'False
            Width           =   1425
         End
      End
   End
End
Attribute VB_Name = "wndDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
''      ControlName:    ISCombo.
''      Filename:       wndDown.frm( Don't modify this form ! !)
''
''      Author:         Alfredo Córdova Pérez ( fred_cpp )
''      e-mail:         fred_cpp@hotmail.com
''                      fred_cpp@yahoo.com.mx
''
''      Description:
''
''      I've Got a lot of problemas with the VB' combo, I couldn't detect
''      when the user changes the selection, and, those combos are relly ugly :(
''      so, I decided made one better.
''      you know, you can use this freely, just give me credit.
''      Votes and suggestions are wellcome.
''



Option Explicit

Public ShowPictureNow As Boolean

Dim iPos As Integer
Dim iItems As Integer
Dim IsInside As Boolean
Dim iPrevPos As Integer
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

Private Const WM_SIZE = &H5
Private Const WM_MOVE = &H3
Private Const WM_WINDOWPOSCHANGED = &H47
Private Const WM_KILLFOCUS = &H8
Private Const GWL_WNDPROC = (-4)
Private OriginalWndProc As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, Y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Public m_Items As New Collection
Public m_Images As New Collection
Public m_ShowingList As Boolean
Public ItemClick As Integer

Event ItemClick(iItem As Integer, sText As String, DontSay As Boolean)
Event MouseHover(Index As Integer)
Event unloaded()

Dim nValue As Long

'' Detect if the Mouse cursor is inside a Window
Private Function InBox(ObjectHWnd As Long) As Boolean
    Dim mpos As PointAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect ObjectHWnd, oRect
    If mpos.x >= oRect.Left And mpos.x <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
End Function

Private Sub DrawAll(ActiveItem As Integer)
    lblhover.Visible = True
    lblhover.Top = lblCaption(ActiveItem).Top
    lblhover.Caption = lblCaption(ActiveItem).Caption
    lblhover.Left = lblCaption(ActiveItem).Left
       
    If ActiveItem <= 0 Then
        iPrevPos = 0
    Else
        iPrevPos = ActiveItem
    End If
    RaiseEvent MouseHover(ActiveItem)
'    Debug.Print "DrawAll"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RaiseEvent unloaded
End Sub

Private Sub imgItem_Click(Index As Integer)
    Reset
    Dim nr As Integer
    nr = iPrevPos + vsb.Value - 1
    If m_Items.Count <= 8 Then nr = nr + 1
    RaiseEvent ItemClick(nr, lblCaption(Index).Caption, False)
End Sub


Private Sub imgItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
lblCaption_MouseMove Index, Button, Shift, x, Y
End Sub

'' Raise the ItemClick event
Private Sub lblCaption_Click(Index As Integer)
    Reset
    Dim nr As Integer
    nr = iPrevPos + vsb.Value - 1
    If m_Items.Count <= 8 Then nr = nr + 1
    RaiseEvent ItemClick(nr, lblCaption(Index).Caption, False)
End Sub

'' Detect the mouse movement
Private Sub lblCaption_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
      
DrawAll Index
  
End Sub

Private Sub lblhover_Click()
    Reset
    Dim nr As Integer
    nr = iPrevPos + vsb.Value - 1
    If m_Items.Count <= 8 Then nr = nr + 1
    RaiseEvent ItemClick(nr, lblhover.Caption, False)
End Sub

Private Sub picGroup_Click()
lblhover_Click
End Sub

''  Hide and unload if the window lost the focus
Private Sub picGroup_LostFocus()
    Reset
End Sub

''  Hide and unload if the window lost the focus
Private Sub Form_LostFocus()
    Reset
End Sub

Private Sub pScroller_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then
    
    If lblhover.Top = lblCaption(0).Top Or lblhover.Visible = False Then
        If vsb.Value = vsb.Min Or lblhover.Visible = False Then
        Unload Me
        Else
        vsb.Value = vsb.Value - 1
        End If
    Else
    iPrevPos = iPrevPos - 1
    lblhover.Caption = lblCaption(iPrevPos).Caption
    lblhover.Top = lblCaption(iPrevPos).Top
    lblhover.Left = lblCaption(iPrevPos).Left
    End If

ElseIf KeyCode = 40 Then

    If iPrevPos + vsb.Value < m_Items.Count And (iPrevPos <> m_Items.Count - 1) Then
    
    If iPrevPos = 7 Then
    vsb.Value = vsb.Value + 1
    Else
    iPrevPos = iPrevPos + 1
    End If
    
    lblhover.Caption = lblCaption(iPrevPos).Caption
    lblhover.Top = lblCaption(iPrevPos).Top
    lblhover.Left = lblCaption(iPrevPos).Left
    
    End If

ElseIf KeyCode = 13 Then

lblhover_Click

End If
End Sub

''  Hide and unload if the window lost the focus
Private Sub pScroller_LostFocus()
    'If vsb.Visible Then Exit Sub
    Reset
End Sub
'' Change the position of the items when the ScrollBar changes
Private Sub vsb_Change()
    On Error Resume Next
    veranderen
    Me.SetFocus
End Sub

'' Hide Window and Save state in Variable
Private Sub Reset()
    Hide
    m_ShowingList = False
End Sub

'' This function Show the cDown Window, And adds the items
Public Function PopUp(x As Long, Y As Long, lWidth As Single, parent As Object) As Boolean
'On Error Resume Next
    
    Dim ni As Integer
    Dim ht As Single
    Dim lHeight As Single
    vsb.Value = vsb.Min
    m_ShowingList = True
    
    
    Dim hoogste As Integer
    
    On Error Resume Next
    vsb.Value = 0
    hoogste = m_Items.Count
    If hoogste > 8 Then hoogste = 8
    Select Case hoogste
    Case 0
    ht = 500
    Case 1 To 7
    ht = 255 * (hoogste) + 45
    Case 8
    If m_Items.Count = 8 Then
    ht = 255 * (hoogste) + 45
    Else
    ht = 255 * (hoogste) + 5
    End If
    
    End Select
    
    For ni = 1 To 8
        Unload imgItem(ni)
        Unload lblCaption(ni)
    Next
        
    
    For ni = 1 To hoogste
        Load lblCaption(ni)
        Load imgItem(ni)
        lblCaption(ni - 1).Visible = True
        lblCaption(ni - 1).Caption = m_Items.Item(ni)
        lblCaption(ni - 1).Width = Me.Width - lblCaption(ni - 1).Left
        lblCaption(ni - 1).BackColor = vbWindowBackground
        
        If ShowPictureNow Then
        imgItem(ni - 1).Visible = True
        Set imgItem(ni - 1).Picture = m_Images(ni)
        imgItem(ni - 1).Move 30, 255 * (ni - 1)
        lblCaption(ni - 1).Move 360, 255 * (ni - 1)
        Else
        lblCaption(ni - 1).Move 30, 255 * (ni - 1)
        imgItem(ni - 1).Visible = False
        End If
        
    Next ni
    
    lblhover.Left = lblCaption(0).Left
    
LimitOfItems:
    
    If m_Items.Count <= 8 Then
        lHeight = ht
        vsb.Visible = False
    Else
        lHeight = 8 * 255 + 60
        vsb.Visible = True
        vsb.Min = 1
        vsb.Max = m_Items.Count - 7
    End If
    
    Visible = True
    Move x, Y, lWidth, lHeight
    Show ', 'parent
    picGroup.Move 15, 15, Width, ht
    pScroller.Move 30, 30, Width - 60, lHeight - 60
    vsb.Move Width - vsb.Width - 2 * Screen.TwipsPerPixelX - 60, 0, vsb.Width, lHeight - 90
    pScroller.Move 15, 15, Width - 30, lHeight - 30
    
    For ni = 1 To hoogste
        lblCaption(ni - 1).Width = Me.Width - lblCaption(ni - 1).Left
    Next
    
    iPrevPos = 0
'    If vsb.Visible Then
'        'vsb.SetFocus
'        pScroller.SetFocus
'    Else
'        picGroup.SetFocus
'    End If
    Me.SetFocus
End Function

Sub veranderen()
Dim x As Integer

'On Error GoTo endje
For x = vsb.Value To vsb.Value + 8
        imgItem(x - vsb.Value) = m_Images(x)
        lblCaption(x - vsb.Value) = m_Items.Item(x)
        If x - vsb.Value = iPrevPos Then lblhover = lblCaption(x - vsb.Value)
Next


endje:
End Sub
