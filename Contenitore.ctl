VERSION 5.00
Begin VB.UserControl MiniForm 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1212
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1296
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   Picture         =   "Contenitore.ctx":0000
   PropertyPages   =   "Contenitore.ctx":038A
   ScaleHeight     =   1212
   ScaleWidth      =   1296
   ToolboxBitmap   =   "Contenitore.ctx":0396
   Begin VB.PictureBox Background 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   648
      Left            =   0
      ScaleHeight     =   648
      ScaleWidth      =   1056
      TabIndex        =   8
      Top             =   300
      Width           =   1056
   End
   Begin VB.PictureBox Barra 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   276
      ScaleWidth      =   1272
      TabIndex        =   2
      Top             =   0
      Width           =   1296
      Begin VB.CommandButton CQuit 
         Height          =   204
         Left            =   984
         Picture         =   "Contenitore.ctx":06A8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   36
         Width           =   228
      End
      Begin VB.CommandButton CMax 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   756
         Picture         =   "Contenitore.ctx":0A32
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "Des"
         ToolTipText     =   "Maximize"
         Top             =   36
         Width           =   228
      End
      Begin VB.CommandButton CMin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   204
         Left            =   528
         Picture         =   "Contenitore.ctx":0DBC
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "Des"
         ToolTipText     =   "Minimize"
         Top             =   36
         Width           =   228
      End
      Begin VB.Image Icona 
         Height          =   204
         Left            =   60
         Stretch         =   -1  'True
         Top             =   48
         Width           =   216
      End
      Begin VB.Label Titolo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   312
         TabIndex        =   5
         Top             =   36
         Width           =   48
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   672
      Left            =   1056
      ScaleHeight     =   672
      ScaleWidth      =   240
      TabIndex        =   1
      Tag             =   "Des"
      Top             =   300
      Visible         =   0   'False
      Width           =   240
      Begin VB.VScrollBar VScroll 
         Height          =   684
         LargeChange     =   50
         Left            =   0
         Max             =   10000
         Min             =   -10000
         SmallChange     =   40
         TabIndex        =   7
         Top             =   -12
         Width           =   252
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1296
      TabIndex        =   0
      Tag             =   "Des"
      Top             =   972
      Visible         =   0   'False
      Width           =   1296
      Begin VB.HScrollBar HScroll 
         Height          =   252
         LargeChange     =   50
         Left            =   0
         Max             =   10000
         Min             =   -10000
         SmallChange     =   40
         TabIndex        =   6
         Top             =   -12
         Width           =   1296
      End
   End
End
Attribute VB_Name = "MiniForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim Altezza As Long, Lunghezza As Long, XLeft As Long, YTop As Long, HContainer As Long, WContainer As Long
Dim SuVisible As Boolean, SxVisible As Boolean
Dim Quitted As Boolean, MiniMized As Boolean, Maximized As Boolean
Dim DefAltezza As Long, DefLunghezza As Long

'Default Property Values:
Const m_def_Caption = ""
Const m_def_WindowStyle = 0
Const m_def_ScrollBars = 0
Const m_def_UseAsForm = 0
Const m_def_Moveable = 1
Const m_def_Gradient = 0
Const m_def_BackColor = &HC0C0C0
Const m_def_BarVisible = 1
Const m_def_BarStyle = 0
Const m_def_QuitVisible = 1
Const m_def_MaxVisible = 1
Const m_def_MinVisible = 1
Const m_def_QuitEnabled = 1
Const m_def_MaxEnabled = 1
Const m_def_MinEnabled = 1
Const m_def_ToolTipText = ""
Const m_def_WhatsThisHelpID = 0
Const m_def_BarColor = &H80000000

'Property Variables:
Dim m_Caption As String
Dim m_WindowStyle As Integer
Dim m_ScrollBars As Integer
Dim m_UseAsForm As Boolean
Dim m_Moveable As Boolean
Dim m_Gradient As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_BarVisible As Boolean
Dim m_BarStyle As Integer
Dim m_QuitVisible As Boolean
Dim m_MaxVisible As Boolean
Dim m_MinVisible As Boolean
Dim m_QuitEnabled As Boolean
Dim m_MaxEnabled As Boolean
Dim m_MinEnabled As Boolean
Dim m_ToolTipText As String
Dim m_WhatsThisHelpID As Long
Dim m_BarColor As OLE_COLOR

'Event Declarations:
Event Paint() 'MappingInfo=UserControl,UserControl,-1,Paint
Attribute Paint.VB_Description = "Occurs when any part of a form or PictureBox control is moved, enlarged, or exposed."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event Load() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Load.VB_Description = "Occurs when the control's Visible property changes to True."
Event Unload() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Unload.VB_Description = "Occurs when the control's Visible property changes to False."

Private Sub Background_Click()
RaiseEvent Click
End Sub

Private Sub Background_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Background_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub Background_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub Background_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Background_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Background_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Background_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Background_Paint()
RaiseEvent Paint
Background.Cls
UserControl.Cls
Barra.Cls
UserControl.PropertyChanged BarStyle
UserControl.PropertyChanged WindowStyle
UserControl.PropertyChanged Gradient
End Sub

Private Sub Barra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long

RaiseEvent MouseDown(Button, Shift, X, Y)

    If UseAsForm = True Then
            If TypeOf UserControl.Extender.Container Is Form Then
              'Permette di spostare il controllo ed il form che lo contiene
               If Button = 1 And Moveable = True Then
                 X = ReleaseCapture()
                 ReturnVal = SendMessage(UserControl.Extender.Parent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
               End If
            End If
       Else
            'Permette di spostare il controllo all'interno di un form
            If Button = 1 And Moveable = True Then
              X = ReleaseCapture()
              ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
            End If
    End If
End Sub

Private Sub Barra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Private Sub Barra_Resize()
Barra.Cls
UserControl.PropertyChanged BarStyle
End Sub

Private Sub CMax_Click()
Maximized = True

If UseAsForm = True Then
        If TypeOf UserControl.Extender.Container Is Form Then
            If UseAsForm = True And MiniMized = True Then
                 On Error Resume Next
                 CMin.Enabled = True
                 UserControl.Extender.Container.Height = HContainer
                 UserControl.Extender.Container.Width = WContainer
                 UserControl.Extender.Top = 0
                 UserControl.Extender.Left = 0
                 UserControl.Width = UserControl.Extender.Container.Width - 100
                 UserControl.Height = UserControl.Extender.Container.Height - 100
                 MiniMized = False
                 'Visualizzando il controllo dopo averlo nascosto, ripristina
                    'le frecce dello scrolling
                    If Quitted = True Then
                        If SxVisible = True Then
                           Picture1.Visible = True
                          Else
                           Picture1.Visible = False
                        End If
                    
                        If SuVisible = True Then
                           Picture2.Visible = True
                          Else
                           Picture2.Visible = False
                        End If
                        Quitted = False
                    End If
            
                    Picture1.Align = 2
                    Picture2.Align = 4
                    
                    UserControl.SetFocus
                    RaiseEvent Paint
                    UserControl.PropertyChanged UseAsForm
                    RaiseEvent Resize
             ElseIf UseAsForm = True And UserControl.Extender.Container.WindowState = vbNormal Then
                 CMin.Enabled = True
                 CMax.Enabled = True
                 UserControl.Extender.Container.WindowState = vbMaximized
                 CMax.Picture = LoadResPicture(101, vbResIcon)
                 UserControl.Extender.Top = 0
                 UserControl.Extender.Left = 0
                 UserControl.Width = UserControl.Extender.Container.Width - 100
                 UserControl.Height = UserControl.Extender.Container.Height - 100
                 
                 'Visualizzando il controllo dopo averlo nascosto, ripristina
                    'le frecce dello scrolling
                    If Quitted = True Then
                        If SxVisible = True Then
                           Picture1.Visible = True
                          Else
                           Picture1.Visible = False
                        End If
                        
                        If SuVisible = True Then
                           Picture2.Visible = True
                          Else
                           Picture2.Visible = False
                        End If
                        Quitted = False
                    End If
            
                    Picture1.Align = 2
                    Picture2.Align = 4
                    
                    UserControl.SetFocus
                    RaiseEvent Paint
                    UserControl.PropertyChanged UseAsForm
                    RaiseEvent Resize
             ElseIf UseAsForm = True And UserControl.Extender.Container.WindowState = vbMaximized Then
                 CMin.Enabled = True
                 CMax.Enabled = True
                 UserControl.Extender.Container.WindowState = vbNormal
                 CMax.Picture = LoadResPicture(100, vbResIcon)
                 UserControl.Extender.Top = 0
                 UserControl.Extender.Left = 0
                 UserControl.Width = UserControl.Extender.Container.Width - 100
                 UserControl.Height = UserControl.Extender.Container.Height - 100
                 
                 'Visualizzando il controllo dopo averlo nascosto, ripristina
                    'le frecce dello scrolling
                    If Quitted = True Then
                        If SxVisible = True Then
                           Picture1.Visible = True
                          Else
                           Picture1.Visible = False
                        End If
                        
                        If SuVisible = True Then
                           Picture2.Visible = True
                          Else
                           Picture2.Visible = False
                        End If
                        Quitted = False
                    End If
            
                    Picture1.Align = 2
                    Picture2.Align = 4
                    
                    UserControl.SetFocus
                    RaiseEvent Paint
                    UserControl.PropertyChanged UseAsForm
                    RaiseEvent Resize
            End If
        End If
    
    Else
        CMin.Enabled = True
        CMax.Enabled = False
        'Massimizzando il controllo dopo averlo minimizzato, ripristina
        'le frecce dello scrolling
        If Quitted = True Then
            If SxVisible = True Then
               Picture1.Visible = True
              Else
               Picture1.Visible = False
            End If
            
            If SuVisible = True Then
               Picture2.Visible = True
              Else
               Picture2.Visible = False
            End If
            Quitted = False
        End If

        Picture1.Align = 2
        Picture2.Align = 4

        If MiniMized = True Then
            UserControl.Extender.Height = Altezza 'Ripristina l'altezza
            UserControl.Extender.Width = Lunghezza  'Ripristina la lunghezza
            MiniMized = False
            RaiseEvent Resize
          Else
            MiniMized = False
            RaiseEvent Resize
        End If
        
        UserControl.SetFocus
        RaiseEvent Paint
End If
End Sub
Private Sub CMin_Click()

If UseAsForm = True Then
        If TypeOf UserControl.Extender.Container Is Form Then
            If UseAsForm = True And UserControl.Extender.Container.WindowState = vbMaximized Then
                 On Error Resume Next
                 MiniMized = True
                 CMin.Enabled = False
                 CMax.Enabled = True
                 'Settando a Min il controllo memorizza se erano presenti i pulsanti dello
                 'scrolling
                 If Picture1.Visible = True Then
                    SxVisible = True
                  Else
                    SxVisible = False
                 End If
                 
                 If Picture2.Visible = True Then
                    SuVisible = True
                  Else
                    SuVisible = False
                 End If
                
                 Picture1.Align = 0
                 Picture2.Align = 0
                
                 HContainer = UserControl.Extender.Container.Height
                 WContainer = UserControl.Extender.Container.Width
                 UserControl.Extender.Container.WindowState = vbMinimized
                 
                 Maximized = False
                 RaiseEvent Resize
                 
             ElseIf UseAsForm = True And UserControl.Extender.Container.WindowState = vbNormal Then
                 MiniMized = True
                 CMin.Enabled = False
                 CMax.Enabled = True
                 'Settando a Min il controllo memorizza se erano presenti i pulsanti dello
                 'scrolling
                 If Picture1.Visible = True Then
                    SxVisible = True
                  Else
                    SxVisible = False
                 End If
                 
                 If Picture2.Visible = True Then
                    SuVisible = True
                  Else
                    SuVisible = False
                 End If
                
                Picture1.Align = 0
                Picture2.Align = 0
                 
                 On Error Resume Next
                 HContainer = UserControl.Extender.Container.Height
                 WContainer = UserControl.Extender.Container.Width
                 UserControl.Extender.Container.WindowState = vbMinimized
                 RaiseEvent Resize
            End If
        End If
    Else
        
        MiniMized = True
        CMin.Enabled = False
        CMax.Enabled = True
         'Settando a Min il controllo memorizza se erano presenti i pulsanti dello
         'scrolling
         If Picture1.Visible = True Then
            SxVisible = True
          Else
            SxVisible = False
         End If
         
         If Picture2.Visible = True Then
            SuVisible = True
          Else
            SuVisible = False
         End If
        
        Picture1.Align = 0
        Picture2.Align = 0

        Lunghezza = UserControl.Width
        Altezza = UserControl.Height
        
        RaiseEvent Resize
        
        UserControl.Height = 345
        
        UserControl.SetFocus
End If
End Sub
Private Sub CQuit_Click()

Quitted = True
 'Settando a invisibile il controllo memorizza se erano presenti i pulsanti dello
 'scrolling
 If Picture1.Visible = True Then
    SxVisible = True
  Else
    SxVisible = False
 End If
 
 If Picture2.Visible = True Then
    SuVisible = True
  Else
    SuVisible = False
 End If
 
If UseAsForm = True Then
   If TypeOf UserControl.Extender.Container Is Form Then
      Unload UserControl.Extender.Parent  'Scarica il form
   End If
  Else
    'Nasconde il controllo utilizzando la proprietà Visible
    'dell'oggetto Extender del custom control
    UserControl.Extender.Visible = False
End If
End Sub
Private Sub HScroll_Change()
Static Valore As Long
Select Case HScroll.Value
   Case Is > Valore
        If UserControl.Ambient.UserMode = True Then
            For X = 0 To UserControl.ContainedControls.Count - 1
                On Error Resume Next
                UserControl.ContainedControls(X).Left = UserControl.ContainedControls(X).Left + HScroll.SmallChange
                Valore = HScroll.Value
            Next
        End If
    Case Is < Valore
        If UserControl.Ambient.UserMode = True Then
            For X = 0 To UserControl.ContainedControls.Count - 1
                On Error Resume Next
                UserControl.ContainedControls(X).Left = UserControl.ContainedControls(X).Left - HScroll.SmallChange
                Valore = HScroll.Value
            Next
        End If
End Select
End Sub

Private Sub HScroll_Scroll()
Static Valore As Long
Select Case HScroll.Value
   Case Is > Valore
        If UserControl.Ambient.UserMode = True Then
            For X = 0 To UserControl.ContainedControls.Count - 1
                On Error Resume Next
                UserControl.ContainedControls(X).Left = UserControl.ContainedControls(X).Left + HScroll.LargeChange
                Valore = HScroll.Value
            Next
        End If
    Case Is < Valore
        If UserControl.Ambient.UserMode = True Then
            For X = 0 To UserControl.ContainedControls.Count - 1
                On Error Resume Next
                UserControl.ContainedControls(X).Left = UserControl.ContainedControls(X).Left - HScroll.LargeChange
                Valore = HScroll.Value
            Next
        End If
End Select
End Sub

Private Sub Titolo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal As Long

RaiseEvent MouseDown(Button, Shift, X, Y)

    If UseAsForm = True Then
            If TypeOf UserControl.Extender.Container Is Form Then
              'Permette di spostare il controllo ed il form che lo contiene
               If Button = 1 And Moveable = True Then
                 X = ReleaseCapture()
                 ReturnVal = SendMessage(UserControl.Extender.Parent.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
               End If
            End If
       Else
            'Permette di spostare il controllo all'interno di un form
            If Button = 1 And Moveable = True Then
              X = ReleaseCapture()
              ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
            End If
    End If
End Sub

Private Sub Titolo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
RaiseEvent Paint
Background.Cls
UserControl.Cls
Barra.Cls
UserControl.PropertyChanged BarStyle
UserControl.PropertyChanged WindowStyle
UserControl.PropertyChanged Gradient
End Sub

Private Sub UserControl_Resize()
  'Impedisce di ridimensionare in larghezza al di sotto di un valore
   'che impedirebbe la visione dei controlli Min, Max, e chiudi

    If UserControl.Width <= 696 Then
         UserControl.Width = 800
    End If
    If UserControl.Height <= 300 Then
       UserControl.Height = 345
    End If

Background.ZOrder 1
Background.Width = UserControl.Width
Background.Height = UserControl.Height

'Riposiziona i pulsanti a seconda di quelli visibili
     If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If


        
        If Picture1.Visible = True And Picture2.Visible = False Then
              HScroll.Width = Picture1.Width
          ElseIf Picture1.Visible = False And Picture2.Visible = True Then
              VScroll.Height = Picture2.Height
         ElseIf Picture1.Visible = True And Picture2.Visible = True Then
              HScroll.Width = Picture1.Width - VScroll.Width
              VScroll.Height = Picture2.Height
        End If

'Aggiorna il gradiente se il controllo viene ridimensionato
UserControl.PropertyChanged Gradient
UserControl.PropertyChanged WindowStyle
UserControl.PropertyChanged BarStyle
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Barra,Barra,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Determines the font to be used in the title bar.\r\n\r\n\r\n"
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = Titolo.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Titolo.Font = New_Font
    PropertyChanged "Font"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Barra,Barra,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns\\sets the color of the text of the title bar."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColor = Titolo.ForeColor
    Barra.ForeColor = ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Titolo.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = ";Misc"
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated context number for an object."
Attribute WhatsThisHelpID.VB_ProcData.VB_Invoke_Property = ";Misc"
    WhatsThisHelpID = m_WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal New_WhatsThisHelpID As Long)
    m_WhatsThisHelpID = New_WhatsThisHelpID
    PropertyChanged "WhatsThisHelpID"
End Property
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ToolTipText = m_def_ToolTipText
    m_WhatsThisHelpID = m_def_WhatsThisHelpID
    m_BarColor = m_def_BarColor
    m_MinEnabled = m_def_MinEnabled
    m_MaxEnabled = m_def_MaxEnabled
    m_QuitEnabled = m_def_QuitEnabled
    m_MinVisible = m_def_MinVisible
    m_MaxVisible = m_def_MaxVisible
    m_QuitVisible = m_def_QuitVisible
    m_BarStyle = m_def_BarStyle
    m_BarVisible = m_def_BarVisible
    m_BackColor = m_def_BackColor
    m_Gradient = m_def_Gradient
    m_Moveable = m_def_Moveable
    m_UseAsForm = m_def_UseAsForm
    m_ScrollBars = m_def_ScrollBars
    m_WindowStyle = m_def_WindowStyle
    m_Caption = m_def_Caption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Titolo.Caption = PropBag.ReadProperty("Caption", "")
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    Titolo.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Barra.BackColor = PropBag.ReadProperty("BarColor", &H80000000)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    m_WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
    m_BarColor = PropBag.ReadProperty("BarColor", m_def_BarColor)
    m_MinEnabled = PropBag.ReadProperty("MinEnabled", m_def_MinEnabled)
    m_MaxEnabled = PropBag.ReadProperty("MaxEnabled", m_def_MaxEnabled)
    m_QuitEnabled = PropBag.ReadProperty("QuitEnabled", m_def_QuitEnabled)
    CMin.Enabled = PropBag.ReadProperty("MinEnabled", True)
    CMax.Enabled = PropBag.ReadProperty("MaxEnabled", True)
    CQuit.Enabled = PropBag.ReadProperty("QuitEnabled", True)
    m_MinVisible = PropBag.ReadProperty("MinVisible", m_def_MinVisible)
    m_MaxVisible = PropBag.ReadProperty("MaxVisible", m_def_MaxVisible)
    m_QuitVisible = PropBag.ReadProperty("QuitVisible", m_def_QuitVisible)
    CMin.Visible = PropBag.ReadProperty("MinVisible", m_def_MinVisible)
    CMax.Visible = PropBag.ReadProperty("MaxVisible", m_def_MaxVisible)
    CQuit.Visible = PropBag.ReadProperty("QuitVisible", m_def_QuitVisible)
    m_BarStyle = PropBag.ReadProperty("BarStyle", m_def_BarStyle)
    m_BarVisible = PropBag.ReadProperty("BarVisible", m_def_BarVisible)
    Barra.Visible = PropBag.ReadProperty("BarVisible", m_def_BarVisible)
   If BarVisible = True And ScrollBars = 0 Then
            Background.Move 0, 300, UserControl.Width, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 1 Then
             Background.Move 0, 300, UserControl.Width, UserControl.Height - 240 - 300
       ElseIf BarVisible = True And ScrollBars = 2 Then
              Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 3 Then
               Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 240 - 300
       ElseIf BarVisible = False And ScrollBars = 0 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 1 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height - 240
       ElseIf BarVisible = False And ScrollBars = 2 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 3 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height - 240
End If
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Background.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Picture2.BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_Gradient = PropBag.ReadProperty("Gradient", m_def_Gradient)
    m_WindowStyle = PropBag.ReadProperty("WindowStyle", m_def_WindowStyle)
    m_Moveable = PropBag.ReadProperty("Moveable", m_def_Moveable)
    Background.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    Background.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 1212)
    Background.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 1296)
    m_UseAsForm = PropBag.ReadProperty("UseAsForm", m_def_UseAsForm)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
    'Riposiziona i pulsanti a seconda di quelli visibili
     If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
    Background.DrawMode = PropBag.ReadProperty("DrawMode", 13)
    Background.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    Background.DrawWidth = PropBag.ReadProperty("DrawWidth", 1)
    Background.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    Background.AutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set Picture = PropBag.ReadProperty("Icon", Nothing)
    Icona.Picture = UserControl.Picture
    Select Case Icona.Picture
       Case 0
          Titolo.Left = 60
       Case Else
          Titolo.Left = 312
    End Select
End Sub

Private Sub UserControl_Show()
RaiseEvent Load

Barra.ZOrder 0
Picture1.ZOrder 0
Picture2.ZOrder 0
Background.ZOrder 1

'Riposiziona i pulsanti a seconda di quelli visibili
     If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
    
  


        If Picture1.Visible = True And Picture2.Visible = False Then
              HScroll.Width = Picture1.Width
          ElseIf Picture1.Visible = False And Picture2.Visible = True Then
              VScroll.Height = Picture2.Height
          ElseIf Picture1.Visible = True And Picture2.Visible = True Then
              HScroll.Width = Picture1.Width - VScroll.Width
              VScroll.Height = Picture2.Height
        End If
        
If BarVisible = True And ScrollBars = 0 Then
            Background.Move 0, 300, UserControl.Width, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 1 Then
             Background.Move 0, 300, UserControl.Width, UserControl.Height - 240 - 300
       ElseIf BarVisible = True And ScrollBars = 2 Then
              Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 3 Then
               Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 240 - 300
       ElseIf BarVisible = False And ScrollBars = 0 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 1 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height - 240
       ElseIf BarVisible = False And ScrollBars = 2 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 3 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height - 240
End If


'Visualizzando il controllo dopo averlo nascosto a run-time, ripristina
'le frecce dello scrolling
If Quitted = True Then
    If SxVisible = True Then
       Picture1.Visible = True
      Else
       Picture1.Visible = False
    End If
    
    If SuVisible = True Then
       Picture2.Visible = True
      Else
       Picture2.Visible = False
    End If
    Quitted = False
End If

'Mostra l'icona
Icona.Picture = UserControl.Picture

'Aggiorna il gradiente e altro se il controllo viene ridimensionato
UserControl.PropertyChanged Gradient
UserControl.PropertyChanged WindowStyle
UserControl.PropertyChanged BarStyle
End Sub

Private Sub UserControl_Terminate()
RaiseEvent Unload
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Titolo.Caption, "")
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", Barra.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
    Call PropBag.WriteProperty("BarColor", m_BarColor, m_def_BarColor)
    Call PropBag.WriteProperty("MinEnabled", m_MinEnabled, m_def_MinEnabled)
    Call PropBag.WriteProperty("MaxEnabled", m_MaxEnabled, m_def_MaxEnabled)
    Call PropBag.WriteProperty("QuitEnabled", m_QuitEnabled, m_def_QuitEnabled)
    Call PropBag.WriteProperty("MinVisible", m_MinVisible, m_def_MinVisible)
    Call PropBag.WriteProperty("MaxVisible", m_MaxVisible, m_def_MaxVisible)
    Call PropBag.WriteProperty("QuitVisible", m_QuitVisible, m_def_QuitVisible)
    Call PropBag.WriteProperty("BarStyle", m_BarStyle, m_def_BarStyle)
    Call PropBag.WriteProperty("BarVisible", m_BarVisible, m_def_BarVisible)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Gradient", m_Gradient, m_def_Gradient)
    Call PropBag.WriteProperty("Moveable", m_Moveable, m_def_Moveable)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 1212)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 1296)
    Call PropBag.WriteProperty("UseAsForm", m_UseAsForm, m_def_UseAsForm)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ScrollBars", m_ScrollBars, m_def_ScrollBars)
    Call PropBag.WriteProperty("WindowStyle", m_WindowStyle, m_def_WindowStyle)
    Call PropBag.WriteProperty("DrawMode", UserControl.DrawMode, 13)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("DrawWidth", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, True)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Icon", Picture, Nothing)
End Sub
Public Property Get MinEnabled() As Boolean
Attribute MinEnabled.VB_Description = "Returns/sets a value that determines whether the minimization button can respond to user-generated events."
Attribute MinEnabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MinEnabled = m_MinEnabled
    CMin.Enabled = MinEnabled
End Property

Public Property Let MinEnabled(ByVal New_MinEnabled As Boolean)
    m_MinEnabled = New_MinEnabled
    PropertyChanged "MinEnabled"
    CMin.Enabled = MinEnabled
End Property

Public Property Get MaxEnabled() As Boolean
Attribute MaxEnabled.VB_Description = "Returns/sets a value that determines whether the maximization button can respond to user-generated events."
Attribute MaxEnabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MaxEnabled = m_MaxEnabled
    CMax.Enabled = MaxEnabled
End Property

Public Property Let MaxEnabled(ByVal New_MaxEnabled As Boolean)
    m_MaxEnabled = New_MaxEnabled
    PropertyChanged "MaxEnabled"
    CMax.Enabled = MaxEnabled
End Property
Public Property Get QuitEnabled() As Boolean
Attribute QuitEnabled.VB_Description = "Returns/sets a value that determines whether the quit button can respond to user-generated events."
Attribute QuitEnabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    QuitEnabled = m_QuitEnabled
    CQuit.Enabled = QuitEnabled
End Property

Public Property Let QuitEnabled(ByVal New_QuitEnabled As Boolean)
    m_QuitEnabled = New_QuitEnabled
    PropertyChanged "QuitEnabled"
    CQuit.Enabled = QuitEnabled
End Property
Public Property Get MinVisible() As Boolean
Attribute MinVisible.VB_Description = "Returns/sets a value that determines whether the minimization button is visible or not."
Attribute MinVisible.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MinVisible = m_MinVisible
    CMin.Visible = MinVisible
    If MinVisible = True Then
       MinEnabled = True
      Else
       MinEnabled = False
    End If
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property

Public Property Let MinVisible(ByVal New_MinVisible As Boolean)
    m_MinVisible = New_MinVisible
    PropertyChanged "MinVisible"
    CMin.Visible = MinVisible
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property

Public Property Get MaxVisible() As Boolean
Attribute MaxVisible.VB_Description = "Returns/sets a value that determines whether the maximization button is visible or not."
Attribute MaxVisible.VB_ProcData.VB_Invoke_Property = ";Behavior"
    MaxVisible = m_MaxVisible
    CMax.Visible = MaxVisible
    If MaxVisible = True Then
       MaxEnabled = True
      Else
       MaxEnabled = False
    End If
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property

Public Property Let MaxVisible(ByVal New_MaxVisible As Boolean)
    m_MaxVisible = New_MaxVisible
    PropertyChanged "MaxVisible"
    CMax.Visible = MaxVisible
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property

Public Property Get QuitVisible() As Boolean
Attribute QuitVisible.VB_Description = "Returns/sets a value that determines whether the quit button is visible or not."
Attribute QuitVisible.VB_ProcData.VB_Invoke_Property = ";Behavior"
    QuitVisible = m_QuitVisible
    CQuit.Visible = QuitVisible
    If QuitVisible = True Then
       QuitEnabled = True
      Else
       QuitEnabled = False
    End If
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property

Public Property Let QuitVisible(ByVal New_QuitVisible As Boolean)
    m_QuitVisible = New_QuitVisible
    PropertyChanged "QuitVisible"
    CQuit.Visible = QuitVisible
    'Riposiziona i pulsanti a seconda di quelli visibili
    If CMin.Visible = True And CMax.Visible = True And CQuit.Visible = False Then
             CMax.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMin.Left = UserControl.Width - 456 - 100
        ElseIf CMin.Visible = True And CMax.Visible = True And CQuit.Visible = True Then
             CQuit.Left = UserControl.Width - 228 - 100
             CMax.Left = UserControl.Width - 456 - 100
             CMin.Left = UserControl.Width - 684 - 100
        ElseIf CMin.Visible = True And CMax.Visible = False And CQuit.Visible = False Then
              CMin.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = True And CQuit.Visible = False Then
              CMax.Left = UserControl.Width - 228 - 100
        ElseIf CMin.Visible = False And CMax.Visible = False And CQuit.Visible = True Then
              CQuit.Left = UserControl.Width - 228 - 100
    End If
End Property
Private Sub UserControl_Hide()
    RaiseEvent Unload
End Sub
Public Property Get BarStyle() As Integer
Attribute BarStyle.VB_Description = "Returns/sets a value that determines the three-dimensional style of the title bar (0=Flat; 1= Inset; 2=Raised)."
Attribute BarStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarStyle = m_BarStyle
    Select Case BarStyle
       Case 0
          Barra.Cls
          Barra.BorderStyle = 1
       Case 1
          Barra.BorderStyle = 0
          Make3DBarra 0, 0, Barra.Width - 10, Barra.Height - 10, 2, 1
       Case 2
          Barra.BorderStyle = 0
          Make3DBarra 0, 0, Barra.Width - 10, Barra.Height - 10, 2, 2
       Case Is > 2
          Barra.Cls
          Barra.BorderStyle = 1
    End Select
End Property

Public Property Let BarStyle(ByVal New_BarStyle As Integer)
    m_BarStyle = New_BarStyle
    PropertyChanged "BarStyle"
   Select Case BarStyle
       Case 0
          Barra.Cls
          Barra.BorderStyle = 1
       Case 1
          Barra.BorderStyle = 0
          Make3DBarra 0, 0, Barra.Width - 10, Barra.Height - 10, 2, 1
       Case 2
          Barra.BorderStyle = 0
          Make3DBarra 0, 0, Barra.Width - 10, Barra.Height - 10, 2, 2
       Case Is > 2
          Barra.Cls
          Barra.BorderStyle = 1
    End Select
End Property
Public Property Get BarVisible() As Boolean
Attribute BarVisible.VB_Description = "Sets\\returns a value that determines if the title bar is visible or not."
Attribute BarVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarVisible = m_BarVisible
    Barra.Visible = m_BarVisible
    
   If BarVisible = True And ScrollBars = 0 Then
            Background.Move 0, 300, UserControl.Width, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 1 Then
             Background.Move 0, 300, UserControl.Width, UserControl.Height - 240 - 300
       ElseIf BarVisible = True And ScrollBars = 2 Then
              Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 3 Then
               Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 240 - 300
       ElseIf BarVisible = False And ScrollBars = 0 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 1 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height - 240
       ElseIf BarVisible = False And ScrollBars = 2 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 3 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height - 240
End If
End Property
Public Property Let BarVisible(ByVal New_BarVisible As Boolean)
    m_BarVisible = New_BarVisible
    PropertyChanged "BarVisible"
    Barra.Visible = m_BarVisible
    
   If BarVisible = True And ScrollBars = 0 Then
            Background.Move 0, 300, UserControl.Width, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 1 Then
             Background.Move 0, 300, UserControl.Width, UserControl.Height - 240 - 300
       ElseIf BarVisible = True And ScrollBars = 2 Then
              Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 300
       ElseIf BarVisible = True And ScrollBars = 3 Then
               Background.Move 0, 300, UserControl.Width - 240, UserControl.Height - 240 - 300
       ElseIf BarVisible = False And ScrollBars = 0 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 1 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width, UserControl.Height - 240
       ElseIf BarVisible = False And ScrollBars = 2 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height
       ElseIf BarVisible = False And ScrollBars = 3 Then
               UserControl.Picture = LoadPicture("")
               Background.Move 0, 0, UserControl.Width - 240, UserControl.Height - 240
End If
End Property



Sub ShadeControl(Col1 As Integer, Col2 As Integer)
Attribute ShadeControl.VB_Description = "This method allows you to apply a shading effect to the background of the control at run time. "
    Dim DS, DW, SM, SH
    Dim i As Integer
    'Colora il controllo di colori diversi ad ogni apertura
    Randomize
    Const Inside_Solid = 6
    Const Copy_Pen = 13

    DS = Background.DrawStyle                  'save 'em
    DW = Background.DrawWidth
    SM = Background.ScaleMode
    SH = Background.ScaleHeight
    Background.DrawStyle = vbInsideSolid
    Background.DrawWidth = 2
    Background.ScaleMode = vbPixels
    Background.ScaleHeight = 256
    For i = 0 To 255
               Background.Line (0, i)-(Background.Width, i + 1), _
                  RGB(Col1, Col2, 255 - i), B
    Next i
    Background.DrawStyle = DS          'restore the settings
    Background.DrawWidth = DW
    Background.ScaleHeight = SH        'must be restored before ScaleMode
    Background.ScaleMode = SM
End Sub
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Sets\\returns a value that determines the background color of the control."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = m_BackColor
    UserControl.BackColor = BackColor
    Picture1.BackColor = BackColor
    Picture2.BackColor = BackColor
    Background.BackColor = BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = BackColor
    Picture1.BackColor = BackColor
    Picture2.BackColor = BackColor
    Background.BackColor = BackColor
End Property

Public Property Get BarColor() As OLE_COLOR
Attribute BarColor.VB_Description = "Sets\\returns a value that determines the background color of the title bar."
Attribute BarColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BarColor = m_BarColor
    Barra.BackColor = BarColor
End Property

Public Property Let BarColor(ByVal New_BarColor As OLE_COLOR)
    m_BarColor = New_BarColor
    PropertyChanged "BarColor"
    Barra.BackColor = BarColor
End Property

Public Property Get Gradient() As Boolean
Attribute Gradient.VB_Description = "Sets\\returns a boolean value that allows you to apply a shading effect to the background of the control. Not available at run-time."
Attribute Gradient.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Gradient = m_Gradient
    Select Case Gradient
       Case True
            Dim DS, DW, SM, SH
            Dim i As Integer

            'Colora il controllo di colori diversi ad ogni apertura
            Randomize
            C1 = Int((255 * Rnd) + 0)
            C2 = Int((255 * Rnd) + 0)
            Const Inside_Solid = 6
            Const Copy_Pen = 13
        
            DS = Background.DrawStyle                  'save 'em
            DW = Background.DrawWidth
            SM = Background.ScaleMode
            SH = Background.ScaleHeight
            Background.DrawStyle = vbInsideSolid
            Background.DrawWidth = 2
            Background.ScaleMode = vbPixels
            Background.ScaleHeight = 256
            For i = 0 To 255
                       Background.Line (0, i)-(Background.Width, i + 1), _
                          RGB(C1, C2, 255 - i), B
            Next i
            Background.DrawStyle = DS          'restore the settings
            Background.DrawWidth = DW
            Background.ScaleHeight = SH        'must be restored before ScaleMode
            Background.ScaleMode = SM
       Case False
          'Background.Cls
          Background.BackColor = BackColor
    End Select
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
   If UserControl.Ambient.UserMode = True Then
      MsgBox ("Property not available at run-time! " & Chr(13) & "" & Chr(10) & "Use method ShadeControl instead.")
    Else
      m_Gradient = New_Gradient
      PropertyChanged "Gradient"
   End If
End Property
Public Property Get Moveable() As Boolean
Attribute Moveable.VB_Description = "Returns\\sets a value that determines if the control can be moved within its container.Read only at run time."
Attribute Moveable.VB_ProcData.VB_Invoke_Property = ";Position"
    Moveable = m_Moveable
End Property

Public Property Let Moveable(ByVal New_Moveable As Boolean)
    m_Moveable = New_Moveable
    PropertyChanged "Moveable"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
    MousePointer = Background.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Background.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
Attribute ScaleHeight.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleHeight = Background.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    Background.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
Attribute ScaleWidth.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleWidth = Background.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    Background.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

Public Property Get UseAsForm() As Boolean
Attribute UseAsForm.VB_Description = "Sets a value that lets the user to use the miniForm control as a Form. Only design time."
Attribute UseAsForm.VB_ProcData.VB_Invoke_Property = ";Misc"
    UseAsForm = m_UseAsForm
    If TypeOf UserControl.Extender.Container Is Form Then
       If UseAsForm = True Then
           On Error Resume Next
           UserControl.Extender.Parent.BorderStyle = 5
           UserControl.Extender.Parent.Caption = ""
           UserControl.Extender.Parent.ControlBox = False
           UserControl.Extender.Top = 0
           UserControl.Extender.Left = 0
           UserControl.Extender.Width = UserControl.Extender.Parent.Width - 100
           UserControl.Extender.Height = UserControl.Extender.Parent.Height - 100
       End If
    End If
End Property

Public Property Let UseAsForm(ByVal New_UseAsForm As Boolean)
    If UserControl.Ambient.UserMode Then
      MsgBox ("Property not available at run-time!")
     Else
      m_UseAsForm = New_UseAsForm
      PropertyChanged "UseAsForm"
    End If
End Property

Private Sub VScroll_Change()
Static Valore As Long
Select Case VScroll.Value
    Case Is > Valore
            If UserControl.Ambient.UserMode = True Then
                    For X = 0 To UserControl.ContainedControls.Count - 1
                        On Error Resume Next
                        UserControl.ContainedControls(X).Top = UserControl.ContainedControls(X).Top + VScroll.SmallChange
                        Valore = VScroll.Value
                    Next
            End If
    Case Is < Valore
            If UserControl.Ambient.UserMode = True Then
                    For X = 0 To UserControl.ContainedControls.Count - 1
                        On Error Resume Next
                        UserControl.ContainedControls(X).Top = UserControl.ContainedControls(X).Top - VScroll.SmallChange
                        Valore = VScroll.Value
                    Next
            End If
End Select
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of the control. "
     Barra.Refresh
     UserControl.Refresh
     Background.Refresh
End Sub
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "Sets\\returns a value that  determines if scrollbars are visible or not. Valid values:0-None,1-Horizontal,\r\n2-Vertical,3-Both.\r\n"
Attribute ScrollBars.VB_ProcData.VB_Invoke_Property = ";Misc"
    ScrollBars = m_ScrollBars
    Select Case ScrollBars
       Case 0
         UserControl.Picture1.Visible = False
         UserControl.Picture2.Visible = False
       Case 1
         UserControl.Picture1.Visible = True
         UserControl.Picture2.Visible = False
       Case 2
         UserControl.Picture1.Visible = False
         UserControl.Picture2.Visible = True
       Case 3
         UserControl.Picture1.Visible = True
         UserControl.Picture2.Visible = True
       Case Else
         ScrollBars = 0
         Exit Property
    End Select
    
    If Picture1.Visible = True And Picture2.Visible = False Then
              HScroll.Width = Picture1.Width
          ElseIf Picture1.Visible = False And Picture2.Visible = True Then
              VScroll.Height = Picture2.Height
         ElseIf Picture1.Visible = True And Picture2.Visible = True Then
              HScroll.Width = Picture1.Width - VScroll.Width
              VScroll.Height = Picture2.Height
    End If
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As Integer)
    m_ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
    Select Case ScrollBars
       Case 0
         UserControl.Picture1.Visible = False
         UserControl.Picture2.Visible = False
       Case 1
         UserControl.Picture1.Visible = True
         UserControl.Picture2.Visible = False
       Case 2
         UserControl.Picture1.Visible = False
         UserControl.Picture2.Visible = True
       Case 3
         UserControl.Picture1.Visible = True
         UserControl.Picture2.Visible = True
       Case Else
         ScrollBars = 0
         Exit Property
    End Select
    
    If Picture1.Visible = True And Picture2.Visible = False Then
              HScroll.Width = Picture1.Width
          ElseIf Picture1.Visible = False And Picture2.Visible = True Then
              VScroll.Height = Picture2.Height
         ElseIf Picture1.Visible = True And Picture2.Visible = True Then
              HScroll.Width = Picture1.Width - VScroll.Width
              VScroll.Height = Picture2.Height
    End If
End Property
Public Property Get WindowStyle() As Integer
Attribute WindowStyle.VB_Description = "Returns/sets a value that determines the three-dimensional style of the control window (0=Flat; 1= Inset; 2=Raised)."
Attribute WindowStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    WindowStyle = m_WindowStyle
            Select Case WindowStyle
                 Case 0
                   If Gradient = False Then
                     Background.BackColor = BackColor
                   End If
                 Case 1
                   Make3D 0, 0, Background.Width - 60, Background.Height - 60, 1, 1
                 Case 2
                   Make3D 0, 0, Background.Width - 60, Background.Height - 60, 1, 2
                 Case Is > 2
                   WindowStyle = 0
           End Select
End Property

Public Property Let WindowStyle(ByVal New_WindowStyle As Integer)
    m_WindowStyle = New_WindowStyle
    PropertyChanged "WindowStyle"
    
         Select Case WindowStyle
                 Case 0
                   If Gradient = False Then
                     Background.BackColor = BackColor
                   End If
                 Case 1
                   Make3D 0, 0, Background.Width - 60, Background.Height - 60, 1, 1
                 Case 2
                   Make3D 0, 0, Background.Width - 60, Background.Height - 60, 1, 2
                 Case Is > 2
                   WindowStyle = 0
           End Select
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawMode
Public Property Get DrawMode() As Integer
Attribute DrawMode.VB_Description = "Returns or sets a value that determines the appearance of output from graphics method or the appearance of a Shape or Line control.\r\n"
Attribute DrawMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DrawMode = Background.DrawMode
End Property

Public Property Let DrawMode(ByVal New_DrawMode As Integer)
    Background.DrawMode() = New_DrawMode
    PropertyChanged "DrawMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Returns or sets a value that determines the line style for output from graphics methods."
Attribute DrawStyle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DrawStyle = Background.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    Background.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawWidth
Public Property Get DrawWidth() As Integer
Attribute DrawWidth.VB_Description = "Returns or sets the line width for output from graphics methods."
Attribute DrawWidth.VB_ProcData.VB_Invoke_Property = ";Behavior"
    DrawWidth = Background.DrawWidth
End Property

Public Property Let DrawWidth(ByVal New_DrawWidth As Integer)
    Background.DrawWidth() = New_DrawWidth
    PropertyChanged "DrawWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns or sets a value indicating the unit of measurement for coordinates of an object when using graphics methods or when positioning controls.\r\n"
Attribute ScaleMode.VB_ProcData.VB_Invoke_Property = ";Scale"
    ScaleMode = Background.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    Background.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

Public Function Make3D(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Bordo As Integer, Tipo As Single)
Attribute Make3D.VB_MemberFlags = "40"
    Dim DS, DW, SM
    On Error GoTo errore
    
    DS = Background.DrawStyle
    DW = Background.DrawWidth
    SM = Background.ScaleMode
    
    Background.DrawStyle = vbSolid
    Background.DrawWidth = 1
    Background.ScaleMode = vbTwips
    
    Bordo = Bordo * 20

    If Bordo = 0 Then
       Bordo = 50
    End If

    Select Case Tipo
      Case 0    'Normale
        Background.Cls
      Case 1    'Inset
        'Trapezio superiore
        For k = 0 To Bordo
          Background.Line (x1, y1 + k)-(x2 - k, y1 + k), RGB(128, 128, 128)
        Next
        
        'Trapezio destro
        For k = 0 To Bordo
           Background.Line (x2 - k, y1 + k)-(x2 - k, y2), RGB(255, 255, 255)
        Next
        
        'Trapezio lat. sinistro
        For k = 0 To Bordo
             Background.Line (x1 + k, y1)-(x1 + k, y2 - k), RGB(128, 128, 128)
        Next
        
        'Trapezio inferiore
        For k = 0 To Bordo
          Background.Line (x1 + k, y2 - k)-(x2, y2 - k), RGB(255, 255, 255)
        Next
      Case 2    'Raised

        'Trapezio superiore
        For k = 0 To Bordo
          Background.Line (x1, y1 + k)-(x2 - k, y1 + k), RGB(255, 255, 255)
        Next

        'Trapezio destro
        For k = 0 To Bordo
           Background.Line (x2 - k, y1 + k)-(x2 - k, y2), RGB(128, 128, 128)
        Next

        'Trapezio lat. sinistro
        For k = 0 To Bordo
             Background.Line (x1 + k, y1)-(x1 + k, y2 - k), RGB(255, 255, 255)
        Next

        'Trapezio inferiore
        For k = 0 To Bordo
          Background.Line (x1 + k, y2 - k)-(x2, y2 - k), RGB(128, 128, 128)
        Next

      Case Else 'Errore
        MsgBox ("Value out of range!")
        Background.Cls
    End Select
    
    Background.DrawStyle = DS          'restore the settings
    Background.DrawWidth = DW
    Background.ScaleMode = SM

    Exit Function

errore:
MsgBox ("Error n°" & Err.Number & "." & Chr(13) & "" & Chr(10) & "" & Err.Description & ".")
Exit Function
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns or sets the output from a graphics method to a persistent graphic."
Attribute AutoRedraw.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoRedraw = Background.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    Background.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Function Make3DBarra(x1 As Long, y1 As Long, x2 As Long, y2 As Long, Bordo As Integer, Tipo As Single)
Attribute Make3DBarra.VB_MemberFlags = "40"
    Dim DS, DW, SM
    On Error GoTo errore
    
    DS = Barra.DrawStyle
    DW = Barra.DrawWidth
    SM = Barra.ScaleMode
    
   Barra.DrawStyle = vbSolid
   Barra.DrawWidth = 1
   Barra.ScaleMode = vbTwips
    
    Bordo = Bordo * 20

    If Bordo = 0 Then
       Bordo = 50
    End If

    Select Case Tipo
      Case 0    'Normale
       Barra.Cls
      Case 1    'Inset
        'Trapezio superiore
        For k = 0 To Bordo
         Barra.Line (x1, y1 + k)-(x2 - k, y1 + k), RGB(128, 128, 128)
        Next
        
        'Trapezio destro
        For k = 0 To Bordo
          Barra.Line (x2 - k, y1 + k)-(x2 - k, y2), RGB(255, 255, 255)
        Next
        
        'Trapezio lat. sinistro
        For k = 0 To Bordo
            Barra.Line (x1 + k, y1)-(x1 + k, y2 - k), RGB(128, 128, 128)
        Next
        
        'Trapezio inferiore
        For k = 0 To Bordo
         Barra.Line (x1 + k, y2 - k)-(x2, y2 - k), RGB(255, 255, 255)
        Next
      Case 2    'Raised

        'Trapezio superiore
        For k = 0 To Bordo
         Barra.Line (x1, y1 + k)-(x2 - k, y1 + k), RGB(255, 255, 255)
        Next

        'Trapezio destro
        For k = 0 To Bordo
          Barra.Line (x2 - k, y1 + k)-(x2 - k, y2), RGB(128, 128, 128)
        Next

        'Trapezio lat. sinistro
        For k = 0 To Bordo
            Barra.Line (x1 + k, y1)-(x1 + k, y2 - k), RGB(255, 255, 255)
        Next

        'Trapezio inferiore
        For k = 0 To Bordo
         Barra.Line (x1 + k, y2 - k)-(x2, y2 - k), RGB(128, 128, 128)
        Next

      Case Else 'Errore
        MsgBox ("Value out of range!")
       Barra.Cls
    End Select
    
   Barra.DrawStyle = DS          'restore the settings
   Barra.DrawWidth = DW
   Barra.ScaleMode = SM

    Exit Function

errore:
MsgBox ("Error n°" & Err.Number & "." & Chr(13) & "" & Chr(10) & "" & Err.Description & ".")
Exit Function
End Function

Private Sub VScroll_Scroll()
Static Valore As Long
Select Case VScroll.Value
    Case Is > Valore
            If UserControl.Ambient.UserMode = True Then
                    For X = 0 To UserControl.ContainedControls.Count - 1
                        On Error Resume Next
                        UserControl.ContainedControls(X).Top = UserControl.ContainedControls(X).Top + VScroll.LargeChange
                        Valore = VScroll.Value
                    Next
            End If
    Case Is < Valore
            If UserControl.Ambient.UserMode = True Then
                    For X = 0 To UserControl.ContainedControls.Count - 1
                        On Error Resume Next
                        UserControl.ContainedControls(X).Top = UserControl.ContainedControls(X).Top - VScroll.LargeChange
                        Valore = VScroll.Value
                    Next
            End If
End Select
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Determines the text to be displayed in the title bar."
    Caption = m_Caption
    Titolo.Caption = Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Titolo.Caption = Caption
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Icon() As Picture
Attribute Icon.VB_Description = "Sets an icon to be displayed in the title bar."
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Icon = UserControl.Picture
    Icona.Picture = UserControl.Picture
    Select Case Icona.Picture
       Case 0
          Titolo.Left = 60
       Case Else
          Titolo.Left = 312
    End Select
End Property

Public Property Set Icon(ByVal New_Icon As Picture)
    Set UserControl.Picture = New_Icon
    PropertyChanged "Icon"
    Icona.Picture = UserControl.Picture
    Select Case Icona.Picture
       Case 0
          Titolo.Left = 60
       Case Else
          Titolo.Left = 312
    End Select
End Property

