VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ascii"
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   -75
      Top             =   765
   End
   Begin VB.ComboBox cmbFontName 
      Height          =   315
      Left            =   225
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   30
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.PictureBox Picture1 
      Height          =   4020
      Left            =   90
      ScaleHeight     =   3960
      ScaleWidth      =   9465
      TabIndex        =   0
      Top             =   915
      Width           =   9525
      Begin VB.Label lblSym 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   210
         Width           =   315
      End
      Begin VB.Label lblAsc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "32"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.Shape Shape16 
      Height          =   465
      Left            =   240
      Top             =   420
      Width           =   2730
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Scroll"
      Height          =   180
      Left            =   2235
      TabIndex        =   17
      Top             =   555
      Width           =   450
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Caps"
      Height          =   210
      Left            =   1260
      TabIndex        =   16
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Num"
      Height          =   165
      Left            =   270
      TabIndex        =   15
      Top             =   540
      Width           =   330
   End
   Begin VB.Shape sScroll 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2685
      Shape           =   3  'Circle
      Top             =   525
      Width           =   255
   End
   Begin VB.Shape sCaps 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   1665
      Shape           =   3  'Circle
      Top             =   525
      Width           =   255
   End
   Begin VB.Shape sNum 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   645
      Shape           =   3  'Circle
      Top             =   525
      Width           =   255
   End
   Begin VB.Shape Shape12 
      Height          =   870
      Left            =   3090
      Top             =   15
      Width           =   3735
   End
   Begin VB.Label lblShift 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   6960
      TabIndex        =   14
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "For KeyCode use keyboard"
      Height          =   195
      Left            =   3180
      TabIndex        =   13
      Top             =   630
      Width           =   2115
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Click on chart or use keyboard to select character"
      Height          =   210
      Left            =   3180
      TabIndex        =   12
      Top             =   450
      Width           =   3615
   End
   Begin VB.Shape Shape10 
      Height          =   870
      Left            =   7875
      Top             =   15
      Width           =   1635
   End
   Begin VB.Shape Shape9 
      Height          =   870
      Left            =   6930
      Top             =   15
      Width           =   795
   End
   Begin VB.Label cmdChangeFont 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Change Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3135
      TabIndex        =   11
      Top             =   60
      Width           =   1545
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6075
      TabIndex        =   10
      Top             =   60
      Width           =   675
   End
   Begin VB.Label cmdAscii 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Show Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4755
      TabIndex        =   9
      Top             =   60
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ascii/Hex"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7905
      TabIndex        =   8
      Top             =   465
      Width           =   840
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "KeyCode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6945
      TabIndex        =   7
      Top             =   45
      Width           =   750
   End
   Begin VB.Label lblKeyCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7050
      TabIndex        =   6
      Top             =   240
      Width           =   450
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   7920
      TabIndex        =   5
      Top             =   60
      Width           =   705
   End
   Begin VB.Shape Shape5 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4035
      Left            =   195
      Top             =   1005
      Width           =   9510
   End
   Begin VB.Shape sFontShadow 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   285
      Top             =   90
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   4815
      Top             =   75
      Width           =   1200
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   6120
      Top             =   90
      Width           =   660
   End
   Begin VB.Label lblChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   8790
      TabIndex        =   4
      Top             =   45
      Width           =   630
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   765
      Left            =   8850
      Top             =   90
      Width           =   600
   End
   Begin VB.Shape Shape6 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   345
      Left            =   7965
      Top             =   105
      Width           =   690
   End
   Begin VB.Shape Shape7 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7095
      Top             =   270
      Width           =   435
   End
   Begin VB.Shape Shape8 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   285
      Left            =   3180
      Top             =   120
      Width           =   1530
   End
   Begin VB.Shape Shape11 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   210
      Left            =   7005
      Top             =   630
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
    '****************************************************
    '*
    '*         Project Name : Ascii/KeyCodes Chart
    '*        Version Number: 1.0.1
    '*           Author Name: Ken Foster
    '*                 Date : February 23, 2007
    '*        Freeware - Use anyway you want.
    '*
    '****************************************************
    '***************** Table of Procedures *************
    '   Private Sub Form_Load
    '   Private Sub Form_Unload
    '   Private Sub cmdAscii_Click
    '   Private Sub cmdClose_Click
    '   Private Sub cmbFontName_Click
    '   Private Sub cmdChangeFont_Click
    '   Private Sub lblAsc_Click
    '   Private Sub lblSym_Click
    '   Private Sub lblSym_MouseDown
    '   Private Sub LoadFonts
    '   Private Sub SetHighlight
    '   Private Sub Picture1_KeyDown
    '   Private Sub Picture1_KeyPress
    '   Private Sub Timer1_Timer
    '***************** End of Table ********************
    
    Private Declare Function GetKeyboardState Lib "user32" _
    (pbKeyState As Byte) As Long
    Private Declare Function SetKeyboardState Lib "user32" _
    (lppbKeyState As Byte) As Long
    
    ' Constant declarations:
    Const VK_NUMLOCK = &H90
    Const VK_SCROLL = &H91
    Const VK_CAPITAL = &H14
    
    Private Const NormBColr As Long = vbWindowBackground
    Private Const NormFColr As Long = vbWindowText
    Private Const HLBackColor As Long = &H808080
    Private Const HLForeColor As Long = vbHighlightText
    Private Const SMFontSize As Integer = 8
    
    Private LastSel As Integer        'stores last label index
    Private aschex As Boolean         'indicates if ascii or hex is shown

Private Sub Form_Load()
    Dim x As Integer
    
    Picture1.Width = 9525
    Picture1.Height = 4020
    
    'load labels
    For x = 1 To 223
        'first row
        Load lblAsc(x)
        lblAsc(x).Left = lblAsc(x - 1).Left + lblAsc(x - 1).Width - 20
        lblAsc(x).Top = lblAsc(x - 1).Top
        lblAsc(x).Visible = True
        lblAsc(x).Caption = x + 32
        Load lblSym(x)
        lblSym(x).Left = lblSym(x - 1).Left + lblSym(x - 1).Width - 20
        lblSym(x).Top = lblSym(x - 1).Top
        lblSym(x).Visible = True
        'second row
        If x = 32 Then
            lblAsc(x).Top = lblAsc(x - 1).Height + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblSym(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
        'third row
        If x = 64 Then
            lblAsc(x).Top = lblAsc(x - 1).Top + (lblAsc(x - 1).Height) + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblAsc(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
        'forth row
        If x = 96 Then
            lblAsc(x).Top = lblAsc(x - 1).Top + (lblAsc(x - 1).Height) + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblAsc(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
        'fifth row
        If x = 128 Then
            lblAsc(x).Top = lblAsc(x - 1).Top + (lblAsc(x - 1).Height) + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblAsc(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
        'sixth row
        If x = 160 Then
            lblAsc(x).Top = lblAsc(x - 1).Top + (lblAsc(x - 1).Height) + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblAsc(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
        'seventh row
        If x = 192 Then
            lblAsc(x).Top = lblAsc(x - 1).Top + (lblAsc(x - 1).Height) + lblSym(x - 1).Height - 20
            lblAsc(x).Left = lblAsc(0).Left
            lblSym(x).Left = lblAsc(0).Left
            lblSym(x).Top = lblAsc(x).Top + (lblAsc(x - 1).Height) - 20
        End If
    Next x
    aschex = False
    
    LastSel = 0
    LoadFonts
    cmbFontName.Text = lblSym(0).FontName    'set default fontname
    cmbFontName_Click                        'load labels with font characters
    lblSym_Click 1                           'display a default character on startup
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub cmdAscii_Click()
    Dim x As Integer
    
    If aschex = True Then                         'show ascii
    For x = 0 To 223
        lblAsc(x).Caption = x + 32
    Next x
    aschex = False
    cmdAscii.Caption = "Show Hex"
    Me.Caption = "Ascii"
    lblCode.Caption = lblAsc(LastSel).Caption
Else                                              'show hex
    For x = 0 To 223
        lblAsc(x).Caption = Hex(lblAsc(x).Caption)
    Next x
    aschex = True
    cmdAscii.Caption = "Show Ascii"
    Me.Caption = "Hex"
    lblCode.Caption = lblAsc(LastSel).Caption
End If
End Sub

Private Sub cmdClose_Click()
    Timer1.Enabled = False
    Unload Me
End Sub

Private Sub cmbFontName_Click()
    Dim x As Integer
    
    'Loop for all of the character labels
    For x = lblSym.LBound To lblSym.uBound
        lblSym(x).FontName = cmbFontName.Text
        lblSym(x).Caption = Chr(32 + x)
        If lblSym(x).FontSize <> SMFontSize Then lblSym(x).FontSize = SMFontSize
    Next x
    lblChar.Font = cmbFontName.Text
    'hide combobox and shape so they can't get the focus
    'pressing the Tab button will not show keycode if they are visible
    cmbFontName.Visible = False
    sFontShadow.Visible = False
End Sub

Private Sub cmdChangeFont_Click()
    'make invisible so when Tab key is pressed it shows the Keycode
    'can't have any control that takes the focus, for this to work
    'thats why I'm using labels for buttons
    cmbFontName.Visible = Not cmbFontName.Visible
    sFontShadow.Visible = Not sFontShadow.Visible
End Sub

Private Sub lblAsc_Click(Index As Integer)
    lblSym_Click Index
End Sub


Private Sub lblSym_Click(Index As Integer)
    lblChar.Font = lblSym(Index).Font
    lblChar.Caption = lblSym(Index).Caption
    lblCode.Caption = lblAsc(Index).Caption
    SetHighlight Index
    lblKeyCode.Caption = ""
    lblShift.Caption = ""
End Sub

Private Sub lblSym_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbLeftButton Then Call SetHighlight(Index)
End Sub

Private Sub LoadFonts()
    Dim x As Integer
    Dim FontType As Integer
    
    'Loop for all fonts on user's system
    For x = 0 To Screen.FontCount - 1
        'FontType = 1
        cmbFontName.AddItem Screen.Fonts(x)
    Next x
End Sub

Private Sub SetHighlight(ByVal Index As Integer)
    
    If LastSel <> Index Then
        'Set the back & fore color of the last highlighted label back to normal
        If LastSel > -1 Then
            lblSym(LastSel).BackColor = NormBColr
            lblSym(LastSel).ForeColor = NormFColr
        End If
        'Set this label's  back & fore color to highlighted
        lblSym(Index).BackColor = HLBackColor
        lblSym(Index).ForeColor = HLForeColor
        LastSel = Index
    End If
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 32 Then
        lblChar.Caption = ""
        lblCode.Caption = ""
    End If
    If Shift = 0 Then lblShift.Caption = ""
    If Shift = 1 Then lblShift.Caption = "Shift"
    If Shift = 2 Then lblShift.Caption = "Control"
    If Shift = 4 Then
       lblShift.Caption = "Alt"
       KeyCode = 0                     ' this keeps the Alt event from firing when Alt is pressed
       lblKeyCode.Caption = "18"       ' show keycode in window
       Exit Sub
    End If
    lblKeyCode.Caption = KeyCode
End Sub

Private Sub Picture1_KeyPress(KeyAscii As Integer)
   Dim x As Integer
    If KeyAscii < 32 Then
        lblChar.Caption = ""
        lblCode.Caption = ""
    End If
    
    'highlight the key that was pressed
    For x = 0 To 223
       If aschex = False Then
         If lblAsc(x).Caption = KeyAscii Then        'ascii mode
             SetHighlight x
             lblChar.Caption = Chr$(KeyAscii)
             lblCode.Caption = KeyAscii
          End If
        Else
         If lblAsc(x).Caption = Hex(KeyAscii) Then   'hex mode
             SetHighlight x
             lblChar.Caption = Chr$(KeyAscii)
             lblCode.Caption = Hex(KeyAscii)
         End If
        End If
    Next x
End Sub

Private Sub Timer1_Timer()
    Dim NumLockState As Boolean
    Dim ScrollLockState As Boolean
    Dim CapsLockState As Boolean
    Dim keys(0 To 255) As Byte
    GetKeyboardState keys(0)
    
    ' NumLock handling:
    NumLockState = keys(VK_NUMLOCK)
    SetKeyboardState keys(0)
    If NumLockState = True Then
        sNum.FillColor = &H80FF80
    Else
        sNum.FillColor = vbWhite
    End If
    
    ' CapsLock handling:
    CapsLockState = keys(VK_CAPITAL)
    SetKeyboardState keys(0)
    If CapsLockState = True Then
        sCaps.FillColor = &H80FF80
    Else
        sCaps.FillColor = vbWhite
    End If
    
    ' ScrollLock handling:
    ScrollLockState = keys(VK_SCROLL)
    SetKeyboardState keys(0)
    If ScrollLockState = True Then
        sScroll.FillColor = &H80FF80
    Else
        sScroll.FillColor = vbWhite
    End If
End Sub
