VERSION 5.00
Begin VB.Form frmPanel 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1845
   ClientLeft      =   5265
   ClientTop       =   3495
   ClientWidth     =   5265
   ControlBox      =   0   'False
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
   Icon            =   "frmPanel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "frmPanel.frx":030A
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   351
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4980
      Picture         =   "frmPanel.frx":3928C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1545
      Width           =   240
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4980
      Picture         =   "frmPanel.frx":3938E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   45
      Width           =   240
   End
   Begin VB.CommandButton cmdHide 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cacher"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   22
      Top             =   675
      Width           =   630
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   9
      Left            =   4380
      Picture         =   "frmPanel.frx":39490
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   7
      Left            =   3390
      Picture         =   "frmPanel.frx":39642
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   6
      Left            =   2895
      Picture         =   "frmPanel.frx":397E4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   5
      Left            =   2400
      Picture         =   "frmPanel.frx":39996
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   4
      Left            =   1905
      Picture         =   "frmPanel.frx":39B38
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   3
      Left            =   1410
      Picture         =   "frmPanel.frx":39CEA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   2
      Left            =   915
      Picture         =   "frmPanel.frx":39E9C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   1
      Left            =   420
      Picture         =   "frmPanel.frx":3A04E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1425
      Width           =   480
   End
   Begin VB.CommandButton Command1 
      Height          =   360
      Index           =   8
      Left            =   3885
      Picture         =   "frmPanel.frx":3A200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1425
      Width           =   480
   End
   Begin Cathodic.cSysTray SysTray 
      Left            =   2400
      Top             =   2055
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmPanel.frx":3A3B2
      TrayTip         =   "Double cliquez pour afficher..."
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      Picture         =   "frmPanel.frx":3A6CC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1545
      Width           =   240
   End
   Begin VB.CommandButton cmdQuit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Power"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4545
      TabIndex        =   8
      Top             =   675
      Width           =   630
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      Height          =   930
      Left            =   840
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   237
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   345
      Width           =   3615
      Begin VB.OptionButton optAffichage 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   525
         Width           =   2250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   4
         Top             =   540
         Width           =   2250
      End
      Begin VB.CommandButton cmdAlterner 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Alterner"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   360
         TabIndex        =   0
         Top             =   300
         Width           =   705
      End
      Begin VB.OptionButton optAffichage 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   285
         Width           =   2250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   240
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   300
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.OptionButton optAffichage 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   5
         Top             =   30
         Value           =   -1  'True
         Width           =   2250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   2250
      End
      Begin VB.CheckBox chkActiver 
         BackColor       =   &H00808080&
         Caption         =   "Échelles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   165
         TabIndex        =   1
         Top             =   330
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblAffichage 
         BackStyle       =   0  'Transparent
         Caption         =   "Rien à modifier"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   1425
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   1680
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      Picture         =   "frmPanel.frx":3A7CE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   9
      Top             =   45
      Width           =   240
   End
   Begin VB.Menu mnuDummy 
      Caption         =   "dummy"
      Visible         =   0   'False
      Begin VB.Menu mnuAfficher 
         Caption         =   "Afficher"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "Quitter"
      End
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MagentaClair = 13 ' &HD%
Const RougeClair = 12 ' &HC%
Const JauneClair = 14 ' &HE%
Const VertClair = 10 ' &HA%
Const CyanClair = 11 ' &HB%
Const BleuClair = 9 ' &H9%

Sub Check1_Click(Index%)
    If Action Then Exit Sub
    Select Case ChoixDeMire
        Case 1
            gv0010(0, Index) = Check1(Index)
        Case 2
            gv0010(1, Index) = Check1(Index)
        Case 3
            gv0010(2, Index) = Check1(Index)
        Case 4
        Case 5
        Case 6
            gv0010(3, Index) = Check1(Index)
        Case 7
            frmPanel.MousePointer = 11
            frmCathodic.MousePointer = 11
            Action = True
            gv0010(3, 2) = Check1(Index)
            frmCathodic.DrawMode = 7
            frmCathodic.Line (0, 0)-(frmCathodic.ScaleWidth - 1, frmCathodic.ScaleHeight - 1), , BF
            frmCathodic.DrawMode = 13
            Action = False
            frmCathodic.MousePointer = 0
            frmPanel.MousePointer = 0
            Exit Sub
        Case 8
            If Index = 1 Then
                If Check1(1) = Checked Then
                    Check1(2).Visible = True
                Else
                    Check1(2).Visible = 0
                End If
            End If
            gv0010(4, Index) = Check1(Index)
    End Select
    frmCathodic.Timer2.Enabled = True
End Sub
Sub Check1_MouseUp(Index%, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Sub chkActiver_Click()
    If Action Then Exit Sub
    Select Case ChoixDeMire
        Case 1
            gv0010(5, 0) = chkActiver
        Case 3
            gv0010(5, 1) = chkActiver
    End Select
    PrépareOptionSelonSélection
    frmCathodic.Timer2.Enabled = True
End Sub
Sub chkActiver_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Sub cmdQuit_Click()
    If gv0046 Then
        Line (cmdQuit.Left + 20, cmdQuit.Top - 15)-Step(18, 10), QBColor(0), BF
    Else
        Line (cmdQuit.Left + 14, cmdQuit.Top - 15)-Step(18, 10), QBColor(0), BF
    End If
    Refresh
    SysTray.InTray = False
    SoundBuffer = LoadResData(2, "JF_Button_SOUND")
    sndPlaySound SoundBuffer(0), SND_SYNC Or SND_NODEFAULT Or SND_MEMORY
    ÉcranPowerOff
End Sub
Sub cmdQuit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Sub cmdHide_Click()
    Me.Hide
End Sub
Sub cmdHide_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Sub cmdAlterner_Click()
    If Action Then Exit Sub
    AlterneCouleur
    BarreDeCouleurs
End Sub
Sub cmdAlterner_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub

Private Sub Command1_Click(Index As Integer)
    ChoixDeMire = Index
    PrépareOptionSelonSélection
    AffichageSelonChoix
End Sub

Sub Form_Load()
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    SysTray.InTray = True
    Action = False
    gv0010(4, 0) = 1
    gv0010(4, 1) = 1
    gv0010(4, 2) = 1
    gv0006 = 1
    DéfinitionPoint = 1
    DéfinitionQuadrilé = 1
    DéfinitionMire = 1
    gv000E = 1
    BarreVerticale(1) = MagentaClair
    BarreVerticale(2) = RougeClair
    BarreVerticale(3) = JauneClair
    BarreVerticale(4) = VertClair
    BarreVerticale(5) = CyanClair
    BarreVerticale(6) = BleuClair
    ChoixDeMire = 9
    gv0040 = True
    PrépareOptionSelonSélection
    frmCathodic.Show
    hwnd1 = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    Me.Show
End Sub
Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Me.Hide
    Me.Visible = False
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    If gv0046 Then
        Line (cmdQuit.Left + 20, cmdQuit.Top - 15)-Step(18, 10), QBColor(0), BF
    Else
        Line (cmdQuit.Left + 14, cmdQuit.Top - 15)-Step(18, 10), QBColor(0), BF
    End If
    Refresh
    SysTray.InTray = False
    ÉcranPowerOff
End Sub
Sub Form_Resize()
    Dim l019E As Integer
    If Me.WindowState = 1 Then
        Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        l019E% = ShowWindow(frmCathodic.hwnd, 0)
    Else
        hwnd1 = FindWindow("Shell_traywnd", "")
        Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        l019E% = ShowWindow(frmCathodic.hwnd, &H4)
    End If
End Sub
Sub Form_Unload(Cancel As Integer)
    SysTray.InTray = False
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End
End Sub
Private Sub mnuAfficher_Click()
    Me.Visible = True
    Me.WindowState = 0
End Sub
Private Sub mnuQuitter_Click()
    Call cmdQuit_Click
End Sub
Sub optAffichage_Click(Index As Integer)
    If Action Then Exit Sub
    Select Case ChoixDeMire
        Case 1
            If gv0006 = Index + 1 Then Exit Sub
            gv0006 = Index + 1
        Case 2
            If DéfinitionPoint = Index + 1 Then Exit Sub
            DéfinitionPoint = Index + 1
        Case 3
            If DéfinitionQuadrilé = Index + 1 Then Exit Sub
            DéfinitionQuadrilé = Index + 1
        Case 4
            If DéfinitionMire = Index + 1 Then Exit Sub
            DéfinitionMire = Index + 1
        Case 5
            If gv000E = Index + 1 Then Exit Sub
            gv000E = Index + 1
    End Select
    AffichageSelonChoix
End Sub
Sub optAffichage_MouseUp(Index%, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Sub Picture4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If Button = 2 Then Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW): Me.Hide
End Sub
Private Sub SysTray_MouseDblClick(Button As Integer, Id As Long)
    Me.Visible = True
    Me.WindowState = 0
End Sub
Private Sub SysTray_MouseUp(Button As Integer, Id As Long)
    If Button = 2 Then
        PopupMenu mnuDummy, , , , mnuAfficher
    End If
End Sub
