VERSION 5.00
Begin VB.Form frmCathodic 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2430
   ClientLeft      =   1080
   ClientTop       =   1470
   ClientWidth     =   3855
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer tmrTermine 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2115
      Top             =   180
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2700
      Top             =   180
   End
   Begin VB.PictureBox picPixel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   345
      Picture         =   "frmCathodic.frx":0000
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   86
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblAffichage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Affichage des coordonnées"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   225
      Left            =   705
      TabIndex        =   1
      Top             =   2025
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   14
      X2              =   14
      Y1              =   51
      Y2              =   150
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   48
      X2              =   143
      Y1              =   87
      Y2              =   87
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   48
      X2              =   143
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   36
      X2              =   36
      Y1              =   51
      Y2              =   150
   End
End
Attribute VB_Name = "frmCathodic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m001A As Integer
Dim m001C As Integer
Sub Form_Load()
    Width = Screen.Width
    Height = Screen.Height
    Left = 0
    Top = 0
    If ScaleHeight < 480 Then
        Beep
        MsgBox Chr$(10) + "Désolé, " + Titre + " nécéssite au moins " + Chr$(10) + "une résolution VGA(640x480) pour fonctionner." + Chr$(10), 16, App.Title$
        End
    End If
    Line1.Y1 = 0
    Line1.X1 = ScaleWidth \ 2
    Line1.Y2 = ScaleHeight
    Line1.X2 = ScaleWidth \ 2
    Line2.Y1 = ScaleHeight \ 2
    Line2.X1 = 0
    Line2.Y2 = ScaleHeight \ 2
    Line2.X2 = ScaleWidth
    Line3.Y1 = 0
    Line3.X1 = ScaleWidth \ 2
    Line3.Y2 = ScaleHeight
    Line3.X2 = ScaleWidth \ 2
    Line4.Y1 = ScaleHeight \ 2
    Line4.X1 = 0
    Line4.Y2 = ScaleHeight \ 2
    Line4.X2 = ScaleWidth
    Positionnement (ScaleWidth / 2) * 0.95 + (ScaleWidth / 2) - 1, (ScaleHeight / 2) * 0.95 + (ScaleHeight / 2) - 1
End Sub
Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    If gv0048 Then Unload Me
    If Button = 1 Then m001C = True
End Sub
Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim l0088 As Integer
    Dim l008A As Integer
    If gv0048 And m001A > 1 Then Unload Me
    If ChoixDeMire = 3 And frmPanel.chkActiver = 1 And Button = 1 And m001A > 1 Then
        If Shift = 1 Then
            If x + 1 > ScaleWidth \ 2 Then x = Abs(x + 1 - ScaleWidth)
            If y + 1 > ScaleHeight \ 2 Then y = Abs(y + 1 - ScaleHeight)
            l0088% = (Abs(x * 2 - ScaleWidth) / ScaleWidth) * 100
            l008A% = (Abs(y * 2 - ScaleHeight) / ScaleHeight) * 100
            If l0088% > l008A% Then l008A% = l0088%
            gv0048 = True
            Positionnement (ScaleWidth - (ScaleWidth * l008A% \ 100)) \ 2, (ScaleHeight - (ScaleHeight * l008A% \ 100)) \ 2
            gv0048 = False
        Else
            Positionnement x, y
        End If
    End If
    If (ChoixDeMire = 3 And Button = 1) Or gv0048 Then
        If m001A < 2 Then m001A = m001A + 1
    End If
End Sub
Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Action Then Exit Sub
    m001C = False
    If Button = 1 And m001A < 2 Then
        m001A = 0
        If Not frmPanel.Visible Then frmPanel.Show 1
        Exit Sub
    End If
    m001A = 0
    If Button = 2 Then
        Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
        frmPanel.WindowState = 1
        frmPanel.Visible = False
        Me.Hide
        Exit Sub
    End If
    If Left$(lblAffichage, 1) = "M" Then lblAffichage = Right$(lblAffichage, Len(lblAffichage) - InStr(lblAffichage, Chr$(13)) + 1)
End Sub
Sub Form_Paint()
    If Action Or Not Visible Then Exit Sub
    Select Case ChoixDeMire
        Case 1
            MireDeCouleur
        Case 2
            MireDePoints
        Case 3
            If frmPanel.chkActiver = 1 Then Exit Sub
            MireQuadrillée
        Case 4
            MiresRondes
        Case 5
            TonDeGris
        Case 6
            LignesFines
        Case 7
            AffichePixel
        Case 8
            BarreDeCouleurs
        Case 9
            BarreDeTV
    End Select
End Sub
Sub Form_Unload(Cancel As Integer)
    Do While ShowCursor(False) >= gv004A
    Loop
    Do While ShowCursor(True) < gv004A
    Loop
    End
End Sub
Sub lblAffichage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And m001A < 2 Then
        m001A = 0
        If Not frmPanel.Visible Then frmPanel.Show 1
        Exit Sub
    End If
    m001A = 0
    If Button = 2 Then
        frmPanel.WindowState = 1
        frmPanel.Show 1
        Exit Sub
    End If
End Sub
Sub Positionnement(Large As Single, Haut As Single)
    Dim LargeurEcran As Variant
    Dim HauteurEcran As Variant
    Dim Position_X As String
    LargeurEcran = ScaleWidth
    HauteurEcran = ScaleHeight
    Line1.X1 = Large
    Line1.X2 = Large
    Line2.Y1 = Haut
    Line2.Y2 = Haut
    Line3.X1 = LargeurEcran - Large - 1
    Line3.X2 = LargeurEcran - Large - 1
    Line4.Y1 = HauteurEcran - Haut - 1
    Line4.Y2 = HauteurEcran - Haut - 1
    If Not gv0048 And m001A And m001C Then
        Position_X = "M=" & Large + 1 & ":" & Haut + 1 & Chr$(13)
    End If
    Position_X = Position_X & "X="
    If Large < LargeurEcran \ 2 Then
        Position_X = Position_X & Format$(Abs(Large - (LargeurEcran \ 2)) * 2, "##0")
    Else
        Position_X = Position_X & Format$(Abs(Large - (LargeurEcran \ 2)) * 2 + 2, "##0")
    End If
    Position_X = Position_X & ":"
    Position_X = Position_X & Format$(Abs(Large - (LargeurEcran \ 2)) / (LargeurEcran \ 2), "##0%")
    Position_X = Position_X & Chr$(13) & "Y="
    If Haut < HauteurEcran \ 2 Then
        Position_X = Position_X & Format$(Abs(Haut - (HauteurEcran \ 2)) * 2, "##0")
    Else
        Position_X = Position_X & Format$(Abs(Haut - (HauteurEcran \ 2)) * 2 + 2, "##0")
    End If
    Position_X = Position_X & ":"
    Position_X = Position_X & Format$(Abs(Haut - (HauteurEcran \ 2)) / (HauteurEcran \ 2), "##0%")
    lblAffichage.Caption = Position_X
    If Large < LargeurEcran \ 2 Then
        lblAffichage.Left = LargeurEcran - Large - lblAffichage.Width - 3
    Else
        lblAffichage.Left = Large - lblAffichage.Width - 2
    End If
    If Haut < HauteurEcran \ 2 Then
        lblAffichage.Top = HauteurEcran - Haut - lblAffichage.Height - 1
    Else
        lblAffichage.Top = Haut - lblAffichage.Height - 0
    End If
End Sub
Sub tmrTermine_Timer()
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End
End Sub
Sub Timer2_Timer()
    Timer2.Enabled = False
    AffichageSelonChoix
End Sub
