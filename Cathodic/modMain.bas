Attribute VB_Name = "modMain"
Option Explicit
Public hwnd1 As Long
Public Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As _
        Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags _
        As Long) As Long
Public Declare Function FindWindow Lib "user32" _
        Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName _
        As String) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal nWidth As Long, _
        ByVal nHeight As Long, _
        ByVal hSrcDC As Long, _
        ByVal xSrc As Long, _
        ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal nCmdShow As Long) As Long
        
Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

Global gv0006 As Integer
Global DéfinitionPoint As Integer
Global DéfinitionQuadrilé As Integer
Global DéfinitionMire As Integer
Global gv000E As Integer
Global gv0010(5, 2) As Integer
Global BarreVerticale(1 To 6) As Integer
Global gv0040 As Integer
Global Action As Boolean
Global ChoixDeMire As Integer
Global gv0046 As Integer
Global gv0048 As Integer
Global gv004A As Integer
Global Const Titre = "Cathodic"
Global Const gc005C = 12 ' &HC%
Global Const gc0060 = 10 ' &HA%
Global Const gc0064 = 9 ' &H9%

Global Const SND_SYNC = &H1      ' Jouer de façon synchrone, et ASYNC de façon asyncrone
Global Const SND_NODEFAULT = &H2 ' Ne pas utiliser le son par défaut.
Global Const SND_MEMORY = &H4    ' lpszSoundName pointe vers un fichier en mémoire.
Global SoundBuffer() As Byte

Sub BarreDeCouleurs()
    Dim l004E As Variant
    Dim l0052 As Variant
    Dim l0056 As Variant
    Dim l005A As Variant
    Dim l005E As Variant
    Dim l0062 As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l004E = DoEvents()
    If frmPanel.Check1(0).Value = Checked Then
        l0052 = frmCathodic.ScaleHeight
        If frmPanel.Check1(1).Value = Checked Then
            l0056 = frmCathodic.ScaleWidth / 18
            If frmPanel.Check1(2).Value = Checked Then l005A = 0 Else l005A = 15
            frmCathodic.Line (0, 0)-Step(l0056, frmCathodic.ScaleHeight), QBColor(BarreVerticale(1)), BF
            frmCathodic.Line (1 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (2 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(2)), BF
            frmCathodic.Line (4 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (5 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(3)), BF
            frmCathodic.Line (7 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (8 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(4)), BF
            frmCathodic.Line (10 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (11 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(5)), BF
            frmCathodic.Line (13 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (14 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(6)), BF
            frmCathodic.Line (16 * l0056, 0)-Step(2 * l0056, l0052), QBColor(l005A), BF
            frmCathodic.Line (17 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(1)), BF
        Else
            l0056 = frmCathodic.ScaleWidth / 12
            frmCathodic.Line (0, 0)-Step(l0056, frmCathodic.ScaleHeight), QBColor(BarreVerticale(1)), BF
            frmCathodic.Line (1 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(2)), BF
            frmCathodic.Line (3 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(3)), BF
            frmCathodic.Line (5 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(4)), BF
            frmCathodic.Line (7 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(5)), BF
            frmCathodic.Line (9 * l0056, 0)-Step(2 * l0056, l0052), QBColor(BarreVerticale(6)), BF
            frmCathodic.Line (11 * l0056, 0)-Step(l0056, l0052), QBColor(BarreVerticale(1)), BF
        End If
    Else
        l005E = frmCathodic.ScaleWidth
        If frmPanel.Check1(1).Value = Checked Then
            l0062 = frmCathodic.ScaleHeight / 18
            If frmPanel.Check1(2).Value = Checked Then l005A = 0 Else l005A = 15
            frmCathodic.Line (0, 0)-Step(frmCathodic.ScaleWidth, l0062), QBColor(BarreVerticale(1)), BF
            frmCathodic.Line (0, 1 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 2 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(2)), BF
            frmCathodic.Line (0, 4 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 5 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(3)), BF
            frmCathodic.Line (0, 7 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 8 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(4)), BF
            frmCathodic.Line (0, 10 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 11 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(5)), BF
            frmCathodic.Line (0, 13 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 14 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(6)), BF
            frmCathodic.Line (0, 16 * l0062)-Step(l005E, 2 * l0062), QBColor(l005A), BF
            frmCathodic.Line (0, 17 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(1)), BF
        Else
            l0062 = frmCathodic.ScaleHeight / 12
            frmCathodic.Line (0, 0)-Step(frmCathodic.ScaleWidth, l0062), QBColor(BarreVerticale(1)), BF
            frmCathodic.Line (0, 1 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(2)), BF
            frmCathodic.Line (0, 3 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(3)), BF
            frmCathodic.Line (0, 5 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(4)), BF
            frmCathodic.Line (0, 7 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(5)), BF
            frmCathodic.Line (0, 9 * l0062)-Step(l005E, 2 * l0062), QBColor(BarreVerticale(6)), BF
            frmCathodic.Line (0, 11 * l0062)-Step(l005E, l0062), QBColor(BarreVerticale(1)), BF
        End If
    End If
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub MiresRondes()
    Dim l0066 As Variant
    Dim l006A As Variant
    Dim l006E As Variant
    Dim l0072 As Variant
    Dim l0076 As Variant
    Dim l007A As Variant
    Dim l007E As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l0066 = DoEvents()
    l006A = frmCathodic.ScaleWidth
    l006E = frmCathodic.ScaleHeight
    'frmCathodic.Line (0, 0)-(l006A - 1, l006E - 1), , B
    If DéfinitionMire < 3 Then
        l0072 = l006A / 16
        l0076 = l006E / 12
        For l007A = l0072 To l006A Step l0072
            frmCathodic.Line (l007A, 1)-(l007A, l006E)
        Next
        For l007A = l0076 To l006E Step l0076
            frmCathodic.Line (1, l007A)-(l006A, l007A)
        Next
    End If
    l0072 = frmCathodic.ScaleWidth / 16
    l0076 = frmCathodic.ScaleHeight / 12
    frmCathodic.FillStyle = 0
    frmCathodic.FillColor = QBColor(0)
    frmCathodic.Circle (l0072 * 2, l0076 * 2), l0072 * 1.4, QBColor(15)
    frmCathodic.PSet (l0072 * 2, l0076 * 2), QBColor(15)
    frmCathodic.Circle (l0072 * 14, l0076 * 2), l0072 * 1.4, QBColor(15)
    frmCathodic.PSet (l0072 * 14, l0076 * 2), QBColor(15)
    frmCathodic.Circle (l0072 * 8, l0076 * 6), l0072 * 4.8, QBColor(15)
    frmCathodic.PSet (l0072 * 8, l0076 * 6), QBColor(15)
    frmCathodic.Circle (l0072 * 2, l0076 * 10), l0072 * 1.4, QBColor(15)
    frmCathodic.PSet (l0072 * 2, l0076 * 10), QBColor(15)
    frmCathodic.Circle (l0072 * 14, l0076 * 10), l0072 * 1.4, QBColor(15)
    frmCathodic.PSet (l0072 * 14, l0076 * 10), QBColor(15)
    frmCathodic.FillStyle = 1
    frmCathodic.FillColor = QBColor(15)
    If DéfinitionMire = 1 Then
        For l007E = (l0072 * 1.4) - 2 To 0 Step -2
            frmCathodic.Circle (l0072 * 2, l0076 * 2), l007E, QBColor(15)
            frmCathodic.Circle (l0072 * 14, l0076 * 2), l007E, QBColor(15)
            frmCathodic.Circle (l0072 * 2, l0076 * 10), l007E, QBColor(15)
            frmCathodic.Circle (l0072 * 14, l0076 * 10), l007E, QBColor(15)
        Next l007E
    End If
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub AffichageSelonChoix()
    If frmCathodic.Line1.Visible And ChoixDeMire <> 3 Or frmPanel.chkActiver = 0 Then
        frmCathodic.Line1.Visible = False
        frmCathodic.Line2.Visible = False
        frmCathodic.Line3.Visible = False
        frmCathodic.Line4.Visible = False
        frmCathodic.lblAffichage.Visible = False
    End If
    Select Case ChoixDeMire
        Case 2, 4, 6
            frmCathodic.Cls
        Case 3
            If frmCathodic.Line1.Visible = False Then frmCathodic.Cls
    End Select
    Select Case ChoixDeMire
        Case 1
            MireDeCouleur
        Case 2
            MireDePoints
        Case 3
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
Sub MireDePoints()
    Dim PasseLaMain As Variant
    Dim LargeurÉcran As Variant
    Dim HauteurÉcran As Variant
    Dim DéfinitionPointLargeur As Variant
    Dim DéfinitionPointHauteur As Variant
    Dim x As Variant
    Dim y As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    PasseLaMain = DoEvents()
    LargeurÉcran = frmCathodic.ScaleWidth
    HauteurÉcran = frmCathodic.ScaleHeight
    frmCathodic.DrawWidth = 2
    Select Case DéfinitionPoint
        Case 1
            DéfinitionPointLargeur = LargeurÉcran / 16
            DéfinitionPointHauteur = HauteurÉcran / 12
        Case 2
            DéfinitionPointLargeur = LargeurÉcran / 8
            DéfinitionPointHauteur = HauteurÉcran / 6
        Case 3
            DéfinitionPointLargeur = LargeurÉcran / 2
            DéfinitionPointHauteur = HauteurÉcran / 2
    End Select
    frmCathodic.ForeColor = ChargeCouleur()
    For x = 0 To LargeurÉcran Step DéfinitionPointLargeur
        For y = 0 To HauteurÉcran Step DéfinitionPointHauteur
            frmCathodic.PSet (x, y)
        Next
    Next
    For x = 0 To LargeurÉcran Step DéfinitionPointLargeur
        frmCathodic.PSet (x, frmCathodic.ScaleHeight - 1)
    Next
    For y = 0 To HauteurÉcran Step DéfinitionPointHauteur
        frmCathodic.PSet (frmCathodic.ScaleWidth - 1, y)
    Next
    frmCathodic.PSet (frmCathodic.ScaleWidth - 1, frmCathodic.ScaleHeight - 1)
    frmCathodic.ForeColor = &HFFFFFF
    frmCathodic.DrawWidth = 1
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub ÉcranPowerOff()
    Dim LargeurÉcran As Variant
    Dim HauteurÉcran As Variant
    Dim Étape As Variant
    Dim a As Integer
    If Action Then Exit Sub
    Action = True
    frmPanel.Hide
    If ChoixDeMire = 3 Then
        frmCathodic.Line1.Visible = False
        frmCathodic.Line2.Visible = False
        frmCathodic.Line3.Visible = False
        frmCathodic.Line4.Visible = False
        frmCathodic.lblAffichage.Visible = False
    End If
    frmCathodic.Cls
    frmCathodic.DrawWidth = 2
    LargeurÉcran = frmCathodic.ScaleWidth / 2
    HauteurÉcran = frmCathodic.ScaleHeight / 2
    Étape = LargeurÉcran / 10
    For a = LargeurÉcran To 1 Step -Étape
        frmCathodic.Circle (LargeurÉcran, HauteurÉcran), a, vbWhite
        frmCathodic.Circle (LargeurÉcran, HauteurÉcran), a + Étape, vbBlack
    Next a
    frmCathodic.FillStyle = vbFSSolid
    frmCathodic.Circle (LargeurÉcran, HauteurÉcran), 2, vbWhite
    frmCathodic.FillStyle = vbFSTransparent
    frmCathodic.Circle (LargeurÉcran, HauteurÉcran), a + Étape, vbBlack
    frmCathodic.tmrTermine.Enabled = True
End Sub
Function ChargeCouleur() As Long
    Dim Couleur As Long
    If frmPanel.Check1(0).Value = Checked Then Couleur = vbRed
    If frmPanel.Check1(1).Value = Checked Then Couleur = Couleur + vbGreen
    If frmPanel.Check1(2).Value = Checked Then Couleur = Couleur + vbBlue
    If Couleur = 0 Then Couleur = vbWhite
    ChargeCouleur = Couleur
End Function
Sub MireQuadrillée()
    Dim Couleur As Variant
    Dim l00C0 As Variant
    Dim l00C4 As Variant
    Dim l00C8 As Variant
    Dim l00CC As Variant
    Dim l00D0 As Variant
    Dim l00D4 As Variant
    If frmPanel.chkActiver = 1 Then
        Couleur = ChargeCouleur()
        frmCathodic.Line1.BorderColor = Couleur
        frmCathodic.Line2.BorderColor = Couleur
        frmCathodic.Line3.BorderColor = Couleur
        frmCathodic.Line4.BorderColor = Couleur
        frmCathodic.Line1.Visible = True
        frmCathodic.Line2.Visible = True
        frmCathodic.Line3.Visible = True
        frmCathodic.Line4.Visible = True
        frmCathodic.lblAffichage.Visible = True
        Exit Sub
    End If
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l00C0 = DoEvents()
    Select Case frmPanel.chkActiver
        Case 0
            l00C4 = frmCathodic.ScaleWidth
            l00C8 = frmCathodic.ScaleHeight
            Select Case DéfinitionQuadrilé
                Case 1
                    l00CC = l00C4 / 16
                    l00D0 = l00C8 / 12
                Case 2
                    l00CC = l00C4 / 8
                    l00D0 = l00C8 / 6
                Case 3
                    l00CC = l00C4 / 2
                    l00D0 = l00C8 / 2
            End Select
            frmCathodic.ForeColor = ChargeCouleur()
            'frmCathodic.Line (0, 0)-(l00C4 - 1, l00C8 - 1), , B
            For l00D4 = l00CC To l00C4 Step l00CC
                frmCathodic.Line (l00D4, 1)-(l00D4, l00C8)
            Next
            For l00D4 = l00D0 To l00C8 Step l00D0
                frmCathodic.Line (1, l00D4)-(l00C4, l00D4)
            Next
            frmCathodic.ForeColor = &HFFFFFF
    End Select
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub BarreDeTV()
    Dim l00D8 As Variant
    Dim l00DC As Variant
    Dim l00E0 As Variant
    Dim l00E4 As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l00D8 = DoEvents()
    l00DC = frmCathodic.ScaleWidth / 7
    l00E0 = frmCathodic.ScaleHeight / 40 * 27
    frmCathodic.Line (0 * l00DC, 0)-Step(l00DC, l00E0), RGB(222, 222, 222), BF
    frmCathodic.Line (1 * l00DC, 0)-Step(l00DC, l00E0), RGB(231, 231, 57), BF
    frmCathodic.Line (2 * l00DC, 0)-Step(l00DC, l00E0), RGB(57, 231, 231), BF
    frmCathodic.Line (3 * l00DC, 0)-Step(l00DC, l00E0), RGB(41, 239, 41), BF
    frmCathodic.Line (4 * l00DC, 0)-Step(l00DC, l00E0), RGB(231, 57, 231), BF
    frmCathodic.Line (5 * l00DC, 0)-Step(l00DC, l00E0), RGB(239, 41, 41), BF
    frmCathodic.Line (6 * l00DC, 0)-Step(l00DC, l00E0), RGB(41, 41, 239), BF
    l00E0 = frmCathodic.ScaleHeight / 40 * 27
    l00E4 = frmCathodic.ScaleHeight / 40 * 3
    frmCathodic.Line (0 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(41, 41, 239), BF
    frmCathodic.Line (1 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    frmCathodic.Line (2 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(231, 57, 231), BF
    frmCathodic.Line (3 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    frmCathodic.Line (4 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(57, 231, 231), BF
    frmCathodic.Line (5 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    frmCathodic.Line (6 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(222, 222, 222), BF
    l00DC = ((frmCathodic.ScaleWidth / 7) * 5) / 4
    l00E0 = frmCathodic.ScaleHeight / 40 * 30
    l00E4 = frmCathodic.ScaleHeight / 40 * 10
    frmCathodic.Line (0 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(0, 33, 123), BF
    frmCathodic.Line (1 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(255, 255, 255), BF
    frmCathodic.Line (2 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(90, 0, 123), BF
    frmCathodic.Line (3 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    l00DC = frmCathodic.ScaleWidth / 21
    frmCathodic.Line (15 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(16, 16, 16), BF
    frmCathodic.Line (16 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    frmCathodic.Line (17 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(41, 41, 41), BF
    l00DC = frmCathodic.ScaleWidth / 7
    frmCathodic.Line (6 * l00DC, l00E0)-Step(l00DC, l00E4), RGB(24, 24, 24), BF
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub MireDeCouleur()
    Dim l00E8 As Variant
    Dim l00EC As Variant
    Dim l00F0 As Variant
    Dim l00F4 As Variant
    Dim l00F8 As Integer
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l00E8 = DoEvents()
    Select Case frmPanel.chkActiver
        Case 0
            Select Case gv0006
                Case 1
                    l00EC = gc005C
                Case 2
                    l00EC = gc0060
                Case 3
                    l00EC = gc0064
            End Select
            frmCathodic.Line (0, 0)-(frmCathodic.ScaleWidth - 1, frmCathodic.ScaleHeight - 1), QBColor(l00EC), BF
        Case 1
            l00F0 = frmCathodic.ScaleWidth / 33
            l00F4 = frmCathodic.ScaleHeight
            frmCathodic.Line (0, 0)-Step(l00F0, l00F4), RGB(0, 0, 0), BF
            Select Case ChargeCouleur()
                Case 255
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, 0, 0), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, 0, 0), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, l00F8% * 16 - 1, l00F8% * 16 - 1), BF
                    Next l00F8%
                Case 65280
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(0, l00F8% * 16 - 1, 0), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(0, 255, 0), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, 255, l00F8% * 16 - 1), BF
                    Next l00F8%
                Case 16711680
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(0, 0, l00F8% * 16 - 1), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(0, 0, 255), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, l00F8% * 16 - 1, 255), BF
                    Next l00F8%
                Case 16776960
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(0, l00F8% * 16 - 1, l00F8% * 16 - 1), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(0, 255, 255), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, 255, 255), BF
                    Next l00F8%
                Case 16711935
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, 0, l00F8% * 16 - 1), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, 0, 255), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, l00F8% * 16 - 1, 255), BF
                    Next l00F8%
                Case 65535
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(l00F8% * 16 - 1, l00F8% * 16 - 1, 0), BF
                    Next l00F8%
                    frmCathodic.Line (l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, 255, 0), BF
                    For l00F8% = 1 To 16
                        frmCathodic.Line (l00F8% * l00F0 + (l00F0 * 16), 0)-Step(l00F0, l00F4), RGB(255, 255, l00F8% * 16 - 1), BF
                    Next l00F8%
                Case 16777215
                    l00F0 = frmCathodic.ScaleWidth / 33
                    For l00F8% = 1 To 32
                        frmCathodic.Line (l00F8% * l00F0, 0)-Step(l00F0, l00F4), RGB(l00F8% * 8 - 1, l00F8% * 8 - 1, l00F8% * 8 - 1), BF
                    Next l00F8%
            End Select
    End Select
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub LignesFines()
    Dim l0100 As Variant
    Dim l0104 As Variant
    Dim l0108 As Variant
    Dim l010C As Variant
    Dim l0110 As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l0100 = frmPanel.Check1(0).Value
    l0104 = frmPanel.Check1(1).Value
    If frmPanel.Check1(0).Value = Unchecked Then
        If frmPanel.Check1(1).Value = Checked Then
            For l0108 = 0 To frmCathodic.ScaleHeight - 1 Step 2
                frmCathodic.Line (0, l0108)-(frmCathodic.ScaleWidth - 1, l0108), QBColor(15)
                l010C = DoEvents()
            Next
        Else
            For l0108 = 0 To frmCathodic.ScaleHeight - 1 Step 4
                frmCathodic.Line (0, l0108)-(frmCathodic.ScaleWidth - 1, l0108 + 1), QBColor(15), B
                l010C = DoEvents()
            Next
        End If
    Else
        If frmPanel.Check1(1).Value = Checked Then
            For l0110 = 0 To frmCathodic.ScaleWidth - 1 Step 2
                frmCathodic.Line (l0110, 0)-(l0110, frmCathodic.ScaleHeight - 1), QBColor(15)
                l010C = DoEvents()
            Next
        Else
            For l0110 = 0 To frmCathodic.ScaleWidth - 1 Step 4
                frmCathodic.Line (l0110, 0)-(l0110 + 1, frmCathodic.ScaleHeight - 1), QBColor(15), B
                l010C = DoEvents()
            Next
        End If
    End If
    If l0100 <> frmPanel.Check1(0).Value Or l0104 <> frmPanel.Check1(1).Value Then
        frmCathodic.Cls
        LignesFines
    End If
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub AlterneCouleur()
    Dim l0114 As Variant
    Dim l0118 As Variant
    l0114 = BarreVerticale(6)
    For l0118 = 6 To 2 Step -1
        BarreVerticale(l0118) = BarreVerticale(l0118 - 1)
    Next
    BarreVerticale(1) = l0114
End Sub
Sub PrépareOptionSelonSélection()
    Dim l011C As Integer
    Action = True
    Select Case ChoixDeMire
        Case 1
            frmPanel.Picture4.Visible = True
            frmPanel.chkActiver = gv0010(5, 0)
            Select Case frmPanel.chkActiver
                Case 0
                    frmPanel.cmdAlterner.Visible = False
                    frmPanel.optAffichage(0).Caption = "Rouge"
                    frmPanel.optAffichage(1).Caption = "Vert"
                    frmPanel.optAffichage(2).Caption = "Bleu"
                    frmPanel.optAffichage(gv0006 - 1) = True
                    If gv0046 Then
                        frmPanel.optAffichage(0).Move 85
                        frmPanel.optAffichage(1).Move 85
                        frmPanel.optAffichage(2).Move 85
                    Else
                        frmPanel.optAffichage(0).Move 85
                        frmPanel.optAffichage(1).Move 85
                        frmPanel.optAffichage(2).Move 85
                    End If
                    frmPanel.optAffichage(0).Visible = True
                    frmPanel.optAffichage(1).Visible = True
                    frmPanel.optAffichage(2).Visible = True
                    frmPanel.chkActiver.Caption = "Dégrader"
                    frmPanel.chkActiver.Visible = True
                    frmPanel.Check1(0).Visible = False
                    frmPanel.Check1(1).Visible = False
                    frmPanel.Check1(2).Visible = False
                Case 1
                    frmPanel.cmdAlterner.Visible = False
                    For l011C% = 0 To 2
                        frmPanel.Check1(l011C%).Value = gv0010(0, l011C%)
                    Next l011C%
                    frmPanel.Check1(0).Caption = "Rouge"
                    frmPanel.Check1(1).Caption = "Vert"
                    frmPanel.Check1(2).Caption = "Bleu"
                    frmPanel.Check1(0).Visible = True
                    frmPanel.Check1(1).Visible = True
                    frmPanel.Check1(2).Visible = True
                    frmPanel.chkActiver.Caption = "Dégrader"
                    frmPanel.chkActiver.Visible = True
                    frmPanel.optAffichage(0).Visible = False
                    frmPanel.optAffichage(1).Visible = False
                    frmPanel.optAffichage(2).Visible = False
            End Select
        Case 2
            frmPanel.Picture4.Visible = True
            For l011C% = 0 To 2
                frmPanel.Check1(l011C%).Value = gv0010(1, l011C%)
            Next l011C%
            frmPanel.Check1(0).Caption = "Rouge"
            frmPanel.Check1(1).Caption = "Vert"
            frmPanel.Check1(2).Caption = "Bleu"
            frmPanel.Check1(0).Visible = True
            frmPanel.Check1(1).Visible = True
            frmPanel.Check1(2).Visible = True
            If gv0046 Then
                frmPanel.optAffichage(0).Move 163
                frmPanel.optAffichage(1).Move 163
                frmPanel.optAffichage(2).Move 163
            Else
                frmPanel.optAffichage(0).Move 140
                frmPanel.optAffichage(1).Move 140
                frmPanel.optAffichage(2).Move 140
            End If
            frmPanel.optAffichage(0).Caption = "Haute"
            frmPanel.optAffichage(1).Caption = "Moyenne"
            frmPanel.optAffichage(2).Caption = "Basse"
            frmPanel.optAffichage(DéfinitionPoint - 1) = True
            frmPanel.optAffichage(0).Visible = True
            frmPanel.optAffichage(1).Visible = True
            frmPanel.optAffichage(2).Visible = True
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
        Case 3
            frmPanel.Picture4.Visible = True
            frmPanel.cmdAlterner.Visible = False
            For l011C% = 0 To 2
                frmPanel.Check1(l011C%).Value = gv0010(2, l011C%)
            Next l011C%
            frmPanel.Check1(0).Caption = "Rouge"
            frmPanel.Check1(1).Caption = "Vert"
            frmPanel.Check1(2).Caption = "Bleu"
            frmPanel.Check1(0).Visible = True
            frmPanel.Check1(1).Visible = True
            frmPanel.Check1(2).Visible = True
            If gv0010(5, 1) = 0 Then
                If gv0046 Then
                    frmPanel.optAffichage(0).Move 163
                    frmPanel.optAffichage(1).Move 163
                    frmPanel.optAffichage(2).Move 163
                Else
                    frmPanel.optAffichage(0).Move 140
                    frmPanel.optAffichage(1).Move 140
                    frmPanel.optAffichage(2).Move 140
                End If
                frmPanel.optAffichage(0).Caption = "Haute"
                frmPanel.optAffichage(1).Caption = "Moyenne"
                frmPanel.optAffichage(2).Caption = "Basse"
                frmPanel.optAffichage(DéfinitionQuadrilé - 1) = True
                frmPanel.optAffichage(0).Visible = True
                frmPanel.optAffichage(1).Visible = True
                frmPanel.optAffichage(2).Visible = True
            Else
                frmPanel.optAffichage(0).Visible = False
                frmPanel.optAffichage(1).Visible = False
                frmPanel.optAffichage(2).Visible = False
            End If
            frmPanel.chkActiver = gv0010(5, 1)
            frmPanel.chkActiver.Caption = "Activer"
            frmPanel.chkActiver.Visible = True
        Case 4
            frmPanel.Picture4.Visible = True
            If gv0046 Then
                frmPanel.optAffichage(0).Move 85
                frmPanel.optAffichage(1).Move 85
                frmPanel.optAffichage(2).Move 85
            Else
                frmPanel.optAffichage(0).Move 68
                frmPanel.optAffichage(1).Move 68
                frmPanel.optAffichage(2).Move 68
            End If
            frmPanel.optAffichage(0).Caption = "Haute"
            frmPanel.optAffichage(1).Caption = "Moyenne"
            frmPanel.optAffichage(2).Caption = "Basse"
            frmPanel.optAffichage(DéfinitionMire - 1) = True
            frmPanel.optAffichage(0).Visible = True
            frmPanel.optAffichage(1).Visible = True
            frmPanel.optAffichage(2).Visible = True
            frmPanel.Check1(0).Visible = False
            frmPanel.Check1(1).Visible = False
            frmPanel.Check1(2).Visible = False
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
        Case 5
            frmPanel.Picture4.Visible = True
            If gv0046 Then
                frmPanel.optAffichage(0).Move 85
                frmPanel.optAffichage(1).Move 85
                frmPanel.optAffichage(2).Move 85
            Else
                frmPanel.optAffichage(0).Move 68
                frmPanel.optAffichage(1).Move 68
                frmPanel.optAffichage(2).Move 68
            End If
            frmPanel.optAffichage(0).Caption = "Blanc"
            frmPanel.optAffichage(1).Caption = "2/3"
            frmPanel.optAffichage(2).Caption = "1/3"
            frmPanel.optAffichage(gv000E - 1) = True
            frmPanel.optAffichage(0).Visible = True
            frmPanel.optAffichage(1).Visible = True
            frmPanel.optAffichage(2).Visible = True
            frmPanel.Check1(0).Visible = False
            frmPanel.Check1(1).Visible = False
            frmPanel.Check1(2).Visible = False
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
        Case 6
            frmPanel.Picture4.Visible = True
            frmPanel.Check1(0).Value = gv0010(3, 0)
            frmPanel.Check1(1).Value = gv0010(3, 1)
            frmPanel.Check1(0).Caption = "Horizontal/Vertical"
            frmPanel.Check1(1).Caption = "1 pixel/2 pixels"
            frmPanel.Check1(0).Visible = True
            frmPanel.Check1(1).Visible = True
            frmPanel.optAffichage(0).Visible = False
            frmPanel.optAffichage(1).Visible = False
            frmPanel.optAffichage(2).Visible = False
            frmPanel.Check1(2).Visible = False
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
        Case 7
            frmPanel.Picture4.Visible = True
            frmPanel.Check1(1).Value = gv0010(3, 2)
            frmPanel.Check1(1).Caption = "Inverser"
            frmPanel.Check1(1).Visible = True
            frmPanel.optAffichage(0).Visible = False
            frmPanel.optAffichage(1).Visible = False
            frmPanel.optAffichage(2).Visible = False
            frmPanel.Check1(0).Visible = False
            frmPanel.Check1(2).Visible = False
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
            frmPanel.lblAffichage.Visible = False
        Case 8
            frmPanel.Picture4.Visible = True
            For l011C% = 0 To 2
                frmPanel.Check1(l011C%).Value = gv0010(4, l011C%)
            Next l011C%
            frmPanel.Check1(0).Caption = "Vertical/Horizontal"
            frmPanel.Check1(1).Caption = "Barres de séparations"
            frmPanel.Check1(2).Caption = "Noir/Blanc"
            frmPanel.Check1(0).Visible = True
            frmPanel.Check1(1).Visible = True
            If frmPanel.Check1(1).Value = Checked Then
                frmPanel.Check1(2).Visible = True
            End If
            frmPanel.cmdAlterner.Visible = True
            frmPanel.optAffichage(0).Visible = False
            frmPanel.optAffichage(1).Visible = False
            frmPanel.optAffichage(2).Visible = False
            frmPanel.chkActiver.Visible = False
        Case 9
            frmPanel.Picture4.Visible = True
            frmPanel.Check1(0).Value = gv0010(3, 2)
            frmPanel.Check1(0).Caption = "Inverser"
            frmPanel.Check1(0).Visible = False
            frmPanel.optAffichage(0).Visible = False
            frmPanel.optAffichage(1).Visible = False
            frmPanel.optAffichage(2).Visible = False
            frmPanel.Check1(1).Visible = False
            frmPanel.Check1(2).Visible = False
            frmPanel.chkActiver.Visible = False
            frmPanel.cmdAlterner.Visible = False
            frmPanel.lblAffichage.Visible = True
    End Select
    Action = False
End Sub
Sub AffichePixel()
    Dim l011E As Variant
    Dim l0122 As Variant
    Dim l0126 As Variant
    Dim l012A As Variant
    Dim l012E As Variant
    Dim l0132 As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l011E = frmPanel.Check1(0).Value
    If frmPanel.Check1(0).Value = Unchecked Then
        l0122 = &H330008
    Else
        l0122 = &HCC0020
    End If
    For l0126 = 0 To frmCathodic.ScaleHeight Step 24
        For l012A = 0 To frmCathodic.ScaleWidth Step 86
            l012E = BitBlt(frmCathodic.hDC, l012A, l0126, frmCathodic.picPixel.ScaleWidth, frmCathodic.picPixel.ScaleHeight, frmCathodic.picPixel.hDC, 0, 0, l0122)
            l0132 = DoEvents()
        Next l012A
    Next l0126
    If l011E <> frmPanel.Check1(0).Value Then
        frmCathodic.DrawMode = 7
        frmCathodic.Line (0, 0)-(frmCathodic.ScaleWidth - 1, frmCathodic.ScaleHeight - 1), , BF
        frmCathodic.DrawMode = 13
    End If
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub
Sub TonDeGris()
    Dim l0136 As Variant
    Dim l013A As Variant
    Action = True
    frmPanel.MousePointer = 11
    frmCathodic.MousePointer = 11
    l0136 = DoEvents()
    Select Case gv000E
        Case 1
            l013A = 15
        Case 2
            l013A = 7
        Case 3
            l013A = 8
    End Select
    frmCathodic.Line (0, 0)-(frmCathodic.ScaleWidth - 1, frmCathodic.ScaleHeight - 1), QBColor(l013A), BF
    frmCathodic.MousePointer = 0
    frmPanel.MousePointer = 0
    Action = False
End Sub

Sub BeginPlaySound(ByVal ResourceId As Integer)
    SoundBuffer = LoadResData(ResourceId, "JF_Button_SOUND")
    sndPlaySound SoundBuffer(0), SND_SYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub

