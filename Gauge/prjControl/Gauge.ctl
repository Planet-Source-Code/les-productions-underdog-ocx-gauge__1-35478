VERSION 5.00
Begin VB.UserControl Gauge 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LockControls    =   -1  'True
   MaskColor       =   &H00808000&
   MaskPicture     =   "Gauge.ctx":0000
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ToolboxBitmap   =   "Gauge.ctx":090A
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   4365
      Picture         =   "Gauge.ctx":0C1C
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   495
      Width           =   75
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   4425
      Picture         =   "Gauge.ctx":17EE
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   2
      Top             =   1290
      Width           =   75
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   4305
      Picture         =   "Gauge.ctx":2070
      ScaleHeight     =   131
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   1
      Top             =   1290
      Width           =   75
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1875
      TabIndex        =   0
      Top             =   3345
      Width           =   360
   End
End
Attribute VB_Name = "Gauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'================================================================================
'
'Nom Du Projet: OldStyleGauge
'
'Auteur:        Les Productions J.F.
'
'Date:          04-06-2002
'
'Description:   An OCX gauge
'
'================================================================================
Option Explicit

Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Dim pic1hDC As Integer
Dim pic2hDC As Integer
Dim pic3hDC As Integer
Dim pic4hDC As Integer
Dim LaValeur
Dim PicGaucheHDC As Long
Dim PicDroiteHDC As Long
Dim PicCenterHDC As Long
Dim GaugeHDC As Long

Dim c0 As Long
Dim c1 As Long
Dim c2 As Long
Dim c3 As Long
Dim c4 As Long
Dim c5 As Long
Dim c6 As Long
Dim c7 As Long
Dim c8 As Long
Dim c9 As Long
Dim c10 As Long
Dim c11 As Long

Dim xret As Long
Dim t As Long
Const kPaToPsi = 0.1450377377
Const Pi = 3.14159265358979
Const m_def_AfficherValeur = 1
Const m_def_GaugeType = 20
Const m_def_Value = 0

Dim m_AfficherValeur As Boolean
Dim m_GaugeType As ModeleDeGauge
Dim m_Value As Variant


'10 BITMAP DISCARDABLE  0_100Degrée.bmp     0
'20 BITMAP DISCARDABLE  0_100Psi.bmp        1
'30 BITMAP DISCARDABLE  0_200Psi.bmp        2
'40 BITMAP DISCARDABLE  0_500Degrée.bmp     3
'50 BITMAP DISCARDABLE  -50_50Degrée.bmp    4
Public Enum ModeleDeGauge
    Gauge100Psi = 20
    Gauge200Psi = 30
    Gauge100C = 10
    Gauge500C = 40
    Gauge5050C = 50
End Enum

Dim ValeurMin As Integer
Dim ValeurMax As Integer

'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
    UserControl.Cls
End Sub

'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property



Private Sub UserControl_Paint()
    Rotate Value
End Sub

'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 3)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 261)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 365)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_GaugeType = PropBag.ReadProperty("GaugeType", m_def_GaugeType)
    UserControl.Picture = LoadResPicture(m_GaugeType, vbResBitmap)
    m_AfficherValeur = PropBag.ReadProperty("AfficherValeur", m_def_AfficherValeur)
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 3915
    UserControl.Width = 4055
End Sub

Private Sub UserControl_Show()
    UserControl.Picture = LoadResPicture(m_GaugeType, vbResBitmap)
    Rotate Value
End Sub

'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 3)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 261)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 365)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("GaugeType", m_GaugeType, m_def_GaugeType)
    Call PropBag.WriteProperty("AfficherValeur", m_AfficherValeur, m_def_AfficherValeur)
End Sub

Private Sub bmp_rotate(pic1 As PictureBox, pic2 As PictureBox, pic3 As PictureBox, Gauge, ByVal theta!)

    Dim PicCentre_c2x As Integer, PicCentre_c2y As Integer 'aiguille gauche
    Dim Gauge_c2x As Integer, Gauge_c2y As Integer 'aiguille droite
    Dim PicGauche_c2x As Integer, PicGauche_c2y As Integer 'aiguille centre
    Dim PicDroite_c2x As Integer, PicDroite_c2y As Integer 'gauge


    Dim a As Single
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim n As Integer, r As Integer

    PicGauche_c2x = pic1.ScaleWidth / 2 + 4 '- pic1.ScaleWidth 'gauche
    PicGauche_c2y = pic1.ScaleHeight / 2 - 18 ' - pic1.ScaleHeight 'gauche

    PicDroite_c2x = pic2.ScaleWidth / 2 - 5 '- pic2.ScaleWidth 'droite
    PicDroite_c2y = pic2.ScaleHeight / 2 - 18 ' - pic2.ScaleHeight 'droite

    PicCentre_c2x = pic3.ScaleWidth \ 2 'Centre
    PicCentre_c2y = pic3.ScaleHeight \ 2 + 15 'centre

    Gauge_c2x = UserControl.ScaleWidth \ 2 'gauge
    Gauge_c2y = UserControl.ScaleHeight \ 2 'gauge

    If Gauge_c2x < Gauge_c2y Then n = Gauge_c2y Else n = Gauge_c2x

    n = n - 1
    PicGaucheHDC = pic1.hdc 'Gauche
    PicDroiteHDC = pic2.hdc 'Droite
    PicCenterHDC = pic3.hdc 'Centre

    GaugeHDC = UserControl.hdc

    For p2x = 0 To n
        For p2y = 0 To n
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0 = GetPixel(PicCenterHDC, PicCentre_c2x + p1x, PicCentre_c2y + p1y)
            c1 = GetPixel(PicCenterHDC, PicCentre_c2x - p1x, PicCentre_c2y - p1y)
            c2 = GetPixel(PicCenterHDC, PicCentre_c2x + p1y, PicCentre_c2y - p1x)
            c3 = GetPixel(PicCenterHDC, PicCentre_c2x - p1y, PicCentre_c2y + p1x)

            c4 = GetPixel(PicGaucheHDC, PicGauche_c2x + p1x, PicGauche_c2y + p1y)
            c5 = GetPixel(PicGaucheHDC, PicGauche_c2x - p1x, PicGauche_c2y - p1y)
            c6 = GetPixel(PicGaucheHDC, PicGauche_c2x + p1y, PicGauche_c2y - p1x)
            c7 = GetPixel(PicGaucheHDC, PicGauche_c2x - p1y, PicGauche_c2y + p1x)

            c8 = GetPixel(PicDroiteHDC, PicDroite_c2x + p1x, PicDroite_c2y + p1y)
            c9 = GetPixel(PicDroiteHDC, PicDroite_c2x - p1x, PicDroite_c2y - p1y)
            c10 = GetPixel(PicDroiteHDC, PicDroite_c2x + p1y, PicDroite_c2y - p1x)
            c11 = GetPixel(PicDroiteHDC, PicDroite_c2x - p1y, PicDroite_c2y + p1x)


            If c0 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2x, Gauge_c2y + p2y, c0&)
            If c1 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2x, Gauge_c2y - p2y, c1&)
            If c2 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2y, Gauge_c2y - p2x, c2&)
            If c3 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2y, Gauge_c2y + p2x, c3&)

            If c4 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2x, Gauge_c2y + p2y, c4&)
            If c5 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2x, Gauge_c2y - p2y, c5&)
            If c6 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2y, Gauge_c2y - p2x, c6&)
            If c7 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2y, Gauge_c2y + p2x, c7&)

            If c8 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2x, Gauge_c2y + p2y, c8&)
            If c9 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2x, Gauge_c2y - p2y, c9&)
            If c10 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x + p2y, Gauge_c2y - p2x, c10&)
            If c11 <> -1 Then xret = SetPixel(GaugeHDC, Gauge_c2x - p2y, Gauge_c2y + p2x, c11&)
    If m_AfficherValeur = True Then
        Label1.Visible = True
        If GaugeType = Gauge5050C Then
            Label1.Caption = Format$(m_Value, "###0.0")
        Else
            Label1.Caption = Format$(m_Value, "###0")
        End If
    Else
        Label1.Visible = False
    End If


        Next
        t = DoEvents()
    Next
End Sub

Private Sub Rotate(AngleRot As Single)

    Dim Angle

    Select Case GaugeType
        Case 20
            Angle = AngleRot + (-49.5) '0-100
        Case 30
            Angle = (AngleRot / 2) + (-49.5) '0-200
        Case 10
            Angle = AngleRot + (-49.5) '0-100
        Case 40
            Angle = (AngleRot / 5) + (-49.5) '0-500
        Case 50
            Angle = (AngleRot + 50) + (-49.5) '-50_+50
    End Select

    Angle = Pi * CDbl(-Angle) / 66.5 'transforme en Theta

    Call bmp_rotate(Picture1, Picture2, Picture3, UserControl.Picture, Angle)

End Sub
'ATTENTION! NE PAS SUPPRIMER OU MODIFIER LES LIGNES DE COMMENTAIRES QUI SUIVENT!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    If New_Value = Value Then Exit Property

    Select Case GaugeType
        Case 20
            New_Value = New_Value * kPaToPsi
            ValeurMin = 0
            ValeurMax = 100
        Case 30
            New_Value = New_Value * kPaToPsi
            ValeurMin = 0
            ValeurMax = 200
        Case 10
            ValeurMin = 0
            ValeurMax = 100
        Case 40
            ValeurMin = 0
            ValeurMax = 500
        Case 50
            ValeurMin = -50
            ValeurMax = 50
    End Select
    If New_Value < ValeurMin Then New_Value = ValeurMin
    If New_Value > ValeurMax Then New_Value = ValeurMax

    m_Value = New_Value
    PropertyChanged "Value"
    If m_AfficherValeur = True Then
        Label1.Visible = True
        If GaugeType = Gauge5050C Then
            Label1.Caption = Format$(m_Value, "###0.0")
        Else
            Label1.Caption = Format$(m_Value, "###0")
        End If
    Else
        Label1.Visible = False
    End If

    UserControl.Cls

    Rotate New_Value

End Property

'Initialiser les propriétés pour le UserControl
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_AfficherValeur = m_def_AfficherValeur
    m_GaugeType = m_def_GaugeType
End Sub

Public Property Get GaugeType() As ModeleDeGauge
    GaugeType = m_GaugeType
End Property

Public Property Let GaugeType(ByVal New_GaugeType As ModeleDeGauge)
    m_GaugeType = New_GaugeType
    Select Case New_GaugeType
        Case 10
            UserControl.Picture = LoadResPicture(10, vbResBitmap)
        Case 20
            UserControl.Picture = LoadResPicture(20, vbResBitmap)
        Case 30
            UserControl.Picture = LoadResPicture(30, vbResBitmap)
        Case 40
            UserControl.Picture = LoadResPicture(40, vbResBitmap)
        Case 50
            UserControl.Picture = LoadResPicture(50, vbResBitmap)
    End Select
    PropertyChanged "GaugeType"
End Property

Public Property Get AfficherValeur() As Boolean
    AfficherValeur = m_AfficherValeur
End Property

Public Property Let AfficherValeur(ByVal New_AfficherValeur As Boolean)
    m_AfficherValeur = New_AfficherValeur
    If m_AfficherValeur = True Then
        Label1.Visible = True
        If GaugeType = Gauge5050C Then
            Label1.Caption = Format$(m_Value, "###0.0")
        Else
            Label1.Caption = Format$(m_Value, "###0")
        End If
    Else
        Label1.Visible = False
    End If
    PropertyChanged "AfficherValeur"
End Property

