VERSION 5.00
Object = "*\ANeedle.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3450
      Top             =   4050
   End
   Begin OldStyleGauge.Gauge Gauge1 
      Height          =   3915
      Left            =   -15
      TabIndex        =   0
      Top             =   -15
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   6906
      ScaleMode       =   0
      ScaleWidth      =   270
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub HScroll1_Change()
Gauge1.Value = HScroll1.Value
End Sub

Private Sub Timer1_Timer()
Gauge1.Value = 450 + Int((50 * Rnd) + 1)
End Sub
