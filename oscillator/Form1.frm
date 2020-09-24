VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Syfon FX Oscillator Created by viper (Drummer) viperc4335@aol.com"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll5 
      Height          =   255
      Left            =   2160
      Max             =   20
      TabIndex        =   11
      Top             =   3600
      Value           =   1
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll9 
      Height          =   255
      Left            =   2160
      Max             =   20
      TabIndex        =   31
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CheckBox Check11 
      Caption         =   "Amplitude Modulation"
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Oscillator Modulation"
      Height          =   255
      Left            =   4560
      TabIndex        =   29
      Top             =   3120
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      Left            =   4560
      Max             =   20
      TabIndex        =   28
      Top             =   3360
      Value           =   1
      Width           =   1815
   End
   Begin VB.HScrollBar HScroll8 
      Height          =   135
      Left            =   120
      Max             =   20
      Min             =   -20
      TabIndex        =   26
      Top             =   2880
      Width           =   2415
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   5
      Height          =   135
      Left            =   3840
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   25
      Top             =   2880
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   1215
      Left            =   3840
      ScaleHeight     =   250
      ScaleMode       =   0  'User
      ScaleWidth      =   360
      TabIndex        =   24
      Top             =   1680
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   255
      Left            =   1440
      Max             =   1000
      TabIndex        =   21
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CheckBox Check10 
      Caption         =   "Highpass"
      Height          =   255
      Left            =   2760
      TabIndex        =   23
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Lowpass"
      Height          =   255
      Left            =   1440
      TabIndex        =   22
      Top             =   2280
      Width           =   975
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      Left            =   1440
      Max             =   100
      TabIndex        =   18
      Top             =   1920
      Value           =   50
      Width           =   1575
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Square"
      Height          =   855
      Left            =   720
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1905
      Width           =   615
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Sine"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1905
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      Left            =   1440
      Max             =   1000
      TabIndex        =   6
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Highpass"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Lowpass"
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   720
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   255
      Left            =   120
      Max             =   100
      Min             =   1
      TabIndex        =   12
      Top             =   4080
      Value           =   1
      Width           =   6255
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Frequency Modulation"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox o1VU 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   5
      Height          =   135
      Left            =   3840
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1440
      Max             =   100
      TabIndex        =   4
      Top             =   360
      Value           =   50
      Width           =   1575
   End
   Begin VB.PictureBox WPB1 
      BackColor       =   &H00008000&
      Height          =   1335
      Left            =   3840
      ScaleHeight     =   250
      ScaleMode       =   0  'User
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Square"
      Height          =   855
      Left            =   720
      Picture         =   "Form1.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sine"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   30
      Top             =   4995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Tune"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Cutoff"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   20
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Volume"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   19
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Oscillator 2"
      Height          =   255
      Index           =   1
      Left            =   105
      TabIndex        =   15
      Top             =   1680
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Cutoff"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   120
      Top             =   3240
      Width           =   4335
   End
   Begin VB.Label Label4 
      Caption         =   "Speed"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Oscillator 1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Volume"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By: David Katrowski.

Private Sub Form_Load()
If Not Init_DX7(Me.Hwnd) Then End
DSB(0).Play DSBPLAY_LOOPING
DSB(1).Play DSBPLAY_LOOPING
Osc1Amp = HScroll1.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
Term_DX7
End Sub



Private Sub Timer1_Timer(): On Error Resume Next
Osc1Amp = HScroll1.Value
Osc2Amp = HScroll3.Value
Osc1FCutoff = HScroll2.Value
Osc2FCutoff = HScroll7.Value

AM_Speed = HScroll4.Value
AM2_Speed = HScroll9.Value
FM_Speed = HScroll5.Value

Temp1 = 0: Temp2 = 0

f = HScroll6.Value
O2F = f + HScroll8.Value

WPB1.Cls: o1VU.Cls: Picture1.Cls: Picture2.Cls
For i = 0 To 360
f2 = Timer
n = i * 0.01745329251994
'If -(FM_Speed And 1) Then FM_Speed = FM_Speed + 1
If Check4.Value = 1 Then fm = Cos(f2 * (FM_Speed / 2)) * f Else fm = f
If fm < 0 Then fm = fm + 4

'---<
If Check1.Value = 1 Then Osc1Samp = Sin(n * fm) * Osc1Amp
If Check2.Value = 1 Then Osc1Samp = Cos(pi * Int(fm * n)) * Osc1Amp

If Check7.Value = 1 Then Osc2Samp = Sin(n * O2F) * Osc2Amp
If Check8.Value = 1 Then Osc2Samp = Cos(pi * Int(O2F * n)) * Osc2Amp

If Check11.Value = 1 Then
Osc1Samp = Cos(f2 * AM2_Speed) * Osc1Samp
Osc2Samp = Cos(f2 * AM2_Speed) * Osc2Samp
End If

If Check3.Value = 1 Then
Osc1Samp = Cos(f2 * AM_Speed) * Osc1Samp 'Cos(2 * pi * f * Timer) 'f2)
Osc2Samp = Sin(f2 * AM_Speed) * Osc2Samp
End If

If Check5.Value = 1 Then If Osc1FCutoff > 0 Then Osc1Samp = Osc1Samp / (10 * Log(1 + (fm / (0.1 * Osc1FCutoff))))
If Check6.Value = 1 Then If Osc1FCutoff > 0 Then Osc1Samp = Osc1Samp / (10 * Exp(1 + (Abs(fm) / (0.1 * (Osc1FCutoff))) * (2 ^ 2)))

If Check9.Value = 1 Then If Osc2FCutoff > 0 Then Osc2Samp = Osc2Samp / (10 * Log(1 + (fm / (0.1 * Osc2FCutoff))))
If Check10.Value = 1 Then If Osc2FCutoff > 0 Then Osc2Samp = Osc2Samp / (10 * Exp(1 + (Abs(fm) / (0.1 * (Osc2FCutoff))) * (2 ^ 2)))
'---<
O1SBuffer(i) = Osc1Samp + 128
O2SBuffer(i) = Osc2Samp + 128
If Osc1Samp > Temp1 Then Temp1 = Osc1Samp
If Osc2Samp > Temp2 Then Temp2 = Osc2Samp
DrawPOINT i, Osc1Samp, WPB1
DrawPOINT i, Osc2Samp, Picture1
Next
DSBWRITE 0, O1SBuffer()
DrawVU Temp1, o1VU
DSBWRITE 1, O2SBuffer()
DrawVU Temp2, Picture2
End Sub
