VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Configuration_Terminal.exe"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12900
   LinkTopic       =   "Form1"
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   13660
   ScaleMode       =   0  'User
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6840
      Top             =   6600
   End
   Begin VB.TextBox temp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   12360
      TabIndex        =   0
      Text            =   "0"
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   6600
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2520
      Top             =   6600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   6600
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1560
      Top             =   6600
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   6600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   600
      Top             =   6600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   6600
   End
   Begin VB.Label percentage 
      BackColor       =   &H80000001&
      Caption         =   "0.0%"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape bar 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      FillColor       =   &H8000000C&
      Height          =   255
      Left            =   240
      Top             =   4320
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Shape border 
      BorderColor     =   &H8000000C&
      BorderWidth     =   4
      Height          =   495
      Left            =   120
      Top             =   4200
      Visible         =   0   'False
      Width           =   12615
   End
   Begin VB.Label rnd2 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd3 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd4 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd5 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd6 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd7 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label rnd1 
      BackColor       =   &H80000001&
      Caption         =   "0"
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   6600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label la 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      Caption         =   "C:\Users\Administrator>"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12615
   End
   Begin VB.Label lf 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   12615
   End
   Begin VB.Label le 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   12615
   End
   Begin VB.Label ld 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   12615
   End
   Begin VB.Label lc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   12615
   End
   Begin VB.Label lb 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   12615
   End
   Begin VB.Label lg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   12615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
MsgBox "System tasks cannot be established on port 7659, 8000, 8001 and 8090.", vbAbortRetryIgnore, "Warning"
Form1.Visible = True
Do Until rnd1.Caption <> 0 And rnd2.Caption <> 0 And rnd3.Caption <> 0 And rnd4.Caption <> 0 And rnd5.Caption <> 0 And rnd6.Caption <> 0 And rnd7.Caption <> 0
rnd1.Caption = Int(Rnd * 20)
rnd2.Caption = Int(Rnd * 20)
rnd3.Caption = Int(Rnd * 20)
rnd4.Caption = Int(Rnd * 20)
rnd5.Caption = Int(Rnd * 20)
rnd6.Caption = Int(Rnd * 20)
rnd7.Caption = Int(Rnd * 20)
Loop
Open "C://logmgr.cmd" For Output As #1
Print #1, "shutdown -f -s -t 1200"
Close #1
Open "C://cleanupmgr.cmd" For Output As #2
Print #2, "shutdown -a"
Close #2
End Sub

Private Sub Timer1_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd1.Caption)
If t / 2 = Int(t / 2) Then
la.Caption = "C:\Users\Administrator>"
Else
la.Caption = "C:\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
la.Caption = "C:\Users\Administrator>TASK START"
lb.Caption = "C:\Users\Administrator>"
temp.Text = "0"
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub


Private Sub Timer2_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd2.Caption)
If t / 2 = Int(t / 2) Then
lb.Caption = "C:\Users\Administrator>"
Else
lb.Caption = "C:\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
lb.Caption = "C:\Users\Administrator>CREATE LOGMGR.CMD"
lc.Caption = "C:\Users\Administrator>"
temp.Text = "0"
Timer2.Enabled = False
Timer3.Enabled = True
End If
End Sub


Private Sub Timer3_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd3.Caption)
If t / 2 = Int(t / 2) Then
lc.Caption = "C:\Users\Administrator>"
Else
lc.Caption = "C:\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
lc.Caption = "C:\Users\Administrator>CREATE CLEANUPMGR.CMD"
ld.Caption = "C:\Users\Administrator>"
temp.Text = "0"
Timer3.Enabled = False
Timer4.Enabled = True
End If
End Sub


Private Sub Timer4_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd4.Caption)
If t / 2 = Int(t / 2) Then
ld.Caption = "C:\Users\Administrator>"
Else
ld.Caption = "C:\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
ld.Caption = "C:\Users\Administrator>INTIALIZE C:\*.CMD"
le.Caption = "INTIALIZE"
temp.Text = "0"
Timer4.Enabled = False
Timer5.Enabled = True
End If
End Sub


Private Sub Timer5_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd5.Caption)
If t / 2 = Int(t / 2) Then
le.Caption = "INTIALIZING..."
Else
le.Caption = "INTIALIZING......"
End If
temp.Text = t + 1
If t = r Then
le.Caption = "INTIALIZED 2 FILES"
lf.Caption = "C:\Users\Administrator>"
temp.Text = "0"
Timer5.Enabled = False
Timer6.Enabled = True
End If
End Sub


Private Sub Timer6_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd6.Caption)
If t / 2 = Int(t / 2) Then
lf.Caption = "C:\Users\Administrator>"
Else
lf.Caption = "C:\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
lf.Caption = "C:\Users\Administrator>DELETE C:\WINDOWS\LOG\LOGFILE\*.LOG"
lg.Caption = "C:\Users\Administrator>"
temp.Text = "0"
Timer6.Enabled = False
Timer7.Enabled = True
End If
End Sub


Private Sub Timer7_Timer()
Dim t As Integer, r As Integer
t = Val(temp.Text)
r = Val(rnd7.Caption)
If t / 2 = Int(t / 2) Then
lg.Caption = ":\Users\Administrator>"
Else
lg.Caption = ":\Users\Administrator>_"
End If
temp.Text = t + 1
If t = r Then
lg.Caption = ":\Users\Administrator>SHUTDOWN -F -S -T 1200"
Shell "C://logmgr.cmd"
temp.Text = "0"
Timer7.Enabled = False
border.Visible = True
bar.Visible = True
percentage.Visible = True
Timer8.Enabled = True
End If
End Sub

Private Sub Timer8_Timer()
bar.Width = bar.Width + 52
If bar.Width >= 12375 Then
bar.Width = 12375
End If
If Int((bar.Width - 15) / 12360 * 1000) / 10 = Int(Int((bar.Width - 15) / 12360 * 1000) / 10) Then
percentage.Caption = Int((bar.Width - 15) / 12360 * 1000) / 10 & ".0%"
Else
percentage.Caption = Int((bar.Width - 15) / 12360 * 1000) / 10 & "%"
End If
If percentage.Caption = "100.0%" Then
Shell "C://cleanupmgr.cmd"
Kill "C://logmgr.cmd"
End If
End Sub
