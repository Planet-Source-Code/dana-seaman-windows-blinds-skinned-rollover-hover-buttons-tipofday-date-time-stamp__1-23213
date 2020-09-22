VERSION 5.00
Begin VB.Form frmDate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   8880
   ControlBox      =   0   'False
   Icon            =   "DateTimeStamp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "DateTimeStamp.frx":000C
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   5460
      Top             =   660
   End
   Begin VB.VScrollBar VScroll 
      Height          =   285
      Index           =   2
      Left            =   4050
      Max             =   59
      TabIndex        =   24
      Top             =   4440
      Width           =   180
   End
   Begin VB.VScrollBar VScroll 
      Height          =   285
      Index           =   1
      Left            =   2700
      Max             =   59
      TabIndex        =   23
      Top             =   4440
      Width           =   180
   End
   Begin VB.VScrollBar VScroll 
      Height          =   285
      Index           =   0
      Left            =   1380
      Max             =   23
      TabIndex        =   22
      Top             =   4440
      Width           =   180
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   4440
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   4440
      Width           =   450
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0"
      Top             =   4440
      Width           =   450
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00CEEFF7&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   5280
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   17
      Top             =   1560
      Width           =   3000
      Begin VB.CheckBox chkWhich 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   32
         Top             =   1920
         Width           =   195
      End
      Begin VB.CheckBox chkWhich 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   31
         Top             =   720
         Width           =   195
      End
      Begin VB.CheckBox chkWhich 
         Caption         =   "Check1"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   30
         Top             =   1320
         Width           =   195
      End
      Begin VB.Label lblWhich 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Accessed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   35
         Top             =   1920
         Width           =   840
      End
      Begin VB.Label lblWhich 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Created"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   34
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label lblWhich 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Modified"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   33
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   2
         Left            =   120
         Picture         =   "DateTimeStamp.frx":823F
         Top             =   1890
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   120
         Picture         =   "DateTimeStamp.frx":87C9
         Top             =   1305
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   120
         Picture         =   "DateTimeStamp.frx":8D53
         Top             =   690
         Width           =   240
      End
      Begin VB.Label lblChange 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Change which dates?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   480
         TabIndex        =   18
         Tag             =   "1500"
         Top             =   120
         Width           =   1920
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   440
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   8
      Top             =   1560
      Width           =   4260
      Begin VB.PictureBox picMonth 
         BackColor       =   &H00CEEFF7&
         ClipControls    =   0   'False
         ForeColor       =   &H80000013&
         Height          =   2295
         Left            =   0
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   277
         TabIndex        =   16
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   7
         Left            =   3600
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   14
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   11
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   10
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblDay 
         Alignment       =   2  'Center
         BackColor       =   &H0098CCD0&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Text            =   "Cbo1(1)"
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Text            =   "Cbo1(0)"
      Top             =   360
      Width           =   1815
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   0
      Left            =   4560
      TabIndex        =   26
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "Cancel"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   0
      TransparentColor=   16711935
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   1
      Left            =   6000
      TabIndex        =   27
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "Apply"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   0
      TransparentColor=   16711935
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   2
      Left            =   7440
      TabIndex        =   28
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "OK"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   0
      TransparentColor=   16711935
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   3
      Left            =   480
      TabIndex        =   29
      Top             =   960
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "<<"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   0
      TransparentColor=   16711935
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   4
      Left            =   3540
      TabIndex        =   25
      Top             =   960
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   ">>"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   0
      TransparentColor=   16711935
      TransparentColor=   16711935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      BorderWidth     =   4
      Height          =   6195
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   8820
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   6300
      Picture         =   "DateTimeStamp.frx":92DD
      Top             =   420
      Width           =   945
   End
   Begin VB.Label lblShortDate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label lblSeconds 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3525
      TabIndex        =   5
      Tag             =   "1305"
      Top             =   4800
      Width           =   765
   End
   Begin VB.Label lblMinutes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2160
      TabIndex        =   4
      Tag             =   "1304"
      Top             =   4800
      Width           =   705
   End
   Begin VB.Label lblHours 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   945
      TabIndex        =   3
      Tag             =   "1303"
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MoYr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2340
      TabIndex        =   2
      Top             =   960
      Width           =   465
   End
   Begin VB.Label lblLongDate 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2400
      TabIndex        =   0
      Top             =   5280
      Width           =   255
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Grid dimensions for days
Private Const GRID_ROWS = 6
Private Const GRID_COLS = 7
'Private variables
Private m_CurrDate As Date  ', m_bAcceptChange As Boolean
Private m_nGridWidth As Integer, m_nGridHeight As Integer
'Private DropHandler As New CFullDrop
'Private WithEvents m_t1 As CTimer
Private MyTime As Date

Private Sub Button_Click(Index As Integer)
   Dim L4 As Long
   On Error GoTo ProcedureError
   WhichDates = 0
   PlaySound 61
   For L4 = 0 To 2
      If chkWhich(L4).Value = 1 Then
         WhichDates = WhichDates + 2 ^ L4
      End If
   Next

   Select Case Index
      Case 0 'Cancel
         Unload Me
      Case 1, 2 'Apply 'OK
         If WhichDates > 0 Then
            Form1.UpdateDates
         End If
         If Index = 2 Then
            Unload Me
         End If
      Case 3 'Prev
         SetNewDate DateAdd("m", -1, m_CurrDate)
      Case 4 'Next
         SetNewDate DateAdd("m", 1, m_CurrDate)
   End Select

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".Button_Click") = vbRetry Then Resume Next

End Sub

Private Sub Button_MouseEnter(Index As Integer)
   PlaySound 60
End Sub

Private Sub Combo1_Click(Index As Integer)
   On Error GoTo ProcedureError
If Index = 0 Then
   If Combo1(0).ListIndex + 1 <> Month(m_CurrDate) Then
      SetNewDate DateSerial(Combo1(1).ListIndex + 1900, Combo1(0).ListIndex + 1, Day(m_CurrDate))
   End If
Else
   If Combo1(1).ListIndex + 1900 <> Year(m_CurrDate) Then
      SetNewDate DateSerial(Combo1(1).ListIndex + 1900, Combo1(0).ListIndex + 1, DatePart("d", m_CurrDate))
   End If
End If

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".combo1(0)_Click") = vbRetry Then Resume Next
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error Resume Next
   Select Case KeyCode
      Case 37, 39 'Left, Right
         Combo1(Abs(Index - 1)).SetFocus
   End Select
End Sub

Private Sub Form_Load()
   Dim L4 As Long
   On Error GoTo ProcedureError
   'Tile.Picture = frmMain.Image1(0).Picture
   'Tile.TileArea Me.hdc, 0, 0, Me.Width, Me.Height

   MakeFormRounded Me, 6   'Round the corners
   ResourceSkinNum = 52    'Base index for resource
   SkinButtons Me

   OutlineControl Picture1, Me
   OutlineControl Picture2, Me
   
   'Set Combo height to show all 12 month names
  ' DropHandler.hWnd = Combo1(0).hWnd
   
  ' SetControlCaptionStrings Me

   m_CurrDate = Date

  ' For L4 = 0 To 2
  '    Image2(L4).Picture = frmMain.m_cIL16.ItemPicture("dt")
  ' Next

   For L4 = 1900 To 2100 'put whatever years you want here,
     Combo1(1).AddItem Str$(L4) 'but don't forget to also change the code in setdate
   Next

   For L4 = 1 To 12 'month names
      Combo1(0).AddItem MonthName(L4)
      If L4 < 8 Then 'day names
         lblDay(L4) = Format$(DateSerial(1900, 4, L4), "ddd")
      End If
   Next

   Combo1(1).ListIndex = DatePart("yyyy", m_CurrDate) - 1900
   Combo1(0).ListIndex = DatePart("m", m_CurrDate) - 1

    'Calculate calendar grid measurements
    m_nGridWidth = ((picMonth.ScaleWidth) \ GRID_COLS)
    m_nGridHeight = ((picMonth.ScaleHeight) \ GRID_ROWS)
   ' m_bAcceptChange = False

   Text1(0).Text = Hour(Time)
   VScroll(0).Value = VScroll(0).Max - Val(Text1(0).Text)
   Text1(1).Text = Minute(Time)
   VScroll(1).Value = VScroll(1).Max - Val(Text1(1).Text)
   Text1(2).Text = Second(Time)
   VScroll(2).Value = VScroll(2).Max - Val(Text1(2).Text)

   UpdateCaptions
   
   ' Set m_t1 = New CTimer
   ' m_t1.Interval = 200
   Timer1.Interval = 200

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".Form_Load") = vbRetry Then Resume Next

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub

Private Sub SetNewDate(NewDate As Date)
   On Error GoTo ProcedureError
    If Month(m_CurrDate) = Month(NewDate) And Year(m_CurrDate) = Year(NewDate) Then
        'DrawSelectionBox False
        m_CurrDate = NewDate
        'DrawSelectionBox True
        picMonth_Paint
    Else
        m_CurrDate = NewDate

        picMonth_Paint
    End If

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".SetNewDate") = vbRetry Then Resume Next

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'm_t1.Interval = 0
   'Set m_t1 = Nothing
   Timer1.Interval = 0
End Sub

Private Sub m_t1_ThatTime()
   On Error GoTo ProcedureError
   Static Once As Boolean
   Dim i As Integer
   If Not (Once) Then
      Once = True
      Text1(0).SetFocus
   End If
   
   UpdateCaptions

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".m_t1_ThatTime") = vbRetry Then Resume Next

End Sub

Private Sub picMonth_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer, MaxDay As Integer
   On Error GoTo ProcedureError
    'Determine which date is being clicked
    i = Weekday(DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)) - 1
    i = (((X \ m_nGridWidth) + 1) + ((Y \ m_nGridHeight) * GRID_COLS)) - i
    'Get last day of current month
    MaxDay = Day(DateAdd("d", -1, DateSerial(Year(m_CurrDate), Month(m_CurrDate) + 1, 1)))
    If i >= 1 And i <= MaxDay Then
        SetNewDate DateSerial(Year(m_CurrDate), Month(m_CurrDate), i)
    End If

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox("picMonth_MouseDown") = vbRetry Then Resume Next

End Sub

Private Sub picMonth_Paint()
    Dim i As Integer, j As Integer, X As Integer, Y As Integer
    Dim NumDays As Integer, CurrPos As Integer, bCurrMonth As Boolean
    Dim MonthStart As Date, Buffer As String
    Dim PrevMonth As Date, PrevDays As Integer

   Dim iMon As Integer
   Dim iYear As Integer
   
   On Error GoTo ProcedureError
   
   iMon = Month(m_CurrDate) - 1 'zero base
   iYear = Year(m_CurrDate)
   If Combo1(0).ListIndex <> iMon Then
      Combo1(0).ListIndex = iMon
   End If
   If Combo1(1).ListIndex + 1900 <> iYear Then
      Combo1(1).ListIndex = iYear - 1900
   End If

    'Determine if this month is today's month
    If Month(m_CurrDate) = Month(Date) And Year(m_CurrDate) = Year(Date) Then
        bCurrMonth = True
    End If
    'Get first date in the month
    MonthStart = DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)
    PrevMonth = DateAdd("m", -1, MonthStart)
    'Number of days in the month
    NumDays = DateDiff("d", MonthStart, DateAdd("m", 1, MonthStart))
    PrevDays = DateDiff("d", PrevMonth, DateAdd("m", 1, PrevMonth))
    'Get first weekday in the month (0 - based)
    j = Weekday(MonthStart) - 1
    'Tweak for 1-based For/Next index
    j = j - 1
    'Show current month/year

    lblMonth = Combo1(0).Text & Format$(m_CurrDate, " yyyy")
    'Clear existing data
    picMonth.Cls
    'Tile.TileArea picMonth.hdc, 0, 0, picMonth.Width, picMonth.Height
    'Display dates for current month as black
    'Display Prev/Next month days as greyed
     For i = 1 To 41 - j         'NumDays
        CurrPos = i + j
        X = (CurrPos Mod GRID_COLS) * m_nGridWidth
        Y = (CurrPos \ GRID_COLS) * m_nGridHeight
        'Show date as bold if today's date
        If i <= NumDays Then
            If bCurrMonth And i = Day(Date) Then
                picMonth.Font.Bold = True
            Else
                picMonth.Font.Bold = False
            End If
            Buffer = CStr(i)
            picMonth.ForeColor = 0
        Else  'Add next month as 1 to ... (greyed)
            Buffer = CStr(i - NumDays)
            picMonth.ForeColor = &H808080
            picMonth.Font.Bold = False
        End If
        'Center date within "date cell"
        picMonth.CurrentX = X + ((m_nGridWidth - picMonth.TextWidth(Buffer)) / 2)
        picMonth.CurrentY = Y + ((m_nGridHeight - picMonth.TextHeight(Buffer)) / 2)
        'Print date
        picMonth.Print Buffer;
    Next i

    For i = 0 To j 'Add previous month days (greyed)
        CurrPos = i
        X = (CurrPos Mod GRID_COLS) * m_nGridWidth
        Y = (CurrPos \ GRID_COLS) * m_nGridHeight
        Buffer = CStr(PrevDays - j + i)
        picMonth.CurrentX = X + ((m_nGridWidth - picMonth.TextWidth(Buffer)) / 2)
        picMonth.CurrentY = Y + ((m_nGridHeight - picMonth.TextHeight(Buffer)) / 2)
        picMonth.Print Buffer;
    Next
    'Indicate selected date
    DrawSelectionBox True

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".picMonth_Paint") = vbRetry Then Resume Next

End Sub
Private Sub DrawSelectionBox(bSelected As Boolean)
    Dim clrTopLeft As Long, clrBottomRight As Long
    Dim i As Integer, X As Integer, Y As Integer
    
    On Error GoTo ProcedureError
    
    'Set highlight and shadow colors
    If bSelected Then
        clrTopLeft = vbBlack      'vbButtonShadow
        clrBottomRight = vbBlack  'vb3DHighlight
    Else
        clrTopLeft = picMonth.BackColor     'was vbButtonFace
        clrBottomRight = picMonth.BackColor 'was vbButtonFace
    End If
    'Compute location for current date
    i = Weekday(DateSerial(Year(m_CurrDate), Month(m_CurrDate), 1)) - 1
    i = i + (Day(m_CurrDate) - 1)
    X = (i Mod GRID_COLS) * m_nGridWidth
    Y = (i \ GRID_COLS) * m_nGridHeight
    'Draw box around date
    picMonth.Line (X, Y + m_nGridHeight)-Step(0, -m_nGridHeight), clrTopLeft
    picMonth.Line -Step(m_nGridWidth, 0), clrTopLeft
    picMonth.Line -Step(0, m_nGridHeight), clrBottomRight
    picMonth.Line -Step(-m_nGridWidth, 0), clrBottomRight

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".DrawSelectionBox") = vbRetry Then Resume Next

End Sub


Private Sub UpdateCaptions()
   On Error GoTo ProcedureError

   MyTime = TimeSerial(Val(Text1(0)), Val(Text1(1)), Val(Text1(2)))
   NewDateTime = Int(m_CurrDate) + MyTime
   
   lblShortDate.Caption = FormatDateTime(NewDateTime, vbGeneralDate)
   lblLongDate.Caption = FormatDateTime(NewDateTime, vbLongDate) & " " & _
                    FormatDateTime(NewDateTime, vbLongTime)

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".UpdateCaptions") = vbRetry Then Resume Next

End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  Dim i As Integer
   
   On Error Resume Next
   
   Select Case KeyCode
      Case 39 'Right
         Index = (Index + 1) Mod 3
         Text1(Index).SetFocus
         'For i = 0 To 2
         '   Text1(Index).Font.Bold = i = Index
         'Next
      Case 37 'Left
         Index = Index - 1
         If Index < 0 Then Index = 2
         Text1(Index).SetFocus
         'For i = 0 To 2
         '   Text1(Index).Font.Bold = i = Index
         'Next
      Case 38 'Up
         VScroll(Index).Value = VScroll(Index).Value - 1
         VScroll_Change (Index)
      Case 40 'Down
         VScroll(Index).Value = VScroll(Index).Value + 1
         VScroll_Change (Index)
   End Select
End Sub

Private Sub VScroll_Change(Index As Integer)
   On Error GoTo ProcedureError
    Text1(Index) = VScroll(Index).Max - VScroll(Index).Value

ProcedureExit:
  Exit Sub

ProcedureError:
  If ErrMsgBox(Me.Name & ".VScroll_Change.Idx " & Index) = vbRetry Then Resume Next

End Sub
