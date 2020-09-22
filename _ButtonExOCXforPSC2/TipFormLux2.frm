VERSION 5.00
Begin VB.Form frmTipOfDay 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5910
   Icon            =   "TipFormLux2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TipFormLux2.frx":000C
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   2460
      TabIndex        =   3
      Top             =   4020
      Width           =   195
   End
   Begin Project1.ButtonEx Button 
      Height          =   330
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   1800
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
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "Next"
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
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "Previous"
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
      TabIndex        =   8
      Top             =   3600
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Caption         =   "Random"
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
      Height          =   4455
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   5850
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't show at startup."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4020
      Width           =   1995
   End
   Begin VB.Label lbltiptext 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "That this is your first tip?, click next button for more ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2520
      TabIndex        =   2
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Did you Know ? . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3000
      TabIndex        =   1
      Top             =   1200
      Width           =   1950
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tip of the day"
      BeginProperty Font 
         Name            =   "Black Chancery"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2520
      TabIndex        =   0
      Top             =   180
      Width           =   2835
   End
End
Attribute VB_Name = "frmTipOfDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurRgn, TempRgn As Long
Dim Tips As New Collection ' The in-memory database of tips.
Dim CurrentTip As Long ' Index in collection of tip currently being displayed.

Const TIP_FILE = "TIPOFDAY.TXT" ' Name of tips file

Public Sub DisplayCurrentTip()
    On Error Resume Next
    If Tips.Count > 0 Then
        lbltiptext.Caption = Tips.Item(CurrentTip)
    End If
End Sub

Private Sub DoNextTip()

    CurrentTip = CurrentTip + 1
    If CurrentTip > Tips.Count Then
        CurrentTip = 1 'Wrap
    End If
    DisplayCurrentTip
    
End Sub

Public Sub DoPreviousTip()
    
    CurrentTip = CurrentTip - 1
    If CurrentTip < 1 Then
       CurrentTip = Tips.Count 'Wrap
    End If
    DisplayCurrentTip
    
End Sub

Private Sub DoRandomTip()
    Static LastTip As Integer
    'Random but don't allow same tip twice in a row
    Do
        CurrentTip = Int((Tips.Count * Rnd) + 1)
        If CurrentTip <> LastTip Then
            LastTip = CurrentTip
            Exit Do
        End If
    Loop
    DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    InFile = FreeFile ' Obtain the next free file descriptor.
    
    If (sFile = "") Or (Dir(sFile) = "") Then ' File exists?
        Exit Function
    End If
 
    Open sFile For Input As InFile ' Read the collection from a text file.
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    DoNextTip
    
    LoadTips = True
    
End Function

Private Sub Button_Click(Index As Integer)
   PlaySound 61
   Select Case Index
      Case 0
         Unload Me
      Case 1
         DoNextTip
      Case 2
         DoPreviousTip
      Case 3
         DoRandomTip
   End Select
   
End Sub

Private Sub Button_MouseEnter(Index As Integer)
   PlaySound 60
End Sub

Private Sub Check1_Click()
   SaveSetting App.EXEName, "Options", "Show Tips at Startup", Check1.Value
End Sub

Private Sub Form_Load()

    Dim ShowAtStartup As Long
   
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 0)
    If ShowAtStartup = 1 Then
        Unload Me
        Exit Sub
    End If
   
    MakeFormRounded Me, 5   'Round the corners
    ResourceSkinNum = 50    'Base index for resource
    SkinButtons Me
   
    ' Set the checkbox, this will force the value to be written back out to the registry
    Check1.Value = vbUnchecked
    
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lbltiptext.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub

Private Sub Label4_Click()
    Check1_Click
End Sub
