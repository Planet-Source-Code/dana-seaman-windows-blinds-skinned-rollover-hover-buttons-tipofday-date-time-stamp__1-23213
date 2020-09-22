VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "ButtonEx Control by J. Pearson"
   ClientHeight    =   4530
   ClientLeft      =   4530
   ClientTop       =   1245
   ClientWidth     =   6810
   Icon            =   "ButtonEx.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "ButtonEx.frx":0ABA
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "DateTime Stamp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3780
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Tip Of Day"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3180
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   3060
      TabIndex        =   9
      Top             =   360
      Width           =   2115
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   360
      Index           =   9
      Left            =   1680
      TabIndex        =   0
      Top             =   2700
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   635
      Appearance      =   2
      Caption         =   "Cancel"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":869C
      SkinUp          =   "ButtonEx.frx":920E
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   300
      Index           =   10
      Left            =   300
      TabIndex        =   1
      Top             =   3300
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   529
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":9D80
      SkinUp          =   "ButtonEx.frx":A772
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx2 
      Height          =   300
      Index           =   0
      Left            =   5460
      TabIndex        =   2
      Top             =   240
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      TransparentColor=   12632256
      SkinOver        =   "ButtonEx.frx":B164
      SkinUp          =   "ButtonEx.frx":B746
      TransparentColor=   12632256
   End
   Begin Project1.ButtonEx ButtonEx2 
      Height          =   300
      Index           =   1
      Left            =   5820
      TabIndex        =   3
      Top             =   240
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      TransparentColor=   14011845
      SkinOver        =   "ButtonEx.frx":BD28
      SkinUp          =   "ButtonEx.frx":C30A
      TransparentColor=   14011845
   End
   Begin Project1.ButtonEx ButtonEx2 
      Height          =   300
      Index           =   2
      Left            =   6180
      TabIndex        =   4
      Top             =   240
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      TransparentColor=   12632256
      SkinOver        =   "ButtonEx.frx":C8EC
      SkinUp          =   "ButtonEx.frx":CECE
      TransparentColor=   12632256
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   6
      Left            =   300
      TabIndex        =   5
      Top             =   2100
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
      Caption         =   "Cancel"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":D4B0
      SkinUp          =   "ButtonEx.frx":DF8A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   345
      Index           =   4
      Left            =   300
      TabIndex        =   6
      Top             =   1440
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   2
      Caption         =   "Apply"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":EA64
      SkinUp          =   "ButtonEx.frx":F58A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   300
      Index           =   14
      Left            =   3060
      TabIndex        =   7
      Top             =   3240
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":100B0
      SkinUp          =   "ButtonEx.frx":10832
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   8
      Top             =   1440
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      Appearance      =   2
      Caption         =   "Cancel"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":10FB4
      SkinUp          =   "ButtonEx.frx":11B0E
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Appearance      =   2
      Caption         =   "Apply"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TransparentColor=   16711935
      SkinOver        =   "ButtonEx.frx":12668
      SkinUp          =   "ButtonEx.frx":13226
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   360
      Index           =   12
      Left            =   300
      TabIndex        =   11
      Top             =   3840
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      Appearance      =   2
      Caption         =   "Apply"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":13DE4
      SkinUp          =   "ButtonEx.frx":148F6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   345
      Index           =   2
      Left            =   300
      TabIndex        =   12
      Top             =   900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":15408
      SkinUp          =   "ButtonEx.frx":15F2E
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   300
      Index           =   15
      Left            =   3540
      TabIndex        =   13
      Top             =   3240
      Width           =   450
      _ExtentX        =   794
      _ExtentY        =   529
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":16A54
      SkinUp          =   "ButtonEx.frx":171D6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   17
      Left            =   4140
      TabIndex        =   14
      Top             =   3840
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":17958
      SkinUp          =   "ButtonEx.frx":18222
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   13
      Left            =   1680
      TabIndex        =   15
      Top             =   3840
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":18AEC
      SkinUp          =   "ButtonEx.frx":195C6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   16
      Left            =   3240
      TabIndex        =   16
      Top             =   3840
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":1A0A0
      SkinUp          =   "ButtonEx.frx":1A96A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   0
      Left            =   300
      TabIndex        =   17
      Top             =   360
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":1B234
      SkinUp          =   "ButtonEx.frx":1BD0E
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   3
      Left            =   1680
      TabIndex        =   18
      Top             =   900
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":1C7E8
      SkinUp          =   "ButtonEx.frx":1D2C2
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   360
      Index           =   7
      Left            =   1680
      TabIndex        =   19
      Top             =   2100
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
      Appearance      =   2
      Caption         =   "Apply"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":1DD9C
      SkinUp          =   "ButtonEx.frx":1E8AE
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   8
      Left            =   300
      TabIndex        =   20
      Top             =   2700
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   582
      Appearance      =   2
      Caption         =   "Cancel"
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":1F3C0
      SkinUp          =   "ButtonEx.frx":1FE9A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   11
      Left            =   1680
      TabIndex        =   21
      Top             =   3300
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":20974
      SkinUp          =   "ButtonEx.frx":2144E
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   210
      Index           =   27
      Left            =   3780
      TabIndex        =   22
      Top             =   2820
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   370
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":21F28
      SkinUp          =   "ButtonEx.frx":2221A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   210
      Index           =   24
      Left            =   4140
      TabIndex        =   23
      Top             =   2820
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   370
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":2250C
      SkinUp          =   "ButtonEx.frx":227FE
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   210
      Index           =   25
      Left            =   4500
      TabIndex        =   24
      Top             =   2820
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   370
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":22AF0
      SkinUp          =   "ButtonEx.frx":22DE2
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   210
      Index           =   26
      Left            =   3420
      TabIndex        =   25
      Top             =   2820
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   370
      Appearance      =   2
      Caption         =   ""
      ForeColor       =   16777215
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
      SkinOver        =   "ButtonEx.frx":230D4
      SkinUp          =   "ButtonEx.frx":233C6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   18
      Left            =   4200
      TabIndex        =   26
      Top             =   3240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":236B8
      SkinUp          =   "ButtonEx.frx":23D72
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   19
      Left            =   4680
      TabIndex        =   27
      Top             =   3240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":2442C
      SkinUp          =   "ButtonEx.frx":24AE6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   330
      Index           =   20
      Left            =   5460
      TabIndex        =   28
      Top             =   720
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   582
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":251A0
      SkinUp          =   "ButtonEx.frx":25C7A
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   420
      Index           =   21
      Left            =   5460
      TabIndex        =   29
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   741
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":26754
      SkinUp          =   "ButtonEx.frx":273F6
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   345
      Index           =   22
      Left            =   5460
      TabIndex        =   30
      Top             =   1920
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      Appearance      =   2
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
      SkinOver        =   "ButtonEx.frx":28098
      SkinUp          =   "ButtonEx.frx":28BBE
      TransparentColor=   16711935
   End
   Begin Project1.ButtonEx ButtonEx 
      Height          =   360
      Index           =   23
      Left            =   5520
      TabIndex        =   31
      Top             =   2460
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   635
      Appearance      =   2
      Caption         =   "Cancel"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      TransparentColor=   16711935
      SkinOver        =   "ButtonEx.frx":296E4
      SkinUp          =   "ButtonEx.frx":2A936
      TransparentColor=   16711935
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00404040&
      BorderStyle     =   3  'Dot
      BorderWidth     =   4
      Height          =   4470
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   6750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ButtonEx6_Click()
   Unload Me
End Sub

Private Sub ButtonEx_Click(Index As Integer)
   PlaySound 61
   Showevent Index, "Click"
End Sub
Private Sub Showevent(Idx As Integer, Nam As String)
   List1.AddItem "Button " & Format(Idx, "00 ") & Nam
   List1.ListIndex = List1.ListCount - 1
End Sub

Private Sub ButtonEx_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Showevent Index, "Down"
End Sub

Private Sub ButtonEx_MouseEnter(Index As Integer)
   PlaySound 60
   Showevent Index, "Enter"
End Sub

Private Sub ButtonEx_MouseExit(Index As Integer)
   Showevent Index, "Exit"
End Sub

Private Sub ButtonEx_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Showevent Index, "Up"
End Sub

Private Sub ButtonEx2_Click(Index As Integer)
   PlaySound 61
   Select Case Index
      Case 0
         Showevent Index, "Minimized"
         Me.WindowState = vbMinimized
      Case 1
         Showevent Index, "Maximized"
      Case 2
         Showevent Index, "Unload"
         If MsgBox("Do you really want to exit?", vbYesNo) = vbYes Then
            Unload Me
         End If
   End Select
End Sub

Private Sub ButtonEx2_MouseEnter(Index As Integer)
   PlaySound 60
   Select Case Index
      Case 0
         Showevent Index, "Min Enter"
      Case 1
         Showevent Index, "Max Enter"
      Case 2
         Showevent Index, "Close Enter"
   End Select
End Sub

Private Sub ButtonEx2_MouseExit(Index As Integer)
   Select Case Index
      Case 0
         Showevent Index, "Min Exit"
      Case 1
         Showevent Index, "Max Exit"
      Case 2
         Showevent Index, "Close Exit"
   End Select
End Sub

Private Sub Command1_Click()
   frmTipOfDay.Show vbModal, Me
End Sub

Private Sub Command2_Click()
   frmDate.Show vbModal, Me
End Sub

Private Sub Form_Load()
   MakeFormRounded Me, 6   'Round the corners
   Shape2.Move 2, 2        'Align shape(border) to form
   InSound = True          'Enable sounds
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   FormDrag Me
End Sub
Public Sub UpdateDates()
   Dim L4         As Long
   Dim hFile      As Long
   Dim FT_CREATE  As Currency
   Dim FT_ACCESS  As Currency
   Dim FT_WRITE   As Currency
   Dim LOC_TIME   As Currency
   Dim UTC_TIME   As Currency
   
   Const FILE_SHARE_READ = &H1
   Const FILE_SHARE_WRITE = &H2
   Const GENERIC_READ = &H80000000
   Const GENERIC_WRITE = &H40000000
   Const OPEN_EXISTING = 3
   Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
   
   On Error GoTo ProcedureError
   
'*** DateTimeStamp code...
   
'If HowManyTags = 0 Then ' No object selected
'   Msg = GetResourceString(1023) & GetResourceString(1024) & GetResourceString(1025)
'   MsgBox Msg
'ElseIf WhichDates = 0 Then ' No Date selected
'   Msg = GetResourceString(1023) & GetResourceString(1112) & " " & GetResourceString(1025)
'   MsgBox Msg
'Else
'   LOC_TIME = VbTimeToWin32Local(NewDateTime) ' VbDate > Local
'   LocalFileTimeToFileTime LOC_TIME, UTC_TIME 'Local > UTC
'   With Grid
'      .Redraw = False
'      For L4 = 1 To .Rows
'         If Grid.CellSelected(L4, 1) Then 'Assumes we're in Row mode
'            Source = SourcePath & .CellText(L4, .ColumnIndex(nam_))
'             hFile = CreateFile(Source, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0&)
'            'If we were able to open file, change it's timestamp.
'          If hFile <> INVALID_HANDLE_VALUE Then
'             GetFileTime hFile, FT_CREATE, FT_ACCESS, FT_WRITE
'             If WhichDates And 1 Then
'                  FT_CREATE = UTC_TIME
'                  .CellText(L4, .ColumnIndex(cre_)) = NewDateTime
'               End If
'               If WhichDates And 2 Then
'                  FT_ACCESS = UTC_TIME
'                  .CellText(L4, .ColumnIndex(acc_)) = NewDateTime
'               End If
'               If WhichDates And 4 Then
'                  FT_WRITE = UTC_TIME
'                  .CellText(L4, .ColumnIndex(mod_)) = NewDateTime
'                  .CellText(L4, .ColumnIndex(tim_)) = NewDateTime
'               End If
'               SetFileTime hFile, FT_CREATE, FT_ACCESS, FT_WRITE
'               CloseHandle hFile
'            End If
'         End If
'      Next
'      .Redraw = True
'   End With
'End If

ProcedureExit:
  Exit Sub
ProcedureError:
  If ErrMsgBox(Me.Name & ".UpdateDates") = vbRetry Then Resume Next

End Sub
Private Sub Form_Unload(Cancel As Integer)
  List1.Clear
End Sub
