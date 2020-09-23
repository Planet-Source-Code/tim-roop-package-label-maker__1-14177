VERSION 5.00
Begin VB.Form fmPrint 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2.917
   ScaleMode       =   5  'Inch
   ScaleWidth      =   4.438
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Image Image2 
      Height          =   1440
      Left            =   240
      Top             =   1800
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   13
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   11760
      Width           =   1440
   End
   Begin VB.Image Image1 
      Height          =   1440
      Index           =   6
      Left            =   240
      Stretch         =   -1  'True
      Top             =   11760
      Width           =   1440
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1440
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1440
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   13
      Left            =   8520
      TabIndex        =   4
      Top             =   12840
      Width           =   2400
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   6
      Left            =   2400
      TabIndex        =   3
      Top             =   12840
      Width           =   2400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   1800
      TabIndex        =   2
      Top             =   11880
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   13
      Left            =   7920
      TabIndex        =   1
      Top             =   11880
      Width           =   3735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "fmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

