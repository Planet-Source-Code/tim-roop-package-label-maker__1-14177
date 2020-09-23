VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fmEdit 
   BackColor       =   &H00800000&
   Caption         =   "Add Item"
   ClientHeight    =   2325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "fmEdit.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2325
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboColor 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtPartNum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   2
      Top             =   1320
      Width           =   1950
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancel"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1845
      Width           =   1425
   End
   Begin MSComDlg.CommonDialog dlgB 
      Left            =   3360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAddItem 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Item"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1845
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddPic 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Picture..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   225
      Width           =   1365
   End
   Begin VB.TextBox txtItem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1170
      TabIndex        =   1
      Top             =   720
      Width           =   4440
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Part #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   270
      TabIndex        =   7
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label lblPicString 
      BackColor       =   &H00800000&
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
      Height          =   375
      Left            =   1635
      TabIndex        =   6
      Top             =   225
      Width           =   3990
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   285
      TabIndex        =   5
      Top             =   765
      Width           =   780
   End
End
Attribute VB_Name = "fmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub cmdAddItem_Click()
    
    NameString = txtItem.Text
    PicString = lblPicString.Caption
    PNumber = txtPartNum.Text
    ColorVal = cboColor.Text
    'i = i + 1
    If txtItem.Text = "" Then
        Dim Result
        Result = MsgBox("Please enter a name for the label.", , "No Label Name")
        Exit Sub
    ElseIf txtPartNum.Text = "" Then
        Dim Result2
        Result2 = MsgBox("Please enter a part number for for this label.", , "No Part Number")
        Exit Sub
    End If
    
    Call fmMain.Save2File
    AddItem = True
    fmMain.lstItems.Refresh
    Unload Me
    
End Sub

Private Sub cmdAddPic_Click()
    
        On Error GoTo OpenError
    
        'Do While FileNum2 > 0
            'Close FileNum2
            'FileNum2 = FileNum2 - 1
        'Loop
        
        'To Do
        'set the flags and attributes of the
        'common dialog control
        SFile = "All Image Files (*.*)|*.*|"
        SFile = SFile + "BMP Files (*.bmp)|*.bmp|  "
        SFile = SFile + "GIF Files (*.gif)|*.gif|  "
        SFile = SFile + "JPG Files (*.jpg)|*.jpg|  "
        dlgB.Filter = SFile
        dlgB.InitDir = "C:\Program Files\LabelMaker"
        dlgB.FilterIndex = 1
        dlgB.ShowOpen
        PicString = dlgB.FileName
        'dlgB.Flags = cdlOFNCreatePrompt
        If Len(PicString) = 0 Then
            Exit Sub
        End If
        lblPicString.Caption = PicString
        'FileNum2 = FreeFile

OpenError:
    Exit Sub

End Sub

Private Sub cmdCancel_Click()
    
    txtItem.Text = ""
    txtPartNum.Text = ""
    lblPicString.Caption = ""
    cboColor.Clear
    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    cmdAddPic.SetFocus
    cboColor.AddItem "NONE"
    cboColor.AddItem "Green"
    cboColor.AddItem "Blue"
    cboColor.AddItem "Red"
    cboColor.AddItem "Pink"
    cboColor.AddItem "Black"
    cboColor.AddItem "Gray"
    cboColor.Text = "Select a border"
    
End Sub
