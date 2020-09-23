VERSION 5.00
Begin VB.Form fmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Package Label Maker"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "fmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSort 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sort by Label Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFC0C0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
      Width           =   2265
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Print Labels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteLbl 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Delete Label item from Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddLabel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Label item to Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.ListBox lstItems 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3660
      ItemData        =   "fmMain.frx":0442
      Left            =   120
      List            =   "fmMain.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label lblPart 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5160
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   5280
      Picture         =   "fmMain.frx":0446
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1500
   End
End
Attribute VB_Name = "fmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAddLabel_Click()
    
    'Show form to edit database
    fmEdit.Show
        
End Sub

Private Sub cmdDeleteLbl_Click()
    Dim TmpFileNum As Integer
    Dim TmpRecord As RecLabel
    Dim TmpRecNum As Long
    'If Labels.tmp already exists, delete it.
    If Dir("Labels.tmp") = "Labels.tmp" Then
        Kill "Labels.tmp"
    End If
    FileNum = FreeFile
    Open FileNam For Random As FileNum Len = RecLength
    TmpFileNum = FreeFile
    TmpLbl = "C:\Program Files\LabelMaker\Labels.tmp"
    Open TmpLbl For Random As TmpFileNum Len = RecLength
    TmpRecNum = 1
    
    'If the nothing is selected on the listbox when the "Delete Label" is pressed,
    'Position will = 0.
    'Trap the error, display a message box, and exit the routine.
    If Position = 0 Then
        MsgBox ("Please select a record from the list to delete.")
        Exit Sub
    End If
    Do While Position < LastRec + 1
        If Position <> CurrentRec Then
            Get #FileNum, Position, TmpRecord
            Put #TmpFileNum, TmpRecNum, TmpRecord
            TmpRecNum = TmpRecNum + 1
        End If
        Position = Position + 1
    Loop
    
    Close FileNum
    Kill FileNam
    
    Close TmpFileNum
    Name TmpLbl As FileNam
    
    Open FileNam For Random As FileNum Len = RecLength
    LastRec = LastRec - 1
    
    If LastRec = 0 Then LastRec = 1
    
    If CurrentRec > LastRec Then
        CurrentRec = LastRec
    End If
    
    Position = lstItems.ListIndex
    lstItems.RemoveItem Position
    
    Close FileNum
    Image1.Picture = LoadPicture("C:\Program Files\LabelMaker\NoPicAvail.jpg")
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()
Dim Row As Single
Dim Col As Single
Dim BSString As String
Dim XL As Integer
Dim SetX
Dim SetY
Dim offX
Dim offY

'This section of code formats Image boxes on fmPrint to print
'7 rows deep and 2 columns of labels.
offX = 4
offY = 21
For Row = 0 To 6
    For Col = 1 To 2
        If Col = 1 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + offX
            SetX = Printer.CurrentX
        ElseIf Col = 2 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + offX
            SetX = Printer.CurrentX
        End If
        fmPrint.Image1(0).Stretch = True
        fmPrint.Image1(0).Picture = LoadPicture(PicString)
            If Col = 1 Then
                If ColorVal <> "NONE" Then
                    Call SetColor
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                Else
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image1(0).Picture, SetX, SetY, 25, 25
                End If
            ElseIf Col = 2 Then
                If ColorVal <> "NONE" Then
                    Call SetColor
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                Else
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture fmPrint.Image1(0).Picture, SetX, SetY, 25, 25
                End If
            End If

    Next Col

Next Row


'This section of code formats the Label Box on fmPrint for
'7 rows deep and columns wide. It also determines the length of
'the description name and if it's greater than 20 characters
'it will break the descriptive name in half.
For Row = 0 To 6

    For Col = 1 To 2
        Printer.CurrentY = (Row * 34) + offY
        SetY = Printer.CurrentY
        If Col = 1 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            If Len(NameString) > 20 Then
                XP = Len(NameString) / 2
                XL = InStr(XP, NameString, " ", vbTextCompare)
                XT = Len(NameString) - XL
                NameStringB = Right(NameString, XT)
                NameString = Replace(NameString, NameStringB, "")
                fmPrint.Label1(0).Caption = Trim(NameString)
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontBold = False
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            Else
                fmPrint.Label1(0).Caption = Trim(NameString)
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontBold = False
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        ElseIf Col = 2 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            If Len(NameString) > 20 Then
                XP = Len(NameString) / 2
                XL = InStr(XP, NameString, " ", vbTextCompare)
                XT = Len(NameString) - XL
                NameStringB = Right(NameString, XT)
                NameString = Replace(NameString, NameStringB, "")
                fmPrint.Label1(0).Caption = Trim(NameString)
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontBold = False
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            Else
                fmPrint.Label1(0).Caption = Trim(NameString)
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontBold = False
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        End If
        
    Next Col

Next Row

'This section will print the second part of the descriptive name if there
'is any.

If NameStringB <> "" Then
    
    For Row = 0 To 6
    
        For Col = 1 To 2
            
            
            If Col = 1 Then
                Printer.CurrentY = (Row * 34) + (offY + 8)
                SetY = Printer.CurrentY
                Printer.ScaleMode = 6
                Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
                Printer.ScaleMode = 4
                'Printer.CurrentX = Printer.CurrentX
                BSString = Trim(NameStringB)
                fmPrint.Label1(0).Caption = BSString
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            ElseIf Col = 2 Then
                Printer.CurrentY = (Row * 34) + (offY + 8)
                Printer.ScaleMode = 6
                Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
                Printer.ScaleMode = 4
                'Printer.CurrentX = Printer.CurrentX
                BSString = Trim(NameStringB)
                fmPrint.Label1(0).Caption = BSString
                Printer.FontSize = 18
                Printer.Font = "Arial"
                Printer.FontItalic = True
                fmPrint.Label1(0).Alignment = 2
                Printer.Print fmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        
        Next Col

    Next Row

End If

'This section of code prints 7 rows and 2 columns worth of
'Part Numbers
For Row = 0 To 6
    
    For Col = 1 To 2
        
        
        If Col = 1 Then
            Printer.CurrentY = (Row * 34) + (offY + 15)
            SetY = Printer.CurrentY
            Printer.ScaleMode = 6 'Set scale to Centimeters
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4 'Set scale to Characters
            'Printer.CurrentX = Printer.CurrentX
            fmPrint.Label1(0).Caption = PNumber
            Printer.FontSize = 24
            Printer.Font = "Arial"
            Printer.FontItalic = False
            fmPrint.Label1(0).Alignment = 2
            Printer.Print fmPrint.Label1(0).Caption
            Printer.ScaleMode = 6 'Set scale back to Cm
        ElseIf Col = 2 Then
            Printer.CurrentY = (Row * 34) + (offY + 15)
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            'Printer.CurrentX = Printer.CurrentX
            fmPrint.Label1(0).Caption = PNumber
            Printer.FontSize = 24
            Printer.Font = "Arial"
            Printer.FontItalic = False
            fmPrint.Label1(0).Alignment = 2
            Printer.Print fmPrint.Label1(0).Caption
            Printer.ScaleMode = 6
        End If
        
    Next Col

Next Row

Printer.EndDoc 'Send Printer object to your printer
Unload fmPrint 'Unload the form


End Sub



Private Sub cmdSort_Click()
    
    FileNum = FreeFile
    RecLength = Len(Record)
    indexVal = 1
    Position = 1
    FileNam = "C:\Program Files\LabelMaker\Labels.nlb"
    Open FileNam For Random As FileNum Len = RecLength
    LastRec = FileLen(FileNam) / RecLength

    If cmdSort.Caption = "Sort by Part Number" Then
        lstItems.Clear
        lblPart.Caption = ""
        Image1.Picture = LoadPicture("C:\Program Files\LabelMaker\NoPicAvail.jpg")
        For Position = 1 To LastRec
            Get #FileNum, Position, Record
            lstItems.AddItem (Trim(Record.PartNum) & "     " & Trim(Record.TName)) & "                            " & Position
        Next Position
        cmdSort.Caption = "Sort by Label Name"
    Else
        lstItems.Clear
        lblPart.Caption = ""
        Image1.Picture = LoadPicture("C:\Program Files\LabelMaker\NoPicAvail.jpg")
        For Position = 1 To LastRec
            Get #FileNum, Position, Record
            lstItems.AddItem (Trim(Record.TName) & "     " & Trim(Record.PartNum)) & "                            " & Position
        Next Position
        cmdSort.Caption = "Sort by Part Number"
    End If

End Sub

Private Sub Form_Activate()
    
    If AddItem = True Then
        If InitVar = "No items in database" Then
            lstItems.Clear
            InitVar = ""
        End If
        If cmdSort.Caption = "Sort by Label Name" Then
            lstItems.AddItem (PartStr(Position) & "     " & NameStr(Position)) & "                               " & Position
            Position = Position + 1
        Else
            lstItems.AddItem (NameStr(Position) & "     " & PartStr(Position)) & "                               " & Position
            Position = Position + 1
        End If
    End If
    AddItem = False
End Sub

Public Sub Form_Load()
    
    FileNum = FreeFile
    RecLength = Len(Record)
    indexVal = 1
    Position = 1
    
    If Dir("C:\Program Files\LabelMaker", vbDirectory) = "" Then
        ChDir ("C:\Program Files")
        MkDir ("LabelMaker")
    End If
    
    FileNam = "C:\Program Files\LabelMaker\Labels.nlb"
    Open FileNam For Random As FileNum Len = RecLength
    LastRec = FileLen(FileNam) / RecLength
    If LastRec = 0 Then
        lstItems.AddItem "No items in database"
        InitVar = "No items in database"
    ElseIf LastRec <> 0 Then
        For Position = 1 To LastRec
            Get #FileNum, Position, Record
            lstItems.AddItem (Trim(Record.PartNum) & "     " & Trim(Record.TName)) & "                            " & Position
        Next Position
    End If
    Close #FileNum
    Position = 0
End Sub

Private Sub LstItems_Click()
    If lstItems.Text = "No items in database" Then
        Exit Sub
    End If
    Position = Right(lstItems.Text, 3)
    Call GetFromFile
    CurrentRec = Position
    Image1.Picture = LoadPicture(PicString)
    lblPart.Caption = PNumber

End Sub

Public Sub Save2File()
    
    FileNum = FreeFile
    Open FileNam For Random As FileNum Len = RecLength

    LastRec = LastRec + 1
    Position = LastRec
    Record.TName = Trim(NameString)
    Record.Picpath = Trim(PicString)
    Record.PartNum = Trim(PNumber)
    Record.Color = Trim(ColorVal)
    
    'Store variable data to Labels.nlb
    Put #FileNum, Position, Record
    Close #FileNum
    NameStr(Position) = NameString
    PicStr(Position) = PicString
    PartStr(Position) = PNumber
    ColrStr(Position) = ColorVal

End Sub

Public Sub GetFromFile()

    FileNum = FreeFile
    Open FileNam For Random As FileNum Len = RecLength
    
    Get #FileNum, Position, Record
    'Retrieves record data from Labels.nlb
    NameString = Trim(Record.TName)
    PicString = Trim(Record.Picpath)
    PNumber = Trim(Record.PartNum)
    ColorVal = Trim(Record.Color)
    
    Close #FileNum
    
    
End Sub

