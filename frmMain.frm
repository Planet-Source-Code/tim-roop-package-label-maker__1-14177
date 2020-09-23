VERSION 5.00
Begin VB.Form fmMain 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Label Maker"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmMain.frx":0000
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
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":0444
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
      Picture         =   "frmMain.frx":0446
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
    
        frmEdit.Show
        
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
    TmpLbl = "\\Renntsvr1\Mfg\7th Street\7th Street Files\LabelMaker\Labels.tmp"
    Open TmpLbl For Random As TmpFileNum Len = RecLength
    Position = 1
    TmpRecNum = 1
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
    Image1.Picture = LoadPicture("\\Renntsvr1\Mfg\7th Street\7th Street Files\LabelMaker\NoPicAvail.jpg")
    
End Sub

Private Sub cmdExit_Click()
    
    Unload Me
    
End Sub

Private Sub cmdPrint_Click()
'Dim Index As Integer
'frmPrint.Show
'For Index = 0 To 13
    'frmPrint.Image1(Index).Picture = LoadPicture(PicString)
    'frmPrint.Label1(Index).Caption = NameString
    'frmPrint.Label3(Index).Caption = PNumber
'Next Index

'frmPrint.PrintForm
'Unload frmPrint
Dim Row As Single
Dim Col As Single
Dim BSString As String
Dim XL As Integer
Dim SetX
Dim SetY
Dim offX
Dim offY

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
        frmPrint.Image1(0).Stretch = True
        frmPrint.Image1(0).Picture = LoadPicture(PicString)
        frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\grey.jpg")
        If Right(PNumber, 1) = "N" Then
            If Col = 1 Then
                If Right(PNumber, 1) = "N" Then
                    If Left(PNumber, 4) = "#185" Then
                        frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\blue.jpg")
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    ElseIf Left(PNumber, 4) = "#143" Then
                        frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\pink.jpg")
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    Else
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    End If
                End If
            ElseIf Col = 2 Then
                If Right(PNumber, 1) = "N" Then
                    If Left(PNumber, 4) = "#185" Then
                        frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\blue.jpg")
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    ElseIf Left(PNumber, 4) = "#143" Then
                        frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\pink.jpg")
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    Else
                        Printer.CurrentY = (Row * 34) + offY
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                        Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                        SetX = Printer.CurrentX
                        Printer.CurrentY = (Row * 34) + (offY + 1)
                        SetY = Printer.CurrentY
                        Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                    End If
                End If
            End If
        Else
            If Col = 1 Then
                If Left(PNumber, 4) = "#185" Then
                    frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\blue.jpg")
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                ElseIf Left(PNumber, 4) = "#143" Then
                    frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\pink.jpg")
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                Else
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 25, 25
                End If
            ElseIf Col = 2 Then
                If Left(PNumber, 4) = "#185" Then
                    frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\blue.jpg")
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                ElseIf Left(PNumber, 4) = "#143" Then
                    frmPrint.Image2.Picture = LoadPicture("\\Renntsvr1\mfg\7th street\7th street files\labelmaker\pink.jpg")
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image2.Picture, SetX, SetY, 25, 25
                    Printer.CurrentX = ((Col - 1) * 106) + (offX + 1)
                    SetX = Printer.CurrentX
                    Printer.CurrentY = (Row * 34) + (offY + 1)
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 23, 23
                Else
                    Printer.CurrentY = (Row * 34) + offY
                    SetY = Printer.CurrentY
                    Printer.PaintPicture frmPrint.Image1(0).Picture, SetX, SetY, 25, 25
                End If
            End If
        End If

    Next Col

Next Row


For Row = 0 To 6

    For Col = 1 To 2
        Printer.CurrentY = (Row * 34) + offY
        SetY = Printer.CurrentY
        If Col = 1 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            If Len(NameString) > 17 Then
                NameStringB = Replace(NameString, " Box Set", "")
                XL = (Len(NameStringB) / 2) - 3
                XP = (Len(NameStringB) / 2) - 6
                frmPrint.Label1(0).Caption = NameStringB
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontBold = False
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            Else
                XL = (Len(NameString) / 2) - 3
                XP = (Len(NameString) / 2) - 6
                frmPrint.Label1(0).Caption = NameString
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontBold = False
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        ElseIf Col = 2 Then
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            If Len(NameString) > 17 Then
                NameStringB = Replace(NameString, " Box Set", "")
                XL = (Len(NameStringB) / 2) - 3
                XP = (Len(NameStringB) / 2) - 6
                frmPrint.Label1(0).Caption = NameStringB
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontBold = False
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            Else
                XL = (Len(NameString) / 2) - 3
                XP = (Len(NameString) / 2) - 6
                frmPrint.Label1(0).Caption = NameString
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontBold = False
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        End If
        
    Next Col

Next Row

If Len(NameString) > 17 Then
    
    For Row = 0 To 6
    
        For Col = 1 To 2
            
            
            If Col = 1 Then
                Printer.CurrentY = (Row * 34) + (offY + 8)
                SetY = Printer.CurrentY
                Printer.ScaleMode = 6
                Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
                Printer.ScaleMode = 4
                Printer.CurrentX = Printer.CurrentX
                BSString = "Box Set"
                frmPrint.Label1(0).Caption = BSString
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            ElseIf Col = 2 Then
                Printer.CurrentY = (Row * 34) + (offY + 8)
                Printer.ScaleMode = 6
                Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
                Printer.ScaleMode = 4
                Printer.CurrentX = Printer.CurrentX
                BSString = "Box Set"
                frmPrint.Label1(0).Caption = BSString
                Printer.FontSize = 18
                Printer.Font = "Bauhaus Md BT"
                Printer.FontItalic = True
                frmPrint.Label1(0).Alignment = 2
                Printer.Print frmPrint.Label1(0).Caption
                Printer.ScaleMode = 6
            End If
        
        Next Col

    Next Row

End If


For Row = 0 To 6
    
    For Col = 1 To 2
        
        
        If Col = 1 Then
            Printer.CurrentY = (Row * 34) + (offY + 15)
            SetY = Printer.CurrentY
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            Printer.CurrentX = Printer.CurrentX
            frmPrint.Label1(0).Caption = PNumber
            Printer.FontSize = 24
            Printer.Font = "Bauhaus Md BT"
            Printer.FontItalic = False
            frmPrint.Label1(0).Alignment = 2
            Printer.Print frmPrint.Label1(0).Caption
            Printer.ScaleMode = 6
        ElseIf Col = 2 Then
            Printer.CurrentY = (Row * 34) + (offY + 15)
            Printer.ScaleMode = 6
            Printer.CurrentX = ((Col - 1) * 106) + (offX + 30)
            Printer.ScaleMode = 4
            Printer.CurrentX = Printer.CurrentX
            frmPrint.Label1(0).Caption = PNumber
            Printer.FontSize = 24
            Printer.Font = "Bauhaus Md BT"
            Printer.FontItalic = False
            frmPrint.Label1(0).Alignment = 2
            Printer.Print frmPrint.Label1(0).Caption
            Printer.ScaleMode = 6
        End If
        
    Next Col

Next Row

Printer.EndDoc
Unload frmPrint


End Sub



Private Sub cmdSort_Click()
    
    FileNum = FreeFile
    RecLength = Len(Record)
    indexVal = 1
    Position = 1
    FileNam = "\\Renntsvr1\Mfg\7th Street\7th Street Files\LabelMaker\Labels.nlb"
    Open FileNam For Random As FileNum Len = RecLength
    LastRec = FileLen(FileNam) / RecLength

    If cmdSort.Caption = "Sort by Part Number" Then
        lstItems.Clear
        lblPart.Caption = ""
        Image1.Picture = LoadPicture("\\Renntsvr1\mfg\7th Street\7th Street Files\LabelMaker\NoPicAvail.jpg")
        For Position = 1 To LastRec
            Get #FileNum, Position, Record
            lstItems.AddItem (Trim(Record.PartNum) & "     " & Trim(Record.TName)) & "                            " & Position
        Next Position
        cmdSort.Caption = "Sort by Label Name"
    Else
        lstItems.Clear
        lblPart.Caption = ""
        Image1.Picture = LoadPicture("\\Renntsvr1\mfg\7th Street\7th Street Files\LabelMaker\NoPicAvail.jpg")
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
    FileNam = "\\Renntsvr1\Mfg\7th Street\7th Street Files\LabelMaker\Labels.nlb"
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
End Sub

Private Sub LstItems_Click()
    
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
    
    Put #FileNum, Position, Record
    Close #FileNum
    NameStr(Position) = NameString
    PicStr(Position) = PicString
    PartStr(Position) = PNumber

End Sub

Public Sub GetFromFile()

    FileNum = FreeFile
    Open FileNam For Random As FileNum Len = RecLength
    
    Get #FileNum, Position, Record
    
    NameString = Trim(Record.TName)
    PicString = Trim(Record.Picpath)
    PNumber = Trim(Record.PartNum)
    
    Close #FileNum
    
    
End Sub

