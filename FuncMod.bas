Attribute VB_Name = "Module1"
Type RecLabel
    TName As String * 35
    Picpath As String * 75
    PartNum As String * 10
    Color As String * 12
End Type
Global FileNum As Integer
Global Record As RecLabel
Global CurrentRec As Integer
Global RecLength As Long
Global indexVal As Integer
Global indx(150) As Integer
Global LastRec As Long
Global FileNam As String
Global Position As Long
Global FileNum2 As Integer
Global SFile As String
Global PicString As String
Global NameString As String
Global NameStringB As String
Global PicStr(150) As String
Global NameStr(150) As String
Global PartStr(150) As String
Global ColrStr(150) As String
Global AddItem As Boolean
Global Item As Integer
Global PNumber As String
Global InitVar As String
Global ColorVal As String

Public Sub SetColor()
    
    Select Case ColorVal
        
        Case Is = "Red"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Red.jpg")
        Case Is = "Blue"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Blue.jpg")
        Case Is = "Green"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Green.jpg")
        Case Is = "Pink"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Pink.jpg")
        Case Is = "Gray"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Gray.jpg")
        Case Is = "Black"
            fmPrint.Image2.Picture = LoadPicture("C:\Program Files\LabelMaker\Black.jpg")
    End Select
    
        
End Sub
