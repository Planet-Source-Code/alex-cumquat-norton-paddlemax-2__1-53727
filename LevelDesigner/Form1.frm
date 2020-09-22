VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PaddleMax Level Design"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "New"
      Height          =   330
      Left            =   2700
      TabIndex        =   3
      Top             =   8175
      Width           =   900
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3735
      Top             =   7995
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PaddleMax Levels (*.lev)|*.lev"
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   8175
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Undo"
      Enabled         =   0   'False
      Height          =   330
      Left            =   900
      TabIndex        =   1
      Top             =   8175
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8175
      Width           =   900
   End
   Begin VB.Line HGridLine 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   11250
      Y1              =   3450
      Y2              =   3450
   End
   Begin VB.Line VGridLine 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8985
   End
   Begin VB.Image Solid 
      Height          =   300
      Left            =   6765
      Picture         =   "Form1.frx":030A
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Green 
      Height          =   300
      Left            =   7515
      Picture         =   "Form1.frx":2820
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Cyan 
      Height          =   300
      Left            =   8265
      Picture         =   "Form1.frx":4F50
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Red 
      Height          =   300
      Left            =   9015
      Picture         =   "Form1.frx":7616
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Orange 
      Height          =   300
      Left            =   9765
      Picture         =   "Form1.frx":9CF8
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Yellow 
      Height          =   300
      Left            =   10515
      Picture         =   "Form1.frx":C3E3
      Top             =   8205
      Width           =   750
   End
   Begin VB.Image Brick 
      Height          =   300
      Index           =   0
      Left            =   10515
      Picture         =   "Form1.frx":EA1D
      Top             =   7905
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Brick_Click(Index As Integer)
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
SEL = True
End Sub

Private Sub Command1_Click()
Dim CC As String
CC = App.Path
If Right(CC, 1) <> "\" Then CC = CC + "\"
Dim FNAME
FNAME = InputBox("Filename?", "Enter Filenname")
If FNAME = vbCancel Then Exit Sub
Open CC + FNAME + ".LEV" For Output As #1
Dim Z As Integer
For Z = 1 To BRICKSEL
    Print #1, Brick(Z).Top
    Print #1, Brick(Z).Left
    If Brick(Z).Picture = Yellow.Picture Then Print #1, "YELLOW"
    If Brick(Z).Picture = Orange.Picture Then Print #1, "ORANGE"
    If Brick(Z).Picture = Red.Picture Then Print #1, "RED"
    If Brick(Z).Picture = Cyan.Picture Then Print #1, "CYAN"
    If Brick(Z).Picture = Green.Picture Then Print #1, "GREEN"
    If Brick(Z).Picture = Solid.Picture Then Print #1, "SOLID"
Next Z
MsgBox "Done!"
Close #1
End Sub

Private Sub Command2_Click()
If BRICKSEL = 0 Then
    Command2.Enabled = False
    Exit Sub
End If
Unload Brick(BRICKSEL)
BRICKSEL = BRICKSEL - 1
End Sub

Private Sub Command3_Click()
CD1.ShowOpen
If CD1.FileName = "" Then Exit Sub

On Error Resume Next
For Z = 1 To 1000
    Unload Brick(Z)
Next Z

Open CD1.FileName For Input As #1
Dim TEMPLINE
BRICKSEL = 1
Do Until EOF(1)
    Line Input #1, TEMPLINE
    Load Brick(BRICKSEL)
    Brick(BRICKSEL).Visible = True
    Brick(BRICKSEL).Enabled = False
    Brick(BRICKSEL).Top = Val(TEMPLINE)
    Brick(BRICKSEL).ZOrder
    Line Input #1, TEMPLINE
    Brick(BRICKSEL).Left = Val(TEMPLINE)
    Line Input #1, TEMPLINE
    If TEMPLINE = "YELLOW" Then Brick(BRICKSEL).Picture = Yellow.Picture
    If TEMPLINE = "ORANGE" Then Brick(BRICKSEL).Picture = Orange.Picture
    If TEMPLINE = "RED" Then Brick(BRICKSEL).Picture = Red.Picture
    If TEMPLINE = "CYAN" Then Brick(BRICKSEL).Picture = Cyan.Picture
    If TEMPLINE = "GREEN" Then Brick(BRICKSEL).Picture = Green.Picture
    If TEMPLINE = "SOLID" Then Brick(BRICKSEL).Picture = Solid.Picture
    BRICKSEL = BRICKSEL + 1
Loop
Close #1
BRICKSEL = BRICKSEL - 1
End Sub

Private Sub Command4_Click()
Dim RESPONSE
RESPONSE = MsgBox("Are you sure you want to create a new level?", vbYesNo, "PaddleMax Level Designer")
If RESPONSE = vbNo Then Exit Sub
On Error Resume Next
For Z = 1 To 1000
    Unload Brick(Z)
Next Z
End Sub

Private Sub Cyan_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Cyan.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub

Private Sub Form_Click()
If SEL = True Then Command2.Enabled = True
SEL = False
Form1.MousePointer = 0
End Sub

Private Sub Form_Load()
BRICKSEL = 0
SEL = False
For Z = 375 To Form1.Width Step 375
    Load VGridLine(Z / 375)
    With VGridLine(Z / 375)
    .Visible = True
    .X1 = Z
    .X2 = Z
    .Y1 = 0
    .Y2 = Form1.Height
    End With
Next Z
For Z = 150 To Form1.Height Step 150
    Load HGridLine(Z / 150)
    With HGridLine(Z / 150)
    .Visible = True
    .X1 = 0
    .X2 = Form1.Width
    .Y1 = Z
    .Y2 = Z
    End With
Next Z
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If SEL = False Then Exit Sub

Brick(BRICKSEL).Left = (Int(X / 375) * 375)
Brick(BRICKSEL).Top = (Int(Y / 150) * 150)
End Sub

Private Sub Green_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Green.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub

Private Sub Orange_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Orange.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub

Private Sub Red_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Red.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub

Private Sub Solid_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Solid.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub

Private Sub Yellow_Click()
BRICKSEL = BRICKSEL + 1
Load Brick(BRICKSEL)
Brick(BRICKSEL).Visible = True
Brick(BRICKSEL).Enabled = False
Brick(BRICKSEL).Picture = Yellow.Picture
Brick(BRICKSEL).ZOrder
SEL = True
Form1.MousePointer = 99
End Sub
