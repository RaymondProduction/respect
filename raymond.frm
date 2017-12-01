VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6015
   ClientLeft      =   3240
   ClientTop       =   1980
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   8865
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Text            =   "4.5"
      Top             =   2160
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4455
      Left            =   2400
      ScaleHeight     =   507.726
      ScaleMode       =   0  'User
      ScaleWidth      =   653.368
      TabIndex        =   0
      Top             =   120
      Width           =   5925
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2640
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private globali, maxvolt As Double

Function Round_Ray(x As Double) As Integer
 If x - Int(x) >= 0.5 Then
  Round_Ray = Int(x) + 1
 Else
  Round_Ray = Int(x)
 End If
End Function
Function test() As Double
 globali = globali + 0.01
 y = Sin(globali) * 4.5
 Label1.Caption = Str(y)
 test = y
End Function
Private Sub Command2_Click()
 Label1.Caption = Combo2.Text
End Sub
Private Sub Form_Load()
 For i = 1 To 10
    Combo1.AddItem Str$(i)
    Combo1.ItemData(Combo1.NewIndex) = i
 Next i
 Combo2.AddItem "4.5"
 Combo2.ItemData(Combo2.NewIndex) = 1
 Combo2.AddItem "10"
 Combo2.ItemData(Combo2.NewIndex) = 2
End Sub
Private Sub Command1_Click()
 maxvolt = Val(Combo2.Text)
 oldy = 240
 For i = 1 To 640
 presy = Round_Ray((test() / maxvolt) * 240) + 240
 Picture1.Line (i - 1, oldy)-(i, presy)
 Rem Picture1.PSet (i, presy)
 oldy = presy
 Next i
End Sub
