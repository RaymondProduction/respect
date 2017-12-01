VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Попередження!!! "
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3000
   Icon            =   "raymond5.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Так"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ні"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Спектр"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "не збережений"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Чи варто виходить з програми?"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Не всі спектри збережені!"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = &H1B Then Unload Me
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
Form1.Enabled = False
End Sub
Private Sub Command1_Click()
Form1.Enabled = True
 Unload Me
End Sub

Private Sub Command2_Click()
 End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Form1.Enabled = True
 Unload Me
End Sub
