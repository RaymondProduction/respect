VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Попередження!!! "
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   Icon            =   "raymond6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Ні"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Так"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Увага заманити графік новим?"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B

Private Sub Command1_Click()
     Form1.StartScan
     Unload Me
End Sub

Private Sub Command2_Click()
 Form1.AllEnabled (True)
 Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = &H1B Then Unload Me
End Sub


Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
Form1.Enabled = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Form1.Enabled = True
End Sub
