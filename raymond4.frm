VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Попередження!!! "
   ClientHeight    =   1275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4035
   Icon            =   "raymond4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1275
   ScaleWidth      =   4035
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
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Увага заманити  графік новим?"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = &H1B Then Unload Me
End Sub

Private Sub Command1_Click()
  Form1.OpenFile (Form2.File1.Path & "\" & Form2.Combo1.Text)
  Unload Me
End Sub

Private Sub Command2_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form2.Enabled = True
 Unload Form2
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
Form2.Enabled = False
End Sub
