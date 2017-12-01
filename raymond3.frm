VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Попередження!!! "
   ClientHeight    =   1440
   ClientLeft      =   3930
   ClientTop       =   4275
   ClientWidth     =   3435
   Icon            =   "raymond3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Ні"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Так"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Увага заманити існуючий файл новим?"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = &H1B Then Unload Me
End Sub

Private Sub Command1_Click()
 Form1.SaveFile (Form2.File1.Path & "\" & Form2.Combo1.Text)
 Unload Form2
 Unload Me 'Закрывает форму текущую
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Form2.Enabled = True
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
End Sub
