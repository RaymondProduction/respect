VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Помилка!!!"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3540
   Icon            =   "raymond8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Виникла помилка при збереженні файла"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = VK_ESCAPE Then Unload Me
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
Form1.Enabled = False
End Sub
Private Sub Command1_Click()
Form1.Enabled = True
 Unload Me
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Form1.Enabled = True
 Unload Me
End Sub

