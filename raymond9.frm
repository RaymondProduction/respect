VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Повідомлення"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4455
   Icon            =   "raymond9.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   $"raymond9.frx":014A
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_ESCAPE = &H1B

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = VK_ESCAPE Then Unload Me
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
 Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
 Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
 SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
 Form1.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Form1.Enabled = True
 Unload Me
End Sub
