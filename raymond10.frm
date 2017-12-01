VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Помилака"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2670
   Icon            =   "raymond10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   240
      Picture         =   "raymond10.frx":014A
      ScaleHeight     =   615
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form10"
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
 'SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE ''не поверх всех
 Command1.Left = Form1.Round_Ray((Me.Width - Command1.Width) / 2)
 Form1.Enabled = False
 Beep
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Form1.Enabled = True
 Unload Me
End Sub

