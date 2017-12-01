VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Помилка!!!"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3675
   Icon            =   "raymond7.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Ок"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Трапилась невідома помилка при відкритті файла. Можливо файл спотворений та невідповідає формату який обробляє программа"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form7"
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

Private Sub Form_Unload(Cancel As Integer)
 Form1.Enabled = True
End Sub

