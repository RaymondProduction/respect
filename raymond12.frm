VERSION 5.00
Begin VB.Form Form12 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Створити папку"
   ClientHeight    =   1305
   ClientLeft      =   840
   ClientTop       =   1365
   ClientWidth     =   5100
   Icon            =   "raymond12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "Нова папка"
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Скасувати"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Назва папки:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   140
      Width           =   1095
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const VK_RETURN = &HD
Const VK_ESCAPE = &H1B
Private Sub Command1_Click()
On Error GoTo ErrorDir
 MkDir (Form2.Dir1.Path + "\" + Text1.Text)
ErrorDir:
 If Err.number Then
  If Err.number = 75 Then
  MsgBox "Папка існує", vbInformation + vbOKOnly, "Інформую"
  Else
  If Err.number <> 0 Then MsgBox "Помилка №" + Str(Err.number), vbCritical + vbOKOnly, "Помилка!"
  End If
  End If
  Form2.Dir1.Refresh
  Unload Me
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = VK_ESCAPE Then Unload Me
 If KeyCode = VK_RETURN Then Command1_Click
End Sub

Private Sub Command3_Click()
Text1.SetFocus
Text1.SelLength = 10
End Sub

Private Sub Form_Load()
Form2.Enabled = False
'Простой пример программного выравнивания формы по центру экрана.
 Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
 Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
 'SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
 If Text1.Enabled And Text1.Visible Then Text1.SetFocus
 Text1.SelStart = 0
 Text1.SelLength = Len(Text1.Text)
End Sub
Private Sub Text_Load()
If Text1.Enabled And Text1.Visible Then Text1.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Form2.Enabled = True
End Sub
