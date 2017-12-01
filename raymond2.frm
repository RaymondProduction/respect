VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   3435
   ClientTop       =   1965
   ClientWidth     =   4950
   FontTransparent =   0   'False
   Icon            =   "raymond2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check1 
      Caption         =   "Зберегти шлях"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   4320
      Picture         =   "raymond2.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Кнопка для створення папки "
      Top             =   3480
      Width           =   495
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   2400
      Pattern         =   "*.dat"
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "*.dat"
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public NameFile As String
Public IOparam As String
Public GrEnd As Integer
Public DirName As String
Const VK_ESCAPE = &H1B
Const VK_TAB = &H9


Rem Процедура для проверки если файл?
Function FileExist(filename As String) As Boolean
    On Error Resume Next
    FileExist = Dir$(filename) <> ""
    If Err.number <> 0 Then FileExist = False
    On Error GoTo 0
End Function
'Функция InStr служит для поиска
' номера символа (или номера байта
'для InStrB), с которого начинается
'в заданной строке образец поиска.
'Поиск идет от указанной позиции слева
'направо. Поиск вхождения одной строки
'в другую весьма часто используемая
'операция. Нумерация символов всегда
'начинается с единицы. Для поиска '
'вхождения с конца строки используйте
'функцию InStrRev.

Private Sub Command2_Click()
 Form12.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = VK_ESCAPE Then Unload Me
End Sub

Private Sub Combo1_Change()
  If Combo1.Text <> File1.filename Then File1.Pattern = Combo1.Text
 
 
End Sub
Private Sub Combo1_Click()
 File1.Pattern = Combo1.Text
End Sub
Private Sub Combo1_KeyPress(keyascii As Integer) 'При нажатии Enter
 If keyascii = 13 Then
    If InStr(Combo1.Text, "*") = 0 Then
      If IOparam <> "Відкрити данні" Then
       If FileExist(File1.Path & "\" & Combo1.Text) Then
           Form3.Show
           Else
           Form1.SaveFile (File1.Path & "\" & Combo1.Text)
           Unload Me
          End If
          Else
      If GrEnd = 0 Then
  Form1.OpenFile (File1.Path & "\" & Combo1.Text)
  Unload Me
  Else
   Form4.Show
   End If
          
          End If
          Else
          File1.Pattern = Combo1.Text
          End If
          End If
End Sub
                                        
                                        
                                     
Private Sub Command1_Click()
 If IOparam <> "Відкрити данні" Then
       If FileExist(File1.Path & "\" & Combo1.Text) Then
           Form3.Show
           Else
           Form1.SaveFile (File1.Path & "\" & Combo1.Text)
           Unload Me
          End If
        Else
      If GrEnd = 0 Then
  Form1.OpenFile (File1.Path & "\" & Combo1.Text)
  Unload Me
  Else
   Form4.Show
   End If

    End If
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form1.Enabled = True
 Dim setsaveway As String
  If Check1.Value = 1 Then
                    setsaveway = "так"
                    Else: setsaveway = "ні"
                    End If

 Form1.WriteSetting ("Зберігати шлях?=" + setsaveway)
 If Check1.Value = 1 Then
    Form1.WriteSetting ("Шлях=" + Dir1.Path)
    If InStr(File1.Pattern, "*") <> 0 Then Form1.WriteSetting ("Розширення файлів=" + File1.Pattern)
 Form1.Enabled = True
 End If
 
End Sub
Private Sub File1_KeyPress(keyascii As Integer) 'При нажатии Enter
 If keyascii = 13 Then
    If IOparam <> "Відкрити данні" Then
    Form2.Enabled = False
    Form3.Show
    Else
      If GrEnd = 0 Then
  Form1.OpenFile (File1.Path & "\" & Combo1.Text)
  Unload Me
  Else
   Form4.Show
   End If
    End If
End If
End Sub

Private Sub Form_Load()
'Простой пример программного выравнивания формы по центру экрана.
Me.Left = Form1.Round_Ray((Screen.Width - Me.Width) / 2)
Me.Top = Form1.Round_Ray((Screen.Height - Me.Height) / 2)
'SetWindowPos Me.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE 'поверх всех
If "ні" = Form1.ReadSetting("Зберігати шлях?") Then Check1.Value = 0
Me.Caption = IOparam
If Check1.Value = 0 Then
                    PathSearch = App.Path
                   Else
                    PathSearch = Form1.ReadSetting("Шлях")
                    If PathSearch = "" Then
                        PathSearch = App.Path
                    End If
                     If Form1.ReadSetting("Розширення файлів") <> "" Then
                        File1.Pattern = Form1.ReadSetting("Розширення файлів")
                        Combo1.Text = Form1.ReadSetting("Розширення файлів")
                        End If
                   End If
On Error GoTo ErrorPatch
Drive1.Drive = PathSearch
Dir1.Path = PathSearch
File1.filename = PathSearch
Combo1.AddItem "*.*"
Combo1.ItemData(Combo1.NewIndex) = 1
Combo1.AddItem "*.txt"
Combo1.ItemData(Combo1.NewIndex) = 2
Combo1.AddItem "*.dat"
Combo1.ItemData(Combo1.NewIndex) = 3
ErrorPatch:
If Err.number <> 0 Then
 If Err.number = 76 Then
    MsgBox "Шлях не знайдений." + Chr(13) + "Мабудь папка знищена, або її ім'я змінено", vbCritical + vbOKOnly, "Помилка!"
 End If
 If Err.number <> 76 Then
    MsgBox "Невідома помилка №" + Str(Err.number), vbCritical + vbOKOnly, "Помилка!"
    End If
Form1.Enabled = False
End If
End Sub
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub
Private Sub Drive1_Change()
DiskName = Left$(Drive1.Drive, 2)
Dir1.Path = DiskName & "\"
File1.Path = Dir1.Path
End Sub
Private Sub File1_Change()
Combo1.Text = File1.filename
End Sub

Private Sub File1_Click()
Combo1.Text = File1.filename
End Sub

Private Sub File1_DblClick()
Combo1.Text = File1.filename
If IOparam <> "Відкрити данні" Then
  Form2.Enabled = False
  Form3.Show
  Else
  
  If GrEnd = 0 Then
  Form1.OpenFile (File1.Path & "\" & Combo1.Text)
  Unload Me
  Else
   Form4.Show
   End If
  End If
End Sub
