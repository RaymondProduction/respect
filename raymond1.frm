VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Вимірювання спектру"
   ClientHeight    =   6570
   ClientLeft      =   2010
   ClientTop       =   2520
   ClientWidth     =   9780
   Icon            =   "raymond1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9780
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame4 
      Caption         =   "COM порт"
      Height          =   615
      Left            =   120
      TabIndex        =   61
      Top             =   5760
      WhatsThisHelpID =   14
      Width           =   3135
      Begin VB.OptionButton Option9 
         Caption         =   "4"
         Height          =   255
         Left            =   2400
         TabIndex        =   65
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton Option8 
         Caption         =   "3"
         Height          =   255
         Left            =   1800
         TabIndex        =   64
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option7 
         Caption         =   "2"
         Height          =   255
         Left            =   1200
         TabIndex        =   63
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option6 
         Caption         =   "1"
         Height          =   255
         Left            =   600
         TabIndex        =   62
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   240
         Picture         =   "raymond1.frx":014A
         Top             =   240
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400000&
      Height          =   4444
      Left            =   3360
      ScaleHeight     =   480
      ScaleMode       =   0  'User
      ScaleWidth      =   640
      TabIndex        =   50
      Top             =   120
      WhatsThisHelpID =   16
      Width           =   5925
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   960
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1440
      Top             =   4560
   End
   Begin VB.Frame Frame3 
      Caption         =   "Показувати графік"
      Height          =   615
      Left            =   3480
      TabIndex        =   39
      Top             =   5760
      WhatsThisHelpID =   13
      Width           =   5895
      Begin VB.CheckBox Check5 
         Caption         =   "5"
         Height          =   255
         Left            =   4920
         TabIndex        =   49
         Top             =   240
         Width           =   555
      End
      Begin VB.CheckBox Check4 
         Caption         =   "4"
         Height          =   255
         Left            =   3840
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check3 
         Caption         =   "3"
         Height          =   255
         Left            =   2640
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check2 
         Caption         =   "2"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H0000C000&
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   44
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   43
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H000080FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   42
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H00FFFF00&
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   41
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture11 
         BackColor       =   &H00FF80FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   5520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   40
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Активний графік"
      Height          =   615
      Left            =   3480
      TabIndex        =   28
      Top             =   5040
      WhatsThisHelpID =   12
      Width           =   5895
      Begin VB.OptionButton Option5 
         Caption         =   "5"
         Height          =   255
         Left            =   5040
         TabIndex        =   38
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option4 
         Caption         =   "4"
         Height          =   255
         Left            =   3840
         TabIndex        =   37
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "3"
         Height          =   255
         Left            =   2640
         TabIndex        =   36
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2"
         Height          =   255
         Left            =   1440
         TabIndex        =   35
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H0000C000&
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   33
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   1920
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   32
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H000080FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   3120
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   31
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00FFFF00&
         Enabled         =   0   'False
         Height          =   255
         Left            =   4320
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   30
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H00FF80FF&
         Enabled         =   0   'False
         Height          =   255
         Left            =   5520
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   29
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton Command5 
      Height          =   517
      Left            =   2640
      Picture         =   "raymond1.frx":04CF
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5137
      Width           =   735
   End
   Begin VB.Frame Frame6 
      Caption         =   "Монітор"
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      WhatsThisHelpID =   11
      Width           =   2415
      Begin VB.PictureBox Picture12 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         Picture         =   "raymond1.frx":090B
         ScaleHeight     =   255
         ScaleWidth      =   135
         TabIndex        =   22
         Top             =   260
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   240
         WhatsThisHelpID =   1
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Сигнал ="
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label22 
         Caption         =   "="
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Параметри"
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   2895
      Begin VB.CheckBox Check6 
         Caption         =   "Зберігати параметри"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2280
         Value           =   1  'Checked
         WhatsThisHelpID =   10
         Width           =   2535
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1560
         TabIndex        =   19
         Text            =   "1"
         Top             =   1800
         WhatsThisHelpID =   9
         Width           =   855
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   1560
         TabIndex        =   18
         Text            =   " 850"
         Top             =   1440
         WhatsThisHelpID =   8
         Width           =   855
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Text            =   " 400"
         Top             =   1080
         WhatsThisHelpID =   7
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Text            =   "4.5"
         Top             =   720
         WhatsThisHelpID =   6
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         Text            =   "1"
         Top             =   360
         WhatsThisHelpID =   5
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "нм"
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label10 
         Caption         =   "сек"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "нм"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "нм"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Крок сканування"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         WhatsThisHelpID =   9
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Кінець"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         WhatsThisHelpID =   8
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Початок"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         WhatsThisHelpID =   7
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Макс. Знач."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         WhatsThisHelpID =   6
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Час вимірювання"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         WhatsThisHelpID =   5
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Робота з спектром "
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton Command4 
         Caption         =   "Відкрити"
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   840
         WhatsThisHelpID =   4
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Зберегти"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         WhatsThisHelpID =   3
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Зупинка"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         WhatsThisHelpID =   2
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Запуск"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   2400
      Left            =   9480
      TabIndex        =   51
      Top             =   2220
      WhatsThisHelpID =   15
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4233
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      SelStart        =   80
      TickStyle       =   3
      TickFrequency   =   5
      Value           =   80
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   360
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      HelpFile        =   "help.chm"
   End
   Begin VB.Label Label12 
      Caption         =   "400"
      Height          =   255
      Left            =   3480
      TabIndex        =   60
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label13 
      Caption         =   " 490"
      Height          =   255
      Left            =   4680
      TabIndex        =   59
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   " 580"
      Height          =   255
      Left            =   5880
      TabIndex        =   58
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label15 
      Caption         =   " 670"
      Height          =   255
      Left            =   7080
      TabIndex        =   57
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label16 
      Caption         =   " 760"
      Height          =   255
      Left            =   8160
      TabIndex        =   56
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label17 
      Caption         =   " 850"
      Height          =   255
      Left            =   9360
      TabIndex        =   55
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   54
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label20 
      Caption         =   " -4.5"
      Height          =   255
      Left            =   3000
      TabIndex        =   53
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label19 
      Caption         =   " 0"
      Height          =   255
      Left            =   3120
      TabIndex        =   52
      Top             =   2280
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub Beep Lib "speaker.dll" ()

Rem Свойство AutoRedraw
Rem В Visual Basic 6.0 графические методы могут быть вызваны из
Rem любой процедуры обработки события; свойство AutoRedraw
Rem используется для сохранения графики, когда графические
Rem методы вызываются из события, отличного от события Paint.
Rem Короче потставив свойство Picture1.AutoRedraw=True теперь
Rem окнами и сварачиванием окон не стирается изображение :-)

Public globali, stepx, zerox, ScanPos, signal, beginx, endx, maxvolt, minvolt, oldy, presy As Double
Const VK_F1 = &H70
Const VK_F2 = &H71
Const VK_F3 = &H72
Const VK_SPACE = &H20
Const VK_ESCAPE = &H1B
Const VK_RETURN = &HD
Const VK_DOWN = &H28
Const VK_UP = &H26
Private StepScan As Double
Dim ComNumber As Integer
Dim NameGraphics(1 To 5) As String
Dim MassiveKey(0 To 255) As Byte
Private stepp, endset As Integer
Dim Settings(1 To 20) As String
Dim SaveOk(1 To 5) As Boolean
Dim ColorGraphic(1 To 5) As Long
Dim GraphicEnd(1 To 5) As Integer
Public SelectGraphic As Integer
Dim NanoMetr(0 To 2000, 1 To 5) As Double, SignalD(0 To 2000, 1 To 5) As Double


' Процедуры изменения размера изображения в максимальном режиме
Dim Border, BorderFirst As Single
Dim Popal As Boolean

Sub InitSizePicture()
Picture1.AutoRedraw = True
Border = Picture1.Top + Picture1.Height
BorderFirst = Border
Popal = False
End Sub
Sub ControlSizePicture()
Form1.MousePointer = 1
Picture1.MousePointer = 1
Popal = False
Border = Picture1.Top + Picture1.Height
Picture1.ToolTipText = ""
End Sub
Sub ReSizePicture(Button As Integer, x, y As Single)
Dim R As Boolean
R = Abs(x - Picture1.Left - Picture1.Width / 2) < Picture1.Width / 2
If (Abs(y - Border) < 100) And (R) Then
  Form1.MousePointer = 7
  Picture1.MousePointer = 7
  Picture1.ToolTipText = "Ви можете підняти або опустити межу поділу між малюнком і шкалою"
  Else
  Picture1.ToolTipText = ""
  Form1.MousePointer = 1
  Picture1.MousePointer = 1
End If
If (Button = 1) And (Abs(y - Border) < 100) Then
Popal = True
End If
If Button = 1 Then
If Screen.Height * 0.93 > y And y > Screen.Height * 0.75 Then
Picture1.Height = y
Border = y
Picture1.ScaleHeight = 480
Picture1.ScaleWidth = 640
Label12.Top = Border
Label13.Top = Border
Label14.Top = Border
Label15.Top = Border
Label16.Top = Border
Label17.Top = Border
GraphicsView
End If
End If
End Sub
Sub MaxScreen()
 'Me.BorderStyle = 0
 Form1.Left = 10
 Form1.Top = 10
 ' - Round_Ray(Screen.Width * 0.1)
 Form1.Width = Screen.Width
 Form1.Height = Screen.Height
 Picture1.Left = 0
 Picture1.Top = 0
 Picture1.Width = Screen.Width
 Picture1.Height = Screen.Height - Round_Ray(Screen.Height * 0.1)
 If ("" <> ReadSetting("Межа")) And (Val(ReadSetting("Межа")) < Screen.Height) Then
    Picture1.Height = Val(ReadSetting("Межа"))
 End If
 Picture1.ScaleHeight = 480
 Picture1.ScaleWidth = 640
 Frame1.Visible = False
 Frame2.Visible = False
 Frame3.Visible = False
 Frame4.Visible = False
 Frame5.Visible = False
 Frame6.Visible = False
 Command5.Visible = False
 Slider1.Visible = False
 BorderFirst = Picture1.Top + Picture1.Height
 Label12.Top = BorderFirst
 Label13.Top = BorderFirst
 Label14.Top = BorderFirst
 Label15.Top = BorderFirst
 Label16.Top = BorderFirst
 Label17.Top = BorderFirst
 Label12.Left = 0
 Label13.Left = Round_Ray(Screen.Width * 0.2)
 Label14.Left = Round_Ray(Screen.Width * 0.4)
 Label15.Left = Round_Ray(Screen.Width * 0.6)
 Label16.Left = Round_Ray(Screen.Width * 0.8)
 Label17.Left = Round_Ray(Screen.Width * 0.95)
 GraphicsView
  ' Для изменения размера изображения в максимальном режиме
InitSizePicture
 '---------------------------
End Sub
Sub MinScreen()
 WriteSetting ("Межа=" + Str(Picture1.Height))
 'Me.BorderStyle = 3
 Picture1.Left = 3480
 Picture1.Top = 120
 Picture1.Width = 5925
 Picture1.Height = 4444
 Picture1.ScaleHeight = 480
 Picture1.ScaleWidth = 640
 GraphicsView
 Form1.Width = 9870
 Form1.Height = 7050
 Form1.Left = Round_Ray((Screen.Width - Form1.Width) / 2)
 Form1.Top = Round_Ray((Screen.Height - Form1.Height) / 2)
 Frame1.Visible = True
 Frame2.Visible = True
 Frame3.Visible = True
 Frame4.Visible = True
 Frame5.Visible = True
 Frame6.Visible = True
 Command5.Visible = True
 Slider1.Visible = True
 Label12.Top = 4680
 Label13.Top = 4680
 Label14.Top = 4680
 Label15.Top = 4680
 Label16.Top = 4680
 Label17.Top = 4680
 Label12.Left = 3480
 Label13.Left = 4680
 Label14.Left = 5880
 Label15.Left = 7080
 Label16.Left = 8160
 Label17.Left = 9360
End Sub
Sub MaxMinScreen()
 If Picture1.WhatsThisHelpID = 16 Then
    MaxScreen
    Picture1.WhatsThisHelpID = 17
    Else
    MinScreen
    Picture1.WhatsThisHelpID = 16
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Picture1.WhatsThisHelpID = 17 Then
 ControlSizePicture
End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Picture1.WhatsThisHelpID = 17 Then
 ReSizePicture Button, x, y
End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim kSize As Double
'Коефіент пропорциональності между координатами на картинке и форме
kSize = Picture1.Height / Picture1.ScaleHeight
If Picture1.WhatsThisHelpID = 17 Then
 ReSizePicture Button, x, Picture1.Top + kSize * y
End If
End Sub
Sub SetComPort()
MSComm1.Settings = "1200,N,7,2" ' установки передачи данных
On Error GoTo ErrorComPort
If Option6.Value Then MSComm1.CommPort = 1 ' номер COM порта
If Option7.Value Then MSComm1.CommPort = 2
If Option8.Value Then MSComm1.CommPort = 3
If Option9.Value Then MSComm1.CommPort = 4
'8002 Неверный номер порта
'8005 Порт уже открыт
'8012 Устройство не является открытым
'8018 Операции только при открытых портах
If Not MSComm1.PortOpen Then MSComm1.PortOpen = True ' открываем ком порт
ErrorComPort:
Dim MsgInfo As String
If Err.number <> 0 Then
    MsgInfo = "Невідома помилка. Пов'язана з Com-портом №" + Str(Err.number)
    Select Case Err.number
        Case 8002: MsgInfo = "Невірний номер порту"
        Case 8012: MsgInfo = "Пристрій не підключено, або їм користується інша програма"
        Case 8005: MsgInfo = "Порт уже відкрито. Внутрішня помилка."
 End Select
 If Err.number = 8002 Then
    Form10.Label1.Caption = MsgInfo
    Form10.Show
    Else
    If Err.number <> 8005 Then MsgBox MsgInfo, vbCritical + vbOKOnly, "Помилка!"
    End If
 End If
End Sub

Sub CloseComPort()
 If MSComm1.PortOpen Then MSComm1.PortOpen = False
End Sub

'Функция на проверку если папка
Public Function FolderExists(ByVal strPathName As String) As Boolean
Dim DirectoryFound As String
Const errPathNotFound As Integer = 76
On Error GoTo 0
DirectoryFound = Dir(strPathName, vbDirectory)
If (Len(DirectoryFound) = 0 Or Err = errPathNotFound) Then
FolderExists = False
Else
FolderExists = True
End If
End Function

'Replace(Expression,Find,Replace[,Start[,Count[,Compare]]])
'Новая функция, которая появилась в Visual Basic 6.0.
'Возвращаемое значение:
'В результате действия функции Replace возвращается исходная строка с замененным строковым фрагментом.
'Параметры:
'Expression - Обязательный аргумент - строка, в которой требуется замена
'Find - Обязательный аргумент - подстрока, которую нужно заменить
'Replace - Обязательный аргумент - подстрока замены
'Start - Необязательный аргумент - указывает позицию
'Count - Необязательный аргумент - указывает число
'Compare - Необязательный аргумент - вид сравнения
Sub Autosave()
    If Not FolderExists("AutoSave") Then MkDir ("AutoSave")
    SaveFile (App.Path + "\AutoSave\" + Date$ + "-" + Replace(Time$, ":", "-") + ".dat")
    SaveOk(SelectGraphic) = False
    WriteSetting ("Збережені дані?=ні")
    SaveSettings
End Sub

Sub HelpShow()
 'Dim RetVal
' RetVal = Shell("hh help.chm", 1) ' Help.
CommonDialog1.HelpFile = App.Path + "\ReSpect.hlp"
'CommonDialog1.HelpCommand = cdlHelpContext
CommonDialog1.HelpCommand = &HB
'Выводит окно Help Topics и активирует вкладку выбранную в прошлый раз
CommonDialog1.ShowHelp
End Sub

Private Sub Command5_Click()
 HelpShow
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case VK_SPACE
      If Command1.Enabled = True Then Command1_Click
    Case VK_ESCAPE
      If (Command3.Enabled = False) And (Picture1.WhatsThisHelpID = 16) Then
         Unload Me
         Else
         If Picture1.WhatsThisHelpID = 17 Then MaxMinScreen
         If Command3.Enabled Then Command3_Click
         End If
    Case VK_RETURN: MaxMinScreen
    Case VK_F1: HelpShow
    Case VK_F2: Command2_Click
    Case VK_F3: Command4_Click
    Case VK_UP: Slider1.Value = Slider1.Value - 1
    Case VK_DOWN: Slider1.Value = Slider1.Value + 1
 End Select
End Sub


Sub AllEnabled(en As Boolean)
 Command1.Enabled = en
 Command2.Enabled = en
 Command4.Enabled = en
 Combo1.Enabled = en
 Combo2.Enabled = en
 Combo3.Enabled = en
 Combo4.Enabled = en
 Combo5.Enabled = en
 Check1.Enabled = en
 Check2.Enabled = en
 Check3.Enabled = en
 Check4.Enabled = en
 Check5.Enabled = en
 Check6.Enabled = en
 Option1.Enabled = en
 Option2.Enabled = en
 Option3.Enabled = en
 Option4.Enabled = en
 Option5.Enabled = en
 Option6.Enabled = en
 Option7.Enabled = en
 Option8.Enabled = en
 Option9.Enabled = en
 Slider1.Enabled = en
End Sub
Sub LoadSettings()
endset = 0
Dim j As Integer
Dim MyFile 'Объявляем переменную для свободного файла
MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
On Error GoTo FileCheck
Open ("ReSpect.ini") For Input As #MyFile 'Открываем файл  для чтения
j = 0
Do Until EOF(MyFile)
   'при каждом вызове оператора Line Input он записывает в
   'переменную новою строку
 j = j + 1
 Line Input #MyFile, Settings(j)
Loop
endset = j
Close #MyFile
FileCheck:
 Dim ErrorNumber As Integer
 ErrorNumber = Err.number
 If ErrorNumber <> 0 Then endset = 0
End Sub

Sub SaveSettings()
 Dim stFile As String
 Dim j As Integer
 Dim MyFile 'Объявляем переменную для свободного файла
 MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
Open ("ReSpect.ini") For Output As #MyFile 'Открываем файл  для записи
For j = 1 To endset
Print #MyFile, Settings(j)
Next j
Close #MyFile 'Закрываем файл
End Sub

Function Round_Ray(x1 As Double) As Long
 If x1 - Int(x1) >= 0.5 Then
  Round_Ray = Int(x1) + 1
 Else
  Round_Ray = Int(x1)
 End If
End Function
Function test() As Double
 Dim y As Double
 On Error GoTo ErrorComPort
 MSComm1.Output = "D"
 y = Val(Mid$(MSComm1.Input, 3, 7))
 test = y
ErrorComPort:
 Dim MsgInfo As String
If Err.number <> 0 Then
    MsgInfo = "Невідома помилка. Пов'язана з Com-портом #" + Str(Err.number)
    Select Case Err.number
        Case 8002: MsgInfo = "Невірний номер порту"
        Case 8018: MsgInfo = "Зчитування сигналу можливо тільки при відкритому"
                   MsgInfo = MsgInfo + " Com-порту. Можливо вольтметр не підключений"
                   MsgInfo = MsgInfo + " або зазначений порт не вірний"
        Case 8012: MsgInfo = "Пристрій не підключено, або їм користується інша програма"
        Case 8005: MsgInfo = "Порт уже відкрито. Внутрішня помилка."
 End Select
 MsgBox MsgInfo, vbCritical + vbOKOnly, "Помилка!"
 End If
End Function
Function test1() As Double
 Dim y As Integer
 globali = globali + 0.1
 y = Sin(globali) * 4.5
 test1 = y
End Function
Sub Beep_Dll()
 On Error GoTo ErrorDll
 Beep
ErrorDll:
If Err.number <> 0 Then
  If (Err.number = 53) Or (Err.number = 48) Then
  Form11.Show
  Else
  MsgBox "Помилка №" + Str(Err.number), vbOKOnly + vbCritical, "Помилка!"
 End If
 End If
 MsgBox "Запис спектру закінчений", vbOKOnly + vbInformation, "Повідомлення"
End Sub



Sub ReperPointSignal()
 Dim tempvolt As Double
 Label18.Caption = Combo2.Text
 Label19.Caption = " 0"
 maxvolt = Val(Combo2.Text)
 minvolt = (-1) * (480 - zerox) * maxvolt / zerox
 Label20.Caption = Str(minvolt)
 GraphicsView
End Sub
Sub ReperPoint()
 Dim b, k As Double
 Label12.Caption = Combo3.Text
 b = Val(Combo3.Text)
 k = (Val(Combo4.Text) - b) / 5
 Label13.Caption = Left(Str(b + k), 4)
 Label14.Caption = Left(Str(b + k * 2), 4)
 Label15.Caption = Left(Str(b + k * 3), 4)
 Label16.Caption = Left(Str(b + k * 4), 4)
 Label17.Caption = Left(Combo4.Text, 4)
 GraphicsView
End Sub

Sub GraphicView(selectview As Integer)
 Dim stepp1, xcor1, xcor2, oldy1, presy1 As Integer
 Dim maxnano As Double
 Dim saveval As String
 saveval = Combo4.Text
 maxnano = Val(Combo4.Text)
 Dim minnano As Double
 minnano = Val(Combo3.Text)
On Error GoTo Errorzero
 For stepp1 = 1 To GraphicEnd(selectview)
 oldy1 = Round_Ray((SignalD(stepp1 - 1, selectview) / maxvolt) * zerox) * (-1) + zerox
 presy1 = Round_Ray((SignalD(stepp1, selectview) / maxvolt) * zerox) * (-1) + zerox
 'Picture1.Line (Round_Ray(stepp1 * stepx) - stepx, oldy1)-(Round_Ray(stepp1 * stepx), presy1), ColorGraphic(selectview)
 xcor1 = Round_Ray(((NanoMetr(stepp1 - 1, selectview) - minnano) / (maxnano - minnano)) * 640)
 xcor2 = Round_Ray(((NanoMetr(stepp1, selectview) - minnano) / (maxnano - minnano)) * 640)
 Picture1.Line (xcor1, oldy1)-(xcor2, presy1), ColorGraphic(selectview)
 Next stepp1
Errorzero:
 If Err.number = 6 Then Combo4.Text = saveval
 
End Sub

Sub GraphicsView()
 GraphicInstall
 If Check1.Value Then GraphicView (1)
 If Check2.Value Then GraphicView (2)
 If Check3.Value Then GraphicView (3)
 If Check4.Value Then GraphicView (4)
 If Check5.Value Then GraphicView (5)
 GraphicView (SelectGraphic)
 Form1.Caption = "Record of Spectrum Ver. 1.2.0.5 Вимірювання спектру " + NameGraphics(SelectGraphic)
End Sub
Sub GraphicInstall()
 Picture1.Cls
 Dim j As Integer
 ColorGraphic(1) = &HC000&
 ColorGraphic(2) = &HFF&
 ColorGraphic(3) = &H80FF&
 ColorGraphic(4) = &HFFFF00
 ColorGraphic(5) = &HFF80FF
 Picture1.DrawStyle = 2
 For j = 1 To 4
  Picture1.Line (j * 128, 0)-(j * 128, 480), &HBF0000
  Next j
  Picture1.Line (0, zerox)-(640, zerox), &HBF0000
 Picture1.DrawStyle = 0
 Label20.Caption = Str(minvolt)
 Slider1.Value = Round_Ray((zerox - 240) * Slider1.max / 240)
 If Slider1.Value <= 80 Then
    Label19.Top = Round_Ray(4400 * zerox / 480 + 40)
 Else
    Label19.Top = 4092
End If
 If Slider1.Value = Slider1.max Then
                        Label19.Caption = ""
                       Else
                       Label19.Caption = " 0"
                       End If
End Sub

Private Sub Check1_Click()
 GraphicsView
End Sub
Private Sub Check2_Click()
 GraphicsView
End Sub
Private Sub Check3_Click()
 GraphicsView
End Sub
Private Sub Check4_Click()
 GraphicsView
End Sub
Private Sub Check5_Click()
 GraphicsView
End Sub
Private Sub Combo2_Change()
  If InStr(Combo2.Text, ".") > 0 Then
   Combo2.Text = Left(Combo2.Text, InStr(Combo2.Text, ".") + 3)
  End If
 ReperPointSignal
End Sub
Private Sub Combo2_Click()
 ReperPointSignal
End Sub
Sub changecombo()
If Val(Combo5.Text) <> 0 And Abs(Val(Combo3.Text) - Val(Combo4.Text)) <> 0 Then
    Form1.ReperPoint
    stepx = 640 / Abs((Val(Combo3.Text) - Val(Combo4.Text)) / Val(Combo5.Text))
    Label2.Caption = Combo3.Text
    GraphicsView
    End If
End Sub
Private Sub Combo3_Change()
changecombo
End Sub
Private Sub Combo4_Change()
changecombo
End Sub
Private Sub Combo3_Click()
changecombo
End Sub
Private Sub Combo4_Click()
changecombo
End Sub


Private Sub Combo5_Change()
 If GraphicEnd(SelectGraphic) <> 0 Then
  Combo5.Text = Str(Abs(NanoMetr(0, SelectGraphic) - NanoMetr(1, SelectGraphic)))
  GraphicsView
 End If
 End Sub
 
 Private Sub Combo5_Click()
 If GraphicEnd(SelectGraphic) <> 0 Then
 Combo5.Text = Str(Abs(NanoMetr(0, SelectGraphic) - NanoMetr(1, SelectGraphic)))
 GraphicsView
 End If
End Sub
Private Sub Command2_Click()
If GraphicEnd(SelectGraphic) <> 0 Then
Form2.IOparam = "Зберегти данні"
Form2.Command1.Caption = "Зберегти"
Form1.Enabled = False
Form2.Show
End If
End Sub

Private Sub Command3_Click()
 Combo4.Text = Str(ScanPos)
 stepx = 640 / Abs((Val(Combo3.Text) - Val(Combo4.Text)) / Val(Combo5.Text))
 Timer1.Enabled = False
 GraphicsView
 Command3.Enabled = False
 AllEnabled (True)
End Sub

Private Sub Command4_Click()
Form2.GrEnd = GraphicEnd(SelectGraphic)
Form2.IOparam = "Відкрити данні"
Form2.Command1.Caption = "Відкрити"
Form1.Enabled = False
Form2.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 WriteSettings 'Прочитать настройки с формы
 WriteSetting ("Збережені дані?=так")
 CloseComPort
 SaveSettings 'Записать настройки
 Cancel = True
  Form5.Label2.Caption = ""
   Dim i As Integer
 For i = 1 To 5
    If Not (SaveOk(i)) Then
     If Form5.Label2.Caption <> "" Then Form5.Label2.Caption = Form5.Label2.Caption + ","
     Form5.Label2.Caption = Form5.Label2.Caption + "  №" + Str(i)
     End If
 Next i
 If Form5.Label2.Caption = "" Then
                Cancel = False
                End
                Else
                Form5.Show
                End If
End Sub


Private Sub Label18_Change()
  Label18.Left = 3300 - Round_Ray(90 * Len(Label18.Caption))
  Label18.Width = Round_Ray(135 * Len(Label18.Caption))
  If InStr(Label18.Caption, ".") > 0 Then
   Label18.Caption = Left(Label18.Caption, InStr(Label18.Caption, ".") + 3)
  End If
End Sub
Private Sub Label20_Change()
  Label20.Left = 3300 - Round_Ray(90 * Len(Label20.Caption))
  Label20.Width = Round_Ray(135 * Len(Label20.Caption))
  If InStr(Label20.Caption, ".") > 0 Then
   Label20.Caption = Left(Label20.Caption, InStr(Label20.Caption, ".") + 3)
  End If
End Sub

Private Sub Picture1_DblClick()
 MaxMinScreen
End Sub

Private Sub Slider1_Change()
zerox = Round_Ray(240 * (Slider1.Value) / Slider1.max) + 240
minvolt = (-1) * (480 - zerox) * maxvolt / zerox
GraphicsView
End Sub


Sub WriteSetting(nameset As String)
 Dim j As Integer
 Dim naydeno As Boolean
 naydeno = False
 If endset <> 0 Then
 For j = 1 To endset
  If InStr(Left(Settings(j), InStr(Settings(j), "=") - 1), Left(nameset, InStr(nameset, "=") - 1)) <> 0 Then
     Settings(j) = nameset
     naydeno = True
     End If
 Next j
 End If
 If Not (naydeno) Then
    endset = endset + 1
    Settings(endset) = nameset
    End If
End Sub

Function ReadSetting(nameset As String) As String
Dim j As Integer
Dim resultset As String
For j = 1 To endset
If nameset = Left(Settings(j), InStr(Settings(j), "=") - 1) Then
resultset = Mid(Settings(j), InStr(Settings(j), "=") + 1, Len(Settings(j)) - InStr(Settings(j), "="))
End If
 Next j
 ReadSetting = resultset
End Function
Sub ReadSettings()
If endset <> 0 Then
 If ReadSetting("Зберігати параметри?") = "ні" Then Check6.Value = 0
 
 If Check6.Value = 1 Then
    Combo1.Text = ReadSetting("Час вимірювання")
    Combo2.Text = ReadSetting("Максимальне значення")
    Combo3.Text = ReadSetting("Початок спектру")
    Combo4.Text = ReadSetting("Кінець спектру")
    Combo5.Text = ReadSetting("Крок сканування спектру")
   End If
  End If
    ComNumber = Val(ReadSetting("Номер Com-порта"))
    Select Case ComNumber
        Case 1: Option6.Value = True
        Case 2: Option7.Value = True
        Case 3: Option8.Value = True
        Case 4: Option9.Value = True
    End Select
End Sub
Sub WriteSettings()
    Dim setsaveok As String
    If Check6.Value = 0 Then
                        setsaveok = "ні"
                        Else
                        setsaveok = "так"
    End If
    WriteSetting ("Зберігати параметри?=" + setsaveok)
If Check6.Value = 1 Then
    WriteSetting ("Час вимірювання=" + Combo1.Text)
    WriteSetting ("Максимальне значення=" + Combo2.Text)
    WriteSetting ("Початок спектру=" + Combo3.Text)
    WriteSetting ("Кінець спектру=" + Combo4.Text)
    WriteSetting ("Крок сканування спектру=" + Combo5.Text)
End If
 WriteSetting ("Номер Com-порта=" + Str(ComNumber))
End Sub
Function numberok(stnum As String) As Boolean
 numberok = ((Val(stnum) <> 0) Or stnum = "0") Or ((stnum = "-") Or (stnum = "."))
 numberok = (numberok) Or (stnum = ",")
End Function

Function transformcorrect(sttrans As String) As String
Dim number As Integer
Dim numberspace As Integer
numberspace = 0
For number = 1 To Len(sttrans)
 If numberok(Mid(sttrans, number, 1)) Then
    numberspace = numberspace + 1
    End If
Next number
If numberspace = 0 Then
sttrans = ""
End If

If Len(sttrans) > 0 Then
 sttrans = Replace(sttrans, Chr(9), " ")
 Do Until numberok(Mid(sttrans, 1, 1))
  sttrans = Right(sttrans, Len(sttrans) - 1)
  Loop
 Do Until numberok(Mid(sttrans, 1, 1))
  sttrans = Left(sttrans, Len(sttrans) - 1)
  Loop
 number = 1
 numberspace = 0
 Do Until number = Len(sttrans)
  If Mid(sttrans, number, 1) = " " Then
   numberspace = number
  End If
  If Not numberok(Mid(sttrans, number, 1)) Then
   sttrans = Left(sttrans, number - 1) + Right(sttrans, Len(sttrans) - number)
   number = number - 1
  End If
  number = number + 1
 Loop
  If numberspace > 0 Then
 sttrans = Left(sttrans, numberspace - 1) + " " + Right(sttrans, Len(sttrans) - numberspace + 1)
 Else
  sttrans = "ERRor"
 End If
 Else
  sttrans = "Empty"
  End If
transformcorrect = sttrans
End Function
Private Sub Form_Load()
 'Простой пример программного выравнивания формы по центру экрана.
 Me.Left = Round_Ray((Screen.Width - Me.Width) / 2)
 Me.Top = Round_Ray((Screen.Height - Me.Height) / 2)

 Label18_Change
 Label20_Change
 LoadSettings 'Загрузить параметри
 If "ні" = ReadSetting("Збережені дані?") Then Form9.Show
 
 App.HelpFile = App.Path + "\ReSpect.hlp" ' Файл для системи What's This

 zerox = Round_Ray(240 * (Slider1.Value) / Slider1.max) + 240
 Label19.Top = 4400 * zerox / 480 + 40
 SelectGraphic = 1
  
  
 ReperPoint
 Dim i As Integer
 For i = 1 To 10
    Combo1.AddItem Str$(i)
    Combo1.ItemData(Combo1.NewIndex) = i
 Next i
  For i = 1 To 4
    Combo5.AddItem Str$(i * 0.5)
    Combo5.ItemData(Combo5.NewIndex) = i
 Next i
 For i = 8 To 17
    Combo3.AddItem Str$(i * 50)
    Combo3.ItemData(Combo3.NewIndex) = i * 50
    Combo4.AddItem Str$(i * 50)
    Combo4.ItemData(Combo4.NewIndex) = i * 50
 Next i
 For i = 1 To 5
    SaveOk(i) = True
 Next i
 maxvolt = 4.5
 minvolt = (-1) * (480 - zerox) * maxvolt / zerox
 Label20.Caption = Str(minvolt)
 Timer1.Enabled = False
 Combo2.AddItem " 4.5"
 Combo2.ItemData(Combo2.NewIndex) = 1
 Combo2.AddItem " 10"
 Combo2.ItemData(Combo2.NewIndex) = 2
 ReadSettings 'прочиатать все настройки
 Label2.Caption = Combo3.Text
 SetComPort 'Открить ком порт
End Sub

Public Sub StartScan()
    Command3.Enabled = True
    GraphicEnd(SelectGraphic) = 0
    SaveOk(SelectGraphic) = False
    GraphicsView
    Timer1.Enabled = True
End Sub
Private Sub Command1_Click()
If Val(Combo5.Text) <> 0 And Abs(Val(Combo3.Text) - Val(Combo4.Text)) <> 0 Then
 AllEnabled (False)
 Timer1.Interval = 1000 * Val(Combo1.Text)
 beginx = Val(Combo3.Text)
 endx = Val(Combo4.Text)
 ScanPos = beginx
 StepScan = Val(Combo5.Text) * Sgn(endx - beginx)
 signal = test()
 stepp = 0
 If maxvolt <> 0 Then oldy = Round_Ray((signal / maxvolt) * zerox) * (-1) + zerox
 stepx = 640 / Abs((beginx - endx) / StepScan)
 If Val(Combo2.Text) <= 0 Then
        MsgBox "Програма не може записувати спектр якщо Мак. Знач. менше нуля або нуль", vbCritical + vbOKOnly, "Помилка!"
        AllEnabled (True)
        Else
 If GraphicEnd(SelectGraphic) <> 0 Then
                Form6.Show
                Else
   StartScan
 End If
 End If
 End If
End Sub
Private Sub Option1_Click()
 SelectGraphic = 1
 GraphicsView
End Sub
Private Sub Option2_Click()
 SelectGraphic = 2
 GraphicsView
End Sub
Private Sub Option3_Click()
 SelectGraphic = 3
 GraphicsView
End Sub
Private Sub Option4_Click()
 SelectGraphic = 4
 GraphicsView
End Sub
Private Sub Option5_Click()
 SelectGraphic = 5
 GraphicsView
End Sub
Private Sub Option6_Click()
 ComNumber = 1
 CloseComPort
 SetComPort
End Sub
Private Sub Option7_Click()
 ComNumber = 2
 CloseComPort
 SetComPort
End Sub
Private Sub Option8_Click()
 ComNumber = 3
 CloseComPort
 SetComPort
End Sub
Private Sub Option9_Click()
 ComNumber = 4
 CloseComPort
 SetComPort
End Sub
Function CorrectStringFile(incorrect As String) As String
Dim j, j1 As Integer
Dim tempcorrect(1) As String
j1 = 0
For j = 1 To Len(incorrect)
 If Mid(incorrect, j, 1) <> " " Then
  tempcorrect(j1) = tempcorrect(j1) & Mid(incorrect, j, 1)
  Else
  If j > 1 Then
   If Mid(incorrect, j - 1, 1) <> " " Then
     j1 = j1 + 1
   End If
   End If
  End If
 Next j
 CorrectStringFile = tempcorrect(0) & " " & tempcorrect(1)
End Function

Private Sub Timer1_Timer()
 NanoMetr(stepp, SelectGraphic) = ScanPos
 SignalD(stepp, SelectGraphic) = signal
 signal = test()
 'корректировка маштаба если сигнал оченъ большой
 If signal > maxvolt Or signal < -maxvolt Then
  maxvolt = Abs(signal) + 0.1 * maxvolt
  Combo2.Text = Str(maxvolt)
  oldy = Round_Ray((SignalD(stepp, SelectGraphic) / maxvolt) * zerox) * (-1) + zerox
  GraphicsView
 End If
 'корректировка если сигнал маленький
 If signal < minvolt And signal >= -maxvolt Then
  minvolt = minvolt - Abs(signal - minvolt)
  If zerox >= 240 Then
   zerox = Round_Ray(480 * maxvolt / (maxvolt - minvolt)) - 2
  End If
  oldy = Round_Ray((SignalD(stepp, SelectGraphic) / maxvolt) * zerox) * (-1) + zerox
  GraphicsView
  End If
 
 stepp = stepp + 1
 presy = Round_Ray((signal / maxvolt) * zerox) * (-1) + zerox
 Picture1.Line (Round_Ray(stepp * stepx) - stepx, oldy)-(Round_Ray(stepp * stepx), presy), ColorGraphic(SelectGraphic)
 oldy = presy

 ScanPos = ScanPos + StepScan
 GraphicEnd(SelectGraphic) = stepp - 1 'отимизировать
 If beginx < endx Then
   If ScanPos > endx Then
      Timer1.Enabled = False
      'Beep 745, 1000
      Beep_Dll
      Rem Timer2.Enabled = True
      Command3.Enabled = False
      ScanPos = endx
      AllEnabled (True)
      Autosave 'авто сохранение
   End If
  Else
   If ScanPos < endx Then
      Timer1.Enabled = False
      'Beep 745, 1000
      Beep_Dll
      Command3.Enabled = False
      ScanPos = endx
      AllEnabled (True)
      Autosave 'авто сохранение
      Rem Timer2.Enabled = True
    End If
  End If
  Label1.Caption = Left(Str(signal), 5)
  Label2.Caption = Str(ScanPos)
End Sub
Sub SaveFile(NameF As String)
 Dim stFile As String
 Dim j As Integer
 Dim MyFile 'Объявляем переменную для свободного файла
 MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами
 On Error GoTo ErrorFile 'Будем проверять на ошибки ввода вывода
Open (NameF) For Output As #MyFile 'Открываем файл  для записи
For j = 0 To GraphicEnd(SelectGraphic)
stFile = Str(NanoMetr(j, SelectGraphic)) & " " & (SignalD(j, SelectGraphic))
Print #MyFile, CorrectStringFile(stFile)
Next j
SaveOk(SelectGraphic) = True
Close #MyFile 'Закрываем файл
ErrorFile:
 If Err.number <> 0 Then Form8.Show
End Sub
Sub OpenFile(NameF As String)
 Dim j As Integer
 Dim NumberAtt As Integer
 Dim stFile As String
 Dim MyFile 'Объявляем переменную для свободного файла
 MyFile = FreeFile ' Присваиваем свободный канал, для работы с файлами

 If InStr(NameF, "\") = 0 Then
  NameGraphics(SelectGraphic) = NameF
    Else
     NameGraphics(SelectGraphic) = Right(NameF, Len(NameF) - InStrRev(NameF, "\"))
     End If
Form1.Caption = "Record of Spectrum Ver. 1.2.0.5 Вимірювання спектру " + NameGraphics(SelectGraphic)
     
 
On Error GoTo ErrorFile 'Будем проверять на ошибки ввода вывода
Open (NameF) For Input As #MyFile 'Открываем файл  для чтения
'For j = 0 To stepp - 1
'Print #MyFile, Str(NanoMetr(j, SelectGraphic)), " ", Str(SignalD(j, SelectGraphic))
'Next j
j = -1
NumberAtt = 0
Do Until EOF(MyFile)
   'при каждом вызове оператора Line Input он записывает в
   'переменную новою строку
Attempt:
 Line Input #MyFile, stFile
 stFile = transformcorrect(stFile) 'исправляем
 
If stFile = "Empty" Then
 NumberAtt = NumberAtt + 1
 If NumberAtt > 5 Then
  MsgBox "Формат не зрозумілий програмі!", vbOKOnly + vbExclamation, "Повідомлення"
  GoTo ErrorFile
  End If
 GoTo Attempt
 End If
If stFile = "ERRor" Then
MsgBox "Файл спотворений!", vbOKOnly + vbExclamation, "Повідомлення"
 GoTo ErrorFile
End If
j = j + 1
  
 NanoMetr(j, SelectGraphic) = Val(Left(stFile, InStr(stFile, " ") - 1))
 SignalD(j, SelectGraphic) = Val(Mid(stFile, InStr(stFile, " ") + 1, Len(stFile) - InStr(stFile, " ")))
 
 signal = SignalD(j, SelectGraphic)

'корректировка маштаба если сигнал оченъ большой
 If signal > Abs(maxvolt) Then
  maxvolt = Abs(signal) + 0.1 * maxvolt
  Combo2.Text = Str(maxvolt)
 End If
 'корректировка если сигнал маленький
 If signal < minvolt And signal >= -maxvolt Then
  minvolt = minvolt - Abs(signal - minvolt)
  If zerox >= 240 Then
   zerox = Round_Ray(480 * maxvolt / (maxvolt - minvolt)) - 2
  End If
  End If


Loop
GraphicEnd(SelectGraphic) = j
'поменять размери учитивая загружонные графикі
 Dim jatt, max, min As Integer
 max = 0
 min = NanoMetr(0, SelectGraphic)
 For jatt = 1 To 5
  If GraphicEnd(jatt) > 0 Then
    If NanoMetr(GraphicEnd(jatt), jatt) > max Then
         max = NanoMetr(GraphicEnd(jatt), jatt)
       End If
    If NanoMetr(0, jatt) < min Then
         min = NanoMetr(0, jatt)
       End If
    End If
Next jatt
 Combo4.Text = Str(max)
 Combo3.Text = Str(min)
 
 'stepx глабальная переменная для GraphicsView которая входит в ReperPoint
 stepx = 640 / Abs((Val(Combo3.Text) - Val(Combo4.Text)) / Val(Combo5.Text))
 Combo5.Text = Str(Abs(NanoMetr(0, SelectGraphic) - NanoMetr(1, SelectGraphic)))
 ReperPoint
 SaveOk(SelectGraphic) = True
Close #MyFile 'Закрываем файл
ErrorFile:
 If Err.number <> 0 Then
 Form7.Caption = "Номер помилки" + Str(Err.number)
 Form7.Show
 End If
 
End Sub
Private Sub Timer2_Timer()
Beep
Timer2.Enabled = False
End Sub
