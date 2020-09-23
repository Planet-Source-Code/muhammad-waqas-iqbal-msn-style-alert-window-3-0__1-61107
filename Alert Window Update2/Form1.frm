VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Alert Window - By Muhammad Waqas Iqbal"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   4260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   -120
      ScaleHeight     =   825
      ScaleWidth      =   4425
      TabIndex        =   8
      Top             =   -120
      Width           =   4455
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         Height          =   195
         Left            =   3480
         MouseIcon       =   "Form1.frx":08CA
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upgrade Your Product."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   840
         MouseIcon       =   "Form1.frx":0BD4
         TabIndex        =   9
         Top             =   360
         Width           =   1860
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form1.frx":0EDE
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Upgrade Your Product."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   210
         Left            =   865
         MouseIcon       =   "Form1.frx":17A8
         TabIndex        =   11
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   3240
      Width           =   4095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   1800
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   1815
      Left            =   -120
      ScaleHeight     =   1755
      ScaleWidth      =   4395
      TabIndex        =   2
      Top             =   720
      Width           =   4455
      Begin VB.OptionButton Option2 
         Caption         =   "Remind me later"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Whats New"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Close"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Download Now"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000010&
         Caption         =   "Label6"
         Height          =   375
         Left            =   2730
         TabIndex        =   13
         Top             =   1120
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000010&
         Caption         =   "Label6"
         Height          =   375
         Left            =   1290
         TabIndex        =   12
         Top             =   1120
         Width           =   1280
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update version to your product is available."
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   120
         Width           =   5985
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image QuestionIcon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":1AB2
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Update2Icon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":237C
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Update1Icon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":2C46
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image ErrorIcon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":3510
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image InformationIcon 
      Height          =   480
      Left            =   5040
      Picture         =   "Form1.frx":3DDA
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image ExclamationIcon 
      Height          =   480
      Left            =   5520
      Picture         =   "Form1.frx":46A4
      Top             =   2160
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Do not resize this picture box"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   2880
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' SetWindowPos() hwndInsertAfter values
Const HWND_TOP = 0
Const HWND_BOTTOM = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Dim resto As Long
Dim TaskBar As Long
Private ClsGradient As New CGradient
Private Enum TransType
    byColor
    byValue
End Enum
Private Enum IconType
    Error = 1
    Exclamation = 2
    Information = 3
    Update1 = 4
    Update2 = 5
    Question = 6
End Enum
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long




Public Function MakePopUpSize(Index As Integer)
Dim HeightInit
    Dim WindowRect As RECT
    Me.Height = Picture1.Height
    Me.Width = Picture1.Width
    HeightInit = Me.Height
    
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
    TaskBar = ((Screen.Height / Screen.TwipsPerPixelX) - WindowRect.PAbj) * Screen.TwipsPerPixelX
    
    Me.Left = Screen.Width - Me.ScaleWidth - 220
    Me.Top = Screen.Height
    resto = Me.Top - ((Me.Height * (Index)) + TaskBar)
    Me.Top = resto + Me.ScaleHeight
    Me.Show
End Function


Private Sub Command1_Click()
Dim L&
If Command1.Caption = "Download" Then
MsgBox "Your will automatically be redirected to the website.", vbInformation
Timer2.Enabled = True
For L = 255 To 0 Step -1
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
Next L
Else
Timer2.Enabled = True
For L = 255 To 0 Step -1
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
Next L
End If
End Sub

Private Sub Command2_Click()
MsgBox "MSN style alert now in windows form with " _
& "title bar, it neither requires activeX control nor " _
& "any DLL file. You can even change sound of Alert " _
& "Window, just by changing 'PlaySoundResource' in " _
& "FormLoad event form 101 to 112. You can also add " _
& "your own sound, just add wave file in the resource " _
& "file under the heading 'WAVE'." _
& vbNewLine & vbNewLine & "Gradient effect is " _
& "now included in the version of Alert Window in prior " _
& "version you could use only one color of heading area " _
& "but in this version you can use gradient effect of " _
& "two colors of your desire. You can not only change " _
& "the colors of gradient of your of desire but also " _
& "change the angle of gardient. For example: 90 or -90. " _
& "All you need to do just change color and anlge in " _
& "FormLoad Event as simple as that." _
& vbNewLine & vbNewLine & "Another upgrade " _
& "version of Alert Window is now here, in this version " _
& "of Alert Window FadeIn and FadeOut effect of form is " _
& "included. You can set FadeIn and FadeOut duration as " _
& "you like by adding 'Step' code in ForNext loop in " _
& "FormLoad Event from 1 to 255 depends on you need. More...", vbInformation, App.Title & " - Whats New"
End Sub

Private Sub Form_Initialize()
Dim x As Long
x = InitCommonControls
End Sub

Private Sub Form_Load()
On Error GoTo err:
'Please Do not open Alert Window more than 5 at
'a time it can be harmful to your system, it could
'low windows resources
   ' Error = 1
   ' Exclamation = 2
   ' Information = 3
   ' Update1 = 4
   ' Update2 = 5
   ' Question = 6
' You can change icon from 1 to 6
SetIcon (4)
SetOnTop (True)
TitleShadow (False)
TitleShadowColor (&HC0C0C0)
'For WindowsXP it is better to set buttonShadow False
'But for Windows 9x you can set it to True
buttonShadow (False)

If Len(Label2.Caption) > 30 Then
Label2.Caption = Left$(Label2.Caption, 30) & "..."
End If
With ClsGradient
        .Angle = 90
        .Color2 = &H80000002  '&H80000015 'RGB(255, 0, 0)
        .Color1 = &H80000003 '&H80000016 'RGB(242, 245, 240)
        .Draw Picture4
    End With
Call PlaySoundResource(106)
MakePopUpSize 1
  'Me.Top = 5890
Timer1.Enabled = True
Dim L As Long
WindowTransparency hWnd, byValue, , 0
Show
For L = 0 To 255 'Step 1 'You can also change fade in and fade out duration
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
Next L
perr:
Screen.MousePointer = 0
Exit Sub
err:
MsgBox err.Description, vbCritical
End
Resume perr:
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim L&
  For L = 255 To 0 Step -1
     WindowTransparency hWnd, byValue, , L
     DoEvents
     Refresh
  Next L
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
Label2.FontUnderline = False
Label2.MousePointer = 0
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = True
Label1.MousePointer = 99
End Sub

Private Sub Label2_Click()
Label2.FontUnderline = False
Label2.MousePointer = 0
'MsgBox Label4.Caption & vbNewLine & vbNewLine & "Created By: Muhammad Waqas Iqbal" & vbCrLf & "E-mail: mwaqasiq007@hotmail.com" & vbNewLine & "            pakistani_muslims@yahoo.com", vbInformation, "Alert Window - " & Label2.Caption
Form2.Show vbModal, Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label2.FontUnderline = True
Label2.MousePointer = 99
End Sub

Private Sub Option1_Click()
Command1.Caption = "Download"
Command1.Enabled = True
End Sub

Private Sub Option2_Click()
Command1.Enabled = True
Command1.Caption = "Close"
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Label1.FontUnderline = False
Label1.MousePointer = 0
Label2.FontUnderline = False
Label2.MousePointer = 0
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 100
If Me.Top <= resto Then
Timer1.Enabled = False
End If
Debug.Print Me.Top
End Sub

Private Sub Timer2_Timer()
Me.Top = Me.Top + 100
If Me.Top >= 10000 Then
Timer2.Enabled = False
MsgBox "Program will now close.", vbInformation
Unload Me
End If
End Sub
Private Sub CreateTransparentWindowStyle(lHwnd)
'-----------------------------------
'this is used to create the window style needed
'to allow transparency to be set/altered with
'calls to SetLayeredWindowAttributes
'-----------------------------------
 On Error GoTo Err_Handler:
 
'VARIABLES:
  Dim Ret As Long
'CODE:
       'Set the window style to 'Layered'
       Ret = GetWindowLong(lHwnd, GWL_EXSTYLE)
       Ret = Ret Or WS_EX_LAYERED
       SetWindowLong lHwnd, GWL_EXSTYLE, Ret
'END CODE:
 
Exit Sub
perr:
Screen.MousePointer = vbDefault
Exit Sub
Err_Handler:
    err.Source = err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox err.Number & vbTab & err.Source & err.Description, vbCritical
    err.Clear
    Resume perr:
End Sub




Private Sub WindowTransparency(lHwnd&, TransparencyBy As TransType, _
                                      Optional Clr As Long, _
                                      Optional TransVal As Long)
On Error GoTo Err_Handler:
'---------------------------------
'sets window transparency
'proper window style must be set first
'with call to CreateTransparentWindowStyle
'that call only has to be made once for the
'life of the form.  After that, this sub
'may be called multiple times by itself
'---------------------------------
'CODE:
    'first create the window style cabable of transparancies
    Call CreateTransparentWindowStyle(lHwnd)
    
    If TransparencyBy = byColor Then
         'the color specified in Clr becomes totally transparent
         SetLayeredWindowAttributes lHwnd, Clr, 0, LWA_COLORKEY
         
    ElseIf TransparencyBy = byValue Then
         If TransVal < 0 Or TransVal > 255 Then
            'makes sure valid transparency number chosen
            '0=totally opaque    255= totally transparent
            err.Raise 2222, "Sub WindowTransparency", _
                    "must choose number between 0-255"
         End If
         SetLayeredWindowAttributes lHwnd, 0, TransVal, LWA_ALPHA
    End If
'END CODE:
Exit Sub
perr:
Exit Sub
Screen.MousePointer = vbDefault
Exit Sub
Err_Handler:
    err.Source = err.Source & "." & VarType(Me) & ".ProcName"
    MsgBox err.Number & vbTab & err.Source & err.Description, vbCritical
Resume Next
End Sub
Private Sub SetOnTop(Value As Boolean)
 Dim i
If Value = True Then
i = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, vbNormal)
ElseIf Value = False Then
i = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, vbNormal)
End If

End Sub

Private Sub TitleShadow(Value As Boolean)
If Value = True Then
Label5.Caption = Label2.Caption
Label5.Visible = True
ElseIf Value = False Then
Label5.Visible = False
End If
End Sub

Private Sub TitleShadowColor(Value As OLE_COLOR)
Label5.ForeColor = Value
End Sub
Private Sub buttonShadow(Value As Boolean)
If Value = True Then
Label6.Visible = True
Label7.Visible = True
ElseIf Value = False Then
Label6.Visible = False
Label7.Visible = False
End If
End Sub

Private Sub SetIcon(Icon As IconType)
If Icon = 1 Then   ' Error = 1
Image1.Picture = ErrorIcon.Picture
ElseIf Icon = 2 Then   ' Exclamation = 2
Image1.Picture = ExclamationIcon.Picture
ElseIf Icon = 3 Then   ' Information = 3
Image1.Picture = InformationIcon.Picture
ElseIf Icon = 4 Then   ' Update1 = 4
Image1.Picture = Update1Icon.Picture
ElseIf Icon = 5 Then    ' Update2 = 5
Image1.Picture = Update2Icon.Picture
ElseIf Icon = 6 Then    ' Question = 6
Image1.Picture = QuestionIcon.Picture
End If
End Sub







