VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Whats New"
   ClientHeight    =   4740
   ClientLeft      =   2925
   ClientTop       =   1140
   ClientWidth     =   6060
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   1935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form2.frx":000C
      Top             =   2040
      Width           =   5775
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1695
      Left            =   -120
      ScaleHeight     =   1635
      ScaleWidth      =   6315
      TabIndex        =   2
      Top             =   -120
      Width           =   6375
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........................."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   27.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   4875
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Window"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   27.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   765
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3630
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H80000002&
         BorderWidth     =   14
         Height          =   975
         Left            =   360
         Top             =   360
         Width           =   975
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H80000002&
         BorderWidth     =   9
         Height          =   855
         Left            =   4440
         Top             =   600
         Width           =   855
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000D&
         BorderWidth     =   9
         FillColor       =   &H00004080&
         Height          =   855
         Left            =   4800
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000D&
         BorderWidth     =   12
         FillColor       =   &H00004080&
         Height          =   1095
         Left            =   -240
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Caption         =   "pakistani_muslims@yahoo.com"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "E-mail: mwaqasiq007@hotmail.com"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Muhammad Waqas Iqbal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3000
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ClsGradient As New CGradient
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
AlwaysOnTop Form2, True
With ClsGradient
        .Angle = 90
        .Color2 = &H80000002
        .Color1 = &H80000003
        .Draw Picture4
    End With
End Sub


