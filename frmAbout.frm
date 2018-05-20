VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Dictionary"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   " „«” »« „«"
      Height          =   400
      Left            =   1440
      TabIndex        =   1
      Top             =   2460
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   " «ÌÌœ"
      Height          =   400
      Left            =   240
      TabIndex        =   0
      Top             =   2460
      Width           =   1100
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   240
      Picture         =   "frmAbout.frx":74F2
      Top             =   90
      Width           =   825
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0099A8AC&
      X1              =   240
      X2              =   4130
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Ê«éÂ ‰«„Â —ÊÌ«· ‰ê«—‘ 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1905
      TabIndex        =   5
      Top             =   255
      Width           =   2220
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   ".»Â ê—ÊÂ ‰—„ «›“«—Ì —ÊÌ«· „Ì »«‘œ"
      Height          =   270
      Left            =   1500
      TabIndex        =   4
      Top             =   1395
      Width           =   2610
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ò·ÌÂ ÕﬁÊﬁ «Ì‰ ‰—„ «›“«— Ê Å—Ê‰œÂ Â«Ì „—»Êÿ »Â ¬‰ „ ⁄·ﬁ "
      Height          =   240
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "»—‰«„Â ‰ÊÌ” : „Ê»«Ì· »«“"
      Height          =   285
      Left            =   1770
      TabIndex        =   2
      Top             =   555
      Width           =   2340
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   240
      X2              =   4130
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------'
'                Royal Dictionary v1.0                '
'-----------------------------------------------------'
'                                                     '
'                  Programmer: R@MiN                  '
'                 03250422@yahoo.com                  '
'                  www.mobilebaz.ir                   '
'                                                     '
'-----------------------------------------------------'

Private Sub Command1_Click()
Unload Me
If Not frmMain.WindowState = vbMinimized Then
frmMain.txtWord.SetFocus
End If
End Sub

Private Sub Command2_Click()
Shell "explorer.exe http://www.mobilebaz.ir/", vbNormalFocus
End Sub

Private Sub Form_Load()
SetIcon Me.hWnd, "AAA"
'Skn.LoadSkin App.Path & "\Skin\" & "winaqua.skn" ' Loads another skin into Skin component
'Skn.ApplySkin Me.hwnd ' Applies the skin to this window and its child controls
End Sub
