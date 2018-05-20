VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Royal Dictionary"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   3960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClipboard 
      Enabled         =   0   'False
      Interval        =   2
      Left            =   0
      Top             =   -135
   End
   Begin VB.CommandButton Command3 
      Caption         =   " ·›Ÿ"
      Enabled         =   0   'False
      Height          =   400
      Left            =   165
      TabIndex        =   5
      Top             =   3390
      Width           =   1100
   End
   Begin VB.CommandButton Command2 
      Caption         =   "...œ—»«—Â"
      Height          =   400
      Left            =   1440
      TabIndex        =   4
      Top             =   3390
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»” ‰"
      Height          =   400
      Left            =   2715
      TabIndex        =   3
      Top             =   3390
      Width           =   1100
   End
   Begin VB.ListBox lstWords 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   180
      TabIndex        =   2
      Top             =   1005
      Width           =   3600
   End
   Begin VB.TextBox txtWord 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   345
      Width           =   3600
   End
   Begin VB.TextBox txtExp 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Index           =   0
      Left            =   180
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2175
      Width           =   3600
   End
   Begin VB.Label Label3 
      Caption         =   ":  —Ã„Â"
      Height          =   240
      Left            =   3225
      TabIndex        =   8
      Top             =   1950
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   ": (Ì«› Â (Â«"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   765
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   ": Ê«éÂ"
      Height          =   255
      Left            =   3405
      TabIndex        =   6
      Top             =   105
      Width           =   375
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu Show_menu 
         Caption         =   "‰„«Ì‘"
         Visible         =   0   'False
      End
      Begin VB.Menu M_menu 
         Caption         =   "‘ò«—çÌ ·€ "
         Begin VB.Menu Word_menu 
            Caption         =   "›⁄«·"
            Checked         =   -1  'True
         End
         Begin VB.Menu blank3 
            Caption         =   "-"
         End
         Begin VB.Menu Ballon_menu 
            Caption         =   "‰„«Ì‘  —Ã„Â œ— »«·Ê‰"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu About_menu 
         Caption         =   "œ—»«—Â Ê«éÂ ‰«„Â —ÊÌ«·"
      End
      Begin VB.Menu blank1 
         Caption         =   "-"
      End
      Begin VB.Menu Close_menu 
         Caption         =   "»” ‰"
      End
   End
End
Attribute VB_Name = "frmMain"
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
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32" () As Long
Private WithEvents f_cSystray As cSystray
Attribute f_cSystray.VB_VarHelpID = -1
Public Speech As SpVoice
Private mblnCaptureClipboard As Boolean
Private BallonVar As Boolean
'
Public Enum FilterConstants
   fltUpper
   fltLower
   fltNumber
   fltPunctuation
   fltUpperLower
   fltUpperNumber
   fltUpperPunctuation
   fltUpperNumberPunctuation
   fltLowerNumber
   fltLowerPunctuation
   fltLowerNumberPunctuation
   fltNumberPunctuation
End Enum
'
Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Ballon_menu_Click()
    Ballon_menu.Checked = Not Ballon_menu.Checked
    BallonVar = Ballon_menu.Checked
End Sub

Private Sub btnStartStop_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set f_cSystray = Nothing
End Sub
Private Sub About_menu_Click()
frmAbout.Show 1
End Sub

Private Sub Close_menu_Click()
Unload Me
End Sub
Private Sub f_cSystray_MouseUp(Button As Integer)
    f_cSystray.BeforePopup
    PopupMenu mnuPopup
End Sub
Private Sub f_cSystray_MouseDblClick(Button As Integer)
             Me.WindowState = vbNormal   ' Or vbMaximized if you feel like it.
             Me.Show
             Set f_cSystray = Nothing
End Sub
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
frmAbout.Show 1
End Sub
Private Sub Command3_Click()
Set Speech = New SpVoice
If Not (txtExp(0).Text = "ò·„Â „Ê—œ ‰Ÿ— Ì«›  ‰‘œ.") Then
'Speech.Speak txtWord.Text
Speech.Speak (lstWords.List(0))
End If
End Sub
Private Sub Form_Load()

    If Command = "-t" Then
        Me.WindowState = vbMinimized
        Me.Hide
        Set f_cSystray = New cSystray
        With f_cSystray
            .SysTrayIconFromRes "AAA"
            .SysTrayToolTip = "Ê«éÂ ‰«„Â —ÊÌ«·"
            .SysTrayShow True
        End With
    End If
    
SetIcon Me.hWnd, "AAA"
'Skn.LoadSkin App.Path & "\Skin\" & "winaqua.skn" ' Loads skin into Skin component
'Skn.ApplySkin Me.hwnd ' Applies the skin to this window and its child controls
Set DBMain = OpenDatabase(App.Path & "\Data.db", True, False, ";pwd=" & "Password Not Found")
Clipboard.Clear
tmrClipboard.Enabled = True
Ballon_menu.Checked = False
End Sub

Private Sub lstWords_Click()
txtWord.Text = lstWords.List(lstWords.ListIndex)
'Call QueryData(lstWords.List(lstWords.ListIndex))
txtWord.SelLength = Len(txtWord.Text)
End Sub
Private Sub tmrClipboard_Timer()
ReadClipBoard
End Sub
Private Sub ReadClipBoard()
Static lastClip As String
Static ctm As Integer
Dim currentClip As String
'take only 1 out of 10 reads
ctm = ctm + 1: If ctm > 10 Then ctm = 0
If ctm > 0 Then Exit Sub
On Error GoTo noClipRead
currentClip = Clipboard.GetText
'=============================================
If Me.WindowState = vbMinimized Then
If BallonVar = True Then
'''''''''''''''''''''''''''''''''''''''''''''''''''
If Not currentClip = "" Then
   If currentClip <> lastClip Then
             lastClip = currentClip
             txtWord.Text = currentClip
             ''''''''''''''''''''
             f_cSystray.BalloonTitle = currentClip
             f_cSystray.BalloonText = txtExp(0).Text
             f_cSystray.BalloonIcon = TTIconUser
             f_cSystray.BalloonShow True, 5000
    End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''
Else
''''''''''''''''''''''''''''''''''''''''''''''''''
If Not currentClip = "" Then
If currentClip <> lastClip Then
    lastClip = currentClip
    txtWord.Text = currentClip
    ''''''''''''''''''''
    Me.WindowState = vbNormal
    Me.Show
    txtWord.SelLength = Len(txtWord.Text)
    Set f_cSystray = Nothing
    ''''''''''''''''''''
End If
End If
End If
Else
'''''''''''''''''''''''''''''''''''''
If Not currentClip = "" Then
If currentClip <> lastClip Then
    lastClip = currentClip
    txtWord.Text = currentClip
    txtWord.SelLength = Len(txtWord.Text)
    Me.SetFocus
    End If
End If
'''''''''''''''''''''''''''''''''''''
'Clipboard.Clear
noClipRead:
End If
'=============================================
End Sub

'Clipboard.SetText thehtmlcode.Text
Private Sub txtWord_Change()
    If txtWord.Text <> "" Then
    Call QueryData(txtWord.Text)
    End If
    If (txtExp(0).Text = "ò·„Â „Ê—œ ‰Ÿ— Ì«›  ‰‘œ.") Or txtWord.Text = "" Then
    Command3.Enabled = False
    Else
    Command3.Enabled = True
    End If
End Sub
Private Sub Form_Activate()
    txtWord.SelLength = Len(txtWord.Text)
End Sub
Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
        Me.Hide
Set f_cSystray = New cSystray
With f_cSystray
                '.SysTrayIconFromFile App.Path & "\01.ico"
                .SysTrayIconFromRes "AAA"
                '.SysTrayIconFromCompRes "shell32.dll", 130
        .SysTrayToolTip = "Ê«éÂ ‰«„Â —ÊÌ«·"
        .SysTrayShow True
        'If .IsBalloonCapable Then .BalloonShow True, 3000
End With
''''''''''''''''''''''''''''''''
    End If
End Sub

Private Sub txtWord_GotFocus()
    txtWord.SelStart = 0
    txtWord.SelLength = Len(txtWord)
End Sub

Private Sub txtWord_KeyPress(KeyAscii As Integer)
KeyAscii = AllowKeys(KeyAscii, fltUpperLower)
End Sub

Private Sub Word_menu_Click()
    Word_menu.Checked = Not Word_menu.Checked
    If Word_menu.Checked = True Then
    Ballon_menu.Enabled = True
    Else
    Ballon_menu.Enabled = False
    End If
    mblnCaptureClipboard = Word_menu.Checked
    tmrClipboard.Enabled = mblnCaptureClipboard
End Sub
Public Function AllowKeys(KeyAscii As Integer, KeysToAllow As FilterConstants, Optional ShowErrMsg As Boolean) As Integer
'Note: all the filters allow the space(32) and back space keys

Dim msg As String

Select Case KeysToAllow
  Case fltLower 'allow lowercase letters only
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Then
          AllowKeys = KeyAscii
        ElseIf KeyAscii < 97 Or KeyAscii > 122 Then
          AllowKeys = 0
          msg = "Only lower case letters are allowed. (E.g. 'a' but not 'A')"
        Else
          AllowKeys = KeyAscii
        End If
        
  Case fltUpper 'allow uppercase letters only
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Then
          AllowKeys = KeyAscii
        ElseIf KeyAscii < 65 Or KeyAscii > 90 Then
          AllowKeys = 0
          msg = "Only upper case letters are allowed. (E.g. 'A' but not 'a')"
        Else
          AllowKeys = KeyAscii
        End If
        
  Case fltNumber 'allow numbers only.The 46 is for the decimal point
        If KeyAscii = vbKeyBack Or KeyAscii = vbKeySpace Or KeyAscii = 46 Then
          AllowKeys = KeyAscii
        ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
          AllowKeys = 0
          msg = "Only digits are allowed. (E.g. '0','3' or '9')"
        Else
          AllowKeys = KeyAscii
        End If
        
  Case fltPunctuation 'allow punctuation keys only
        Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only punctuatuon characters are allowed. (E.g. '?', ';' or ',')"
        End Select
        
   Case fltUpperLower  'allow upper and lowercase letters only
        Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 65 To 90, 97 To 122
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only lower case and upper case letters are allowed. (E.g. 'a','b','A' or 'B')"
        End Select
        
   Case fltUpperNumber 'allow uppercase letters and numbers only
        Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 65 To 90, 48 To 57
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only upper case letters and digits are allowed. (E.g. '0','5','A' or 'B')"
        End Select
        
  Case fltUpperPunctuation 'allow uppercase letters and punctuation only
        Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 65 To 90, 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only upper case letters and punctuation characters are allowed. (E.g. '?',';','A' or 'B')"
        End Select
        
 Case fltLowerNumber 'allow lowercase letters and numbers only
        Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 97 To 122, 48 To 57
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only lower case letters and digits are allowed. (E.g. '0','5','a' or 'b')"
        End Select
        
 Case fltLowerPunctuation 'allow lowercase letters and punctuation only
       Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 97 To 122, 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only lower case letters and punctuation characters are allowed. (E.g. '?',';','a' or 'b')"
        End Select
        
 Case fltNumberPunctuation 'allow numbers and punctuation only
       Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 48 To 57, 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only digits and punctuation characters are allowed. (E.g. '?',';','6' or '2')"
        End Select
        
  Case fltUpperNumberPunctuation 'allow uppercase,numbers and punctuation only
         Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 65 To 90, 48 To 57, 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only upper case letters, digits and punctuation characters are allowed. (E.g. '?','A' or '9')"
        End Select
        
   Case fltLowerNumberPunctuation 'allow lowercase,numbers and punctuation only
         Select Case KeyAscii
           Case vbKeyBack
             AllowKeys = KeyAscii
           Case vbKeySpace
             AllowKeys = KeyAscii
           Case 48 To 57, 97 To 122, 33 To 47, 58 To 64, 91 To 96, 123 To 126
             AllowKeys = KeyAscii
           Case Else
             AllowKeys = 0
             msg = "Only lower case letters, digits and punctuation characters are allowed. (E.g. '?','a' or '9')"
        End Select
 End Select
 
 If AllowKeys = 0 And ShowErrMsg = True Then
    MsgBox msg, vbInformation + vbOKOnly, "Invalid key"
 End If
End Function

