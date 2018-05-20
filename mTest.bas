Attribute VB_Name = "mTest"
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
' Windowless and class array project demostration
'// Windows messages
Private Const WM_MOUSEMOVE                                  As Long = &H200
Private Const WM_RBUTTONDBLCLK                              As Long = &H206
Private Const WM_RBUTTONDOWN                                As Long = &H204
Private Const WM_RBUTTONUP                                  As Long = &H205
Private Const WM_MBUTTONDBLCLK                              As Long = &H209
Private Const WM_MBUTTONDOWN                                As Long = &H207
Private Const WM_MBUTTONUP                                  As Long = &H208
Private Const WM_LBUTTONDBLCLK                              As Long = &H203
Private Const WM_LBUTTONDOWN                                As Long = &H201
Private Const WM_LBUTTONUP                                  As Long = &H202
Private Const WM_USER                                       As Long = &H400

'// Balloon messages
Private Const NIN_BALLOONSHOW                               As Long = (WM_USER + 2)
Private Const NIN_BALLOONHIDE                               As Long = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT                            As Long = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK                          As Long = (WM_USER + 5)

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Sub InitCommonControls Lib "Comctl32" ()

Private m_cSystray(4) As cSystray
Private m_bCancel As Boolean
Public Function EventsProc(ByVal lhWnd As Long, _
       ByVal lMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long) As Long
       
    Select Case lMsg
        Case WM_LBUTTONDBLCLK:                  Debug.Print "MouseDblClick", IndexFromWnd(lhWnd)
        Case WM_LBUTTONDOWN:                    Debug.Print "MouseDown(vbLeftButton)", IndexFromWnd(lhWnd)
        Case WM_LBUTTONUP:                      Debug.Print "MouseUp(vbLeftButton)", IndexFromWnd(lhWnd)
            'Flag to exit the main loop
            m_bCancel = True
        Case WM_MBUTTONDBLCLK:                  Debug.Print "MouseDblClick(vbMiddleButton)", IndexFromWnd(lhWnd)
        Case WM_MBUTTONDOWN:                    Debug.Print "MouseDown(vbMiddleButton)", IndexFromWnd(lhWnd)
        Case WM_MBUTTONUP:                      Debug.Print "MouseUp(vbMiddleButton)", IndexFromWnd(lhWnd)
        Case WM_RBUTTONDBLCLK:                  Debug.Print "MouseDblClick(vbRightButton)", IndexFromWnd(lhWnd)
        Case WM_RBUTTONDOWN:                    Debug.Print "MouseDown(vbRightButton)", IndexFromWnd(lhWnd)
        Case WM_RBUTTONUP:                      Debug.Print "MouseUp(vbRightButton)", IndexFromWnd(lhWnd)
            'The wparam is our extraparam and in _
            this case is used to get the index of the class
            m_cSystray(wParam).BalloonShow
        Case WM_MOUSEMOVE:                      Debug.Print "MouseMove", IndexFromWnd(lhWnd)
        Case NIN_BALLOONUSERCLICK:              Debug.Print "BalloonClick", IndexFromWnd(lhWnd)
        Case NIN_BALLOONTIMEOUT:                Debug.Print "BalloonClose", IndexFromWnd(lhWnd)
        Case NIN_BALLOONSHOW:                   Debug.Print "BalloonShow", IndexFromWnd(lhWnd)
        Case NIN_BALLOONHIDE:                   Debug.Print "BalloonHide", IndexFromWnd(lhWnd)
    End Select
    
End Function

Private Function IndexFromWnd(ByVal hWnd As Long) As Long
    Dim i As Long
    
    For i = 0 To UBound(m_cSystray)
        If m_cSystray(i).lhWnd = hWnd Then
            IndexFromWnd = i
            Exit Function
        End If
    Next
End Function
