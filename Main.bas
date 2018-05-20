Attribute VB_Name = "DBModule"
'-----------------------------------------------------'
'                Royal Dictionary v1.0                '
'-----------------------------------------------------'
'                                                     '
'                  Programmer: R@MiN                  '
'                 03250422@yahoo.com                  '
'                  www.mobilebaz.ir                   '
'                                                     '
'-----------------------------------------------------'
Global DBMain As Database
Global RecSet As Recordset
Public Sub Main()
    'Load frmMain
    frmMain.Show
End Sub
Public Sub QueryData(reqText As String)
Dim SQLText As String
    SQLText = "SELECT *"
    SQLText = SQLText + " FROM Words"
    SQLText = SQLText + " WHERE English LIKE '" & reqText & "*';"   '''' & "*';"
    
    'And Create a Recordset object with this SQLText
    Set RecSet = DBMain.OpenRecordset(SQLText)
    'Clear the list box
    frmMain.lstWords.Clear
    If RecSet.RecordCount = 0 Then
       frmMain.txtExp(0).Text = "ò·„Â „Ê—œ ‰Ÿ— Ì«›  ‰‘œ."
        Exit Sub
    End If
    RecSet.MoveLast: RecSet.MoveFirst
    'fill the list box
    Do Until RecSet.EOF
        frmMain.lstWords.AddItem RecSet.Fields(0)
        RecSet.MoveNext
    Loop
     
    If Not frmMain.lstWords.ListCount = 0 Then
        RecSet.MoveFirst
        frmMain.txtExp(0).Text = RecSet.Fields(1)
    Else
     frmMain.txtExp(0).Text = ""
    End If
End Sub
