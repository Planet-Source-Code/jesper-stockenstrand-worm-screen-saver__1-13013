Attribute VB_Name = "Module1"
Sub Main()
Dim strCmdLine As String
    strCmdLine = Left(Command, 2)


    If strCmdLine = "/p" Then
        'End
        'on select of your screen saver
    ElseIf strCmdLine = "/c" Then
        MsgBox "Worm (C)2000 by Jesper Stockenstrand", vbOKOnly, "Info"
        
        'Call Load(frmConfig)
        'frmConfig.Show
        'Function to call when
        'Settings'
        'button is pushed
    Else
        Call Load(Form1)
       Form1.Show
        'your screen saver has been loaded
    End If

End Sub
