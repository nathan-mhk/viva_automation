Private Sub Form_Load()
Me.TimerInterval = 5000
End Sub

Private Sub Form_Timer()
Me.TimerInterval = 0
'Me.InsideHeight = 1
'Me.InsideWidth = 1

Forms.frmlineinfo.SetFocus
Call Forms.frmlineinfo.AutoRead_Click
End Sub
