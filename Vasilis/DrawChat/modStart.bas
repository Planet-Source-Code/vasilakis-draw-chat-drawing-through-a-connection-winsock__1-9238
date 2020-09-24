Attribute VB_Name = "modStart"

Sub Main()
frmSup.Show
frmSup.Refresh
End Sub


Sub CF(frm As Form)
frm.Left = (Screen.Width / 2) - (frm.Width / 2)
frm.Top = (Screen.Height / 2) - (frm.Height / 2)
End Sub


