Public Class Form2
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        MyBase.OnLoad(e)
        Me.ControlBox = False
        Me.WindowState = FormWindowState.Maximized
        Me.BringToFront()
    End Sub
End Class