Public Class Form1

    Private Sub RibbonButton1_Click(sender As System.Object, e As System.EventArgs) Handles RibbonButton1.Click
        For Each f As Form In Me.MdiChildren
            If TypeOf (f) Is Form2 Then
                f.Activate()
                Return
            End If
        Next

        Dim f2 As New Form2()
        f2.MdiParent = Me
        f2.Show()
    End Sub

    Private Sub RibbonButton2_Click(sender As System.Object, e As System.EventArgs) Handles RibbonButton2.Click
        For Each f As Form In Me.MdiChildren
            If TypeOf (f) Is Form3 Then
                f.Activate()
                Return
            End If
        Next

        Dim f2 As New Form3()
        f2.MdiParent = Me
        f2.Show()
    End Sub

    Private Sub RibbonButton_Close_Click(sender As System.Object, e As System.EventArgs) Handles RibbonButton_Close.Click
        While Me.ActiveMdiChild IsNot Nothing
            Me.ActiveMdiChild.Close()
        End While
    End Sub
End Class
