Public Class ValidatedControlsForms
    Private frmIndexacion As New frmIndexacion()
    Dim errores As Boolean = False

    Public Function check() As Boolean

        If frmIndexacion.TextBox2.Text.Trim = "" Then
            frmIndexacion.Label3.BackColor = Color.Red
            frmIndexacion.Label3.ForeColor = Color.White
            errores = True
        Else
            frmIndexacion.Label3.BackColor = Color.Transparent
            frmIndexacion.Label3.ForeColor = Color.Black
            errores = False
        End If

        If frmIndexacion.TextBox3.Text.Trim = "" Then
            frmIndexacion.Label6.BackColor = Color.Red
            frmIndexacion.Label6.ForeColor = Color.White
            errores = True
        Else
            frmIndexacion.Label6.BackColor = Color.Transparent
            frmIndexacion.Label6.ForeColor = Color.Black
            errores = False
        End If

        If (frmIndexacion.ComboBox1.SelectedIndex = -1) Then
            frmIndexacion.Label5.BackColor = Color.Red
            frmIndexacion.Label5.ForeColor = Color.White
            errores = True
        Else
            frmIndexacion.Label5.BackColor = Color.Transparent
            frmIndexacion.Label5.ForeColor = Color.Black
            errores = False
        End If

        If (frmIndexacion.ComboBox2.SelectedIndex = -1) Then
            frmIndexacion.Label9.BackColor = Color.Red
            frmIndexacion.Label9.ForeColor = Color.White
            errores = True
        Else
            frmIndexacion.Label9.BackColor = Color.Transparent
            frmIndexacion.Label9.ForeColor = Color.Black
            errores = False
        End If

        Return errores

    End Function

End Class
