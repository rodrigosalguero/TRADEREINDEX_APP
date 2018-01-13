Public Class ValidatedControlsForms

    Dim errores As Boolean = False

    Public Function check() As List(Of Control)

        Dim controlsLabel As New List(Of Label)
        Dim controlsResturn As New List(Of Control)

        If frmIndexacion.TextBox2.Text.Trim = "" Then
            controlsResturn.Add(frmIndexacion.Label3)
        Else
            frmIndexacion.Label3.BackColor = Color.Transparent
            frmIndexacion.Label3.ForeColor = Color.Black
        End If

        If frmIndexacion.TextBox3.Text.Trim = "" Then
            controlsResturn.Add(frmIndexacion.Label6)
        Else
            frmIndexacion.Label6.BackColor = Color.Transparent
            frmIndexacion.Label6.ForeColor = Color.Black
        End If


        If (frmIndexacion.ComboBox1.SelectedIndex = -1) Then
            controlsResturn.Add(frmIndexacion.Label5)
        Else
            frmIndexacion.Label5.BackColor = Color.Transparent
            frmIndexacion.Label5.ForeColor = Color.Black
        End If

        If (frmIndexacion.ComboBox5.SelectedIndex = -1) Then
            controlsResturn.Add(frmIndexacion.Label14)
        Else
            frmIndexacion.Label14.BackColor = Color.Transparent
            frmIndexacion.Label14.ForeColor = Color.Black
        End If

        If (frmIndexacion.DateTimePicker1.Text.Trim = "/  /") Then
            controlsResturn.Add(frmIndexacion.Label2)
        Else
            frmIndexacion.Label2.BackColor = Color.Transparent
            frmIndexacion.Label2.ForeColor = Color.Black
        End If

        If (frmIndexacion.MaskedTextBox2.Text.Trim = "/  /") Then
            controlsResturn.Add(frmIndexacion.Label19)
        Else
            frmIndexacion.Label19.BackColor = Color.Transparent
            frmIndexacion.Label19.ForeColor = Color.Black
        End If

        If (frmIndexacion.TextBox8.Text.Trim = "") Then
            controlsResturn.Add(frmIndexacion.Label9)
        Else
            frmIndexacion.Label9.BackColor = Color.Transparent
            frmIndexacion.Label9.ForeColor = Color.Black
        End If

        If (frmIndexacion.ComboBox4.SelectedIndex = -1) Then
            controlsResturn.Add(frmIndexacion.Label12)
        Else
            frmIndexacion.Label12.BackColor = Color.Transparent
            frmIndexacion.Label12.ForeColor = Color.Black
        End If

        If frmIndexacion.TextBox6.Text.Trim = "" Then
            controlsResturn.Add(frmIndexacion.Label15)
        Else
            frmIndexacion.Label15.BackColor = Color.Transparent
            frmIndexacion.Label15.ForeColor = Color.Black
        End If

        If (frmIndexacion.DateTimePicker2.Text.Trim = "/  /") Then
            controlsResturn.Add(frmIndexacion.Label16)
        Else
            frmIndexacion.Label16.BackColor = Color.Transparent
            frmIndexacion.Label16.ForeColor = Color.Black
        End If

        If (frmIndexacion.ComboBox6.SelectedIndex = -1) Then
            controlsResturn.Add(frmIndexacion.Label10)
        Else
            frmIndexacion.Label10.BackColor = Color.Transparent
            frmIndexacion.Label10.ForeColor = Color.Black
        End If

        If (frmIndexacion.ComboBox7.SelectedIndex = -1) Then
            controlsResturn.Add(frmIndexacion.Label15)
        Else
            frmIndexacion.Label15.BackColor = Color.Transparent
            frmIndexacion.Label15.ForeColor = Color.Black
        End If

        Return controlsResturn

    End Function

End Class
