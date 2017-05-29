Public Class UserEntry

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
        'Quitar la siguiente linea para version final
        Button1.Enabled = True
        Return
        '********************************************
        ValidaCedula(TextBox1.Text.Trim())
    End Sub

    Public Function ValidaCedula(ByVal numCedula As String) As Boolean
        Dim suma As Short
        Dim dv As Short
        Dim d1 As Short
        Dim d2 As Short
        Dim d3 As Short
        Dim d4 As Short
        Dim d5 As Short
        Dim d6 As Short
        Dim d7 As Short
        Dim d8 As Short
        Dim d9 As Short
        If Len(TextBox1.Text.Trim) = 0 Then
            Return False
        End If
        If Len(TextBox1.Text.Trim) = 10 Then
            If TextBox1.Text.Substring(0, 2) > 22 Then
                MsgBox("Error: Primeros 2 dígitos mayores a 22", MsgBoxStyle.OkOnly)
                TextBox1.Focus()
                Return False
            ElseIf TextBox1.Text.Substring(2, 1) > 5 Then
                MsgBox("Error: Tercer dígito mayor a 5", MsgBoxStyle.OkOnly)
                TextBox1.Focus()
                Return False
            Else
                d1 = IIf(TextBox1.Text.Substring(0, 1) * 2 > 9, TextBox1.Text.Substring(0, 1) * 2 - 9, TextBox1.Text.Substring(0, 1) * 2)
                d2 = TextBox1.Text.Substring(1, 1)
                d3 = IIf(TextBox1.Text.Substring(2, 1) * 2 > 9, TextBox1.Text.Substring(2, 1) * 2 - 9, TextBox1.Text.Substring(2, 1) * 2)
                d4 = TextBox1.Text.Substring(3, 1)
                d5 = IIf(TextBox1.Text.Substring(4, 1) * 2 > 9, TextBox1.Text.Substring(4, 1) * 2 - 9, TextBox1.Text.Substring(4, 1) * 2)
                d6 = TextBox1.Text.Substring(5, 1)
                d7 = IIf(TextBox1.Text.Substring(6, 1) * 2 > 9, TextBox1.Text.Substring(6, 1) * 2 - 9, TextBox1.Text.Substring(6, 1) * 2)
                d8 = TextBox1.Text.Substring(7, 1)
                d9 = IIf(TextBox1.Text.Substring(8, 1) * 2 > 9, TextBox1.Text.Substring(8, 1) * 2 - 9, TextBox1.Text.Substring(8, 1) * 2)
                suma = (d1 + d2 + d3 + d4 + d5 + d6 + d7 + d8 + d9)
                dv = Math.Ceiling(suma / 10)
                dv = dv * 10 - suma
                If TextBox1.Text.Substring(9, 1) = dv Then
                    Button1.Enabled = True
                    Return True
                Else
                    MsgBox("Último dígito incorrecto", MsgBoxStyle.OkOnly)
                    TextBox1.Focus()
                    Return False
                End If
            End If
        Else
            MsgBox("Error: Número de dígitos incompleto", MsgBoxStyle.OkOnly)
            TextBox1.Focus()
            Return False
        End If
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Mover las siguientes lineas para version final
        MainForm.GestiónDocumentalToolStripMenuItem.Visible = True
        MainForm.GestiónDocumentalToolStripMenuItem.Enabled = True
        MainForm.InscripcionesToolStripMenuItem.Visible = True
        MainForm.InscripcionesToolStripMenuItem.Enabled = True
        MainForm.CertificacionesToolStripMenuItem.Visible = True
        MainForm.CertificacionesToolStripMenuItem.Enabled = True
        MainForm.FolioPersonalToolStripMenuItem.Visible = True
        MainForm.FolioPersonalToolStripMenuItem.Enabled = True
        MainForm.ReportesToolStripMenuItem.Visible = False
        MainForm.FolioRealToolStripMenuItem.Visible = True
        MainForm.FolioRealToolStripMenuItem.Enabled = True
        MainForm.HerramientasToolStripMenuItem.Visible = True
        MainForm.HerramientasToolStripMenuItem.Enabled = True
        'remover para produccion
        MainForm.nombres = "Rodrigo Xavier"
        MainForm.apellidos = "Salguero Carvajal"
        MainForm.identificacion = "1706551171"
        MainForm.ToolStripStatusLabel1.Text = "Bienvenido " & MainForm.nombres & " " & MainForm.apellidos & ". Hoy es " + Today.ToLongDateString
        '        MsgBox(My.Settings.BD_TRADEREConnectionString, vbOKOnly)
        Me.Close()
        Return
        '**************************************************

        Dim reg_contraseña As String
        Dim registros() As DataRow = BD_TRADEREDataSet.Usuarios.Select("identificacion=" & TextBox1.Text.Trim)
        '        MsgBox(TestEncoding(TextBox2.Text.Trim), MsgBoxStyle.OkOnly)
        '     My.Computer.FileSystem.WriteAllText("c:\users\rpq01\documents\ultimopassword.txt", TestEncoding(TextBox2.Text.Trim), False)
        If registros.Count > 0 Then
            reg_contraseña = registros(0).Item("contraseña")
            MainForm.nombres = registros(0).Item("nombres").trim()
            MainForm.apellidos = registros(0).Item("apellidos").trim()
            MainForm.identificacion = TextBox1.Text.Trim
        Else
            MsgBox("Usuario no encontrado!", MsgBoxStyle.OkOnly)
            TextBox2.Text = ""
            TextBox1.Focus()
            Return
        End If
        If TestDecoding(reg_contraseña) = TextBox2.Text.Trim Then
            MainForm.GestiónDocumentalToolStripMenuItem.Visible = True
            MainForm.GestiónDocumentalToolStripMenuItem.Enabled = True
            MainForm.InscripcionesToolStripMenuItem.Visible = True
            MainForm.InscripcionesToolStripMenuItem.Enabled = True
            MainForm.CertificacionesToolStripMenuItem.Visible = True
            MainForm.CertificacionesToolStripMenuItem.Enabled = True
            MainForm.FolioPersonalToolStripMenuItem.Visible = True
            MainForm.FolioPersonalToolStripMenuItem.Enabled = True
            MainForm.ReportesToolStripMenuItem.Visible = False
            MainForm.FolioRealToolStripMenuItem.Visible = True
            MainForm.FolioRealToolStripMenuItem.Enabled = True
            MainForm.HerramientasToolStripMenuItem.Visible = True
            MainForm.HerramientasToolStripMenuItem.Enabled = True
            '            MsgBox(TestDecoding(reg_contraseña), MsgBoxStyle.OkOnly)
            MainForm.ToolStripStatusLabel1.Text = "Bienvenido " & MainForm.nombres & " " & MainForm.apellidos & ". Hoy es " + Today.ToLongDateString
            enviarnotificacion(MainForm.nombres & " " & MainForm.apellidos, Me.TextBox1.Text.Trim())
            Me.Close()
        Else
            MsgBox("Contraseña incorrecta!", MsgBoxStyle.OkOnly)
            TextBox2.Text = ""
            TextBox2.Focus()
            Return
        End If

    End Sub

    Private Sub TextBox2_Click(sender As Object, e As EventArgs) Handles TextBox2.Click
        If ValidaCedula(TextBox1.Text.Trim()) Then
            TextBox2.Focus()
        Else
            TextBox1.Focus()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        TestDecoding("")
    End Sub


    Function TestEncoding(ByVal plaintext As String)
        Dim password As String = "123456"
        Dim wrapper As New Simple3Des(password)
        Dim cipherText As String = wrapper.EncryptData(plaintext)

        '        MsgBox("The cipher text is: " & cipherText)
        '        My.Computer.FileSystem.WriteAllText(My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt", cipherText, False)
        Return cipherText
    End Function

    Function TestDecoding(ByVal cipherText As String)
        '        Dim cipherText As String = My.Computer.FileSystem.ReadAllText(
        '       My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\cipherText.txt")
        Dim password As String = "123456"
        Dim wrapper As New Simple3Des(password)
        Dim plaintext As String = ""
        ' DecryptData throws if the wrong password is used.
        Try
            plaintext = wrapper.DecryptData(cipherText)
        Catch ex As System.Security.Cryptography.CryptographicException
            MsgBox("The data could not be decrypted with the password.")
        End Try
        Return plaintext
    End Function

    Private Sub enviarnotificacion(ByVal nombre_usuario As String, identificacion As String)
        ' Quitar Return para version final
        Return
        '*********************************
        Dim EmailMessage As New MailMessage()
        Try
            EmailMessage.From = New MailAddress("notificaciones.tradere@gmail.com")
            EmailMessage.To.Add("rodrigosalguero@live.com")
            EmailMessage.Subject = "Sistema Registral TRADERE"
            EmailMessage.Body = "Login de " + nombre_usuario + "(" + identificacion + ")." + Chr(10) + Chr(13) + "Efectuado el " + Today.Date.ToLongDateString + " a las " + TimeOfDay.ToShortTimeString + "."
            Dim SMTP As New SmtpClient("smtp.gmail.com")
            SMTP.Port = 587
            SMTP.EnableSsl = True
            SMTP.Credentials = New System.Net.NetworkCredential("notificaciones.tradere@gmail.com", "0669rxsc")
            SMTP.Send(EmailMessage)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmUsersEntry_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'BD_TRADEREDataSet.Usuarios' table. You can move, or remove it, as needed.
        Me.UsuariosTableAdapter.Fill(Me.BD_TRADEREDataSet.Usuarios)
        TextBox1.Focus()
    End Sub

End Class