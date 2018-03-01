Imports System
Imports System.IO
Imports System.Text

Public Class UserEntry
    Public rd_cedula As String
    Public rd_nombre As String
    Public rd_clave As String
    Dim variables As New VariablesGlobalesYfunciones
    Public modeUsers As Integer = 0 ''0=> digitador, 1=> inspector,2 => create

    Private Sub TextBox1_Validated(sender As Object, e As EventArgs) Handles TextBox1.Validated
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

        MainForm.cedControlCalidad = TextBox1.Text

        If modeUsers = 2 Or rd_cedula = Nothing Then
            Dim writer As New StreamWriter(variables.archivoLoginInspector)
            writer.WriteLine(TextBox1.Text)
            writer.WriteLine("Inspector")
            writer.WriteLine(TestEncoding(TextBox2.Text))
            writer.Close()
            MainForm.IndexarToolStripMenuItem.Enabled = True
            MainForm.mode = 1
            Me.Close()
            Exit Sub
        End If

        If modeUsers = 1 Then
            MainForm.mode = 1
        End If

        If TestDecoding(rd_clave) = TextBox2.Text.Trim And rd_cedula = TextBox1.Text Then
            MainForm.IndexarToolStripMenuItem.Enabled = True
            Me.Close()
        Else
            MsgBox("Datos incorrectos!", MsgBoxStyle.OkOnly)
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

    Private Sub UserEntry_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Console.WriteLine(TestEncoding("699"))

        If Not System.IO.File.Exists(variables.ruta(0).ToString + variables.archivotext1) Then
            Dim creartxt As System.IO.FileStream
            creartxt = File.Create(variables.ruta(0).ToString + variables.archivotext1)
            creartxt.Close()

            Dim folderPdf As New DirectoryInfo(variables.ruta(0) + "\pdf")
            Dim file1 As New System.IO.StreamWriter(variables.ruta(0).ToString + variables.archivotext1)

            For Each folder As FileInfo In folderPdf.GetFiles
                file1.WriteLine(TestEncoding(folder.Name.ToString) + "|" + "|" + "|" + "|" + "|")
            Next
            'PERMITE INSERTAR CONTENIDO A UN ARCHIVO
            file1.Close()
        End If

        If Not System.IO.File.Exists(variables.ruta(0).ToString + variables.archivoGuia) Then
            Dim crearGuia As FileStream
            crearGuia = File.Create(variables.ruta(0).ToString + variables.archivoGuia)
            crearGuia.Close()

            Dim escribir As New StreamWriter(variables.ruta(0).ToString + variables.archivoGuia)
            escribir.WriteLine(TestEncoding("0"))
            escribir.Close()
        End If

        If File.Exists(variables.ruta(0).ToString + variables.archivoEstadisticas) Then
            modeLabel.Text = "CONTROL DE CALIDAD"
            If Not File.Exists(variables.archivoLoginInspector) Then
                File.Create(variables.archivoLoginInspector).Close()
                modeUsers = 2
                showMsm()
            Else
                loadData(variables.archivoLoginInspector)
                If rd_cedula = Nothing Then
                    showMsm()
                End If
                modeUsers = 1
            End If

            loadStatistics(variables.ruta(0).ToString + variables.archivoEstadisticas)
            loadPercentage()
        Else
            modeLabel.Text = "INDEXADOR"
            loadData("UConfig.txt")
        End If
        TextBox1.Focus()
    End Sub


    Public Function showMsm()
        MsgBox("LOS PDF HAN SIDO INDEXADOS")
        MsgBox("AHORA EL SISTEMA VA A PASAR AL MODO CONTROL DE CALIDAD, DEBES INGRESAR TU ID DE USUARIO Y UNA CONTRASEÑA PARA EL NUEVO INGRESO")
    End Function

    Public Function loadData(ByVal path As String)
        ' Open the file using a stream reader.
        Using sr As New StreamReader(path)
            ' Read the stream to a string and write the string to the console.
            rd_cedula = sr.ReadLine()
            rd_nombre = sr.ReadLine()
            rd_clave = sr.ReadLine()
            'MsgBox(rd_cedula & rd_nombre & rd_clave, vbOKOnly)
        End Using
    End Function

    Public Sub loadStatistics(ByVal path As String)

        Dim lector As New StreamReader(path, Encoding.Default)

        Dim linea As String = lector.ReadLine()

        While Not linea Is Nothing
            Dim arrayString() As String = linea.Split("|")
            Dim id As String = TestDecoding(arrayString(0))

            If Not MainForm.listDocCheck.IndexOf(id) > -1 Then
                MainForm.listDocCheck.Add(id)
            End If
            linea = lector.ReadLine()
        End While

        lector.Close()

    End Sub

    Public Sub loadPercentage()
        Dim lineas() As String = File.ReadAllLines("UConfig.txt")
        'Console.WriteLine(lineas(lineas.Length - 1))
        MainForm.porcentajeRevision = Convert.ToDecimal(TestDecoding(lineas(lineas.Length - 2))) / 100

    End Sub

End Class