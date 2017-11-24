Imports System.Globalization
Imports System.IO
Imports System.Speech.Recognition
Imports System.Speech.Synthesis
Imports System.Text
Imports System.Threading
Imports TELEINDEXAPP.validedDecimal

Public Class frmIndexacion
    Private rutapdf As String
    Public variables As New VariablesGlobalesYfunciones()
    Public ComboInit As New FillCombo()
    Public seleccion As Integer = -1
    Public ModoEdit As Boolean = False
    Dim elemento As New Control(Nothing)
    Dim microfono As New SpeechRecognitionEngine
    Dim digitosRepertorio As Integer = 4
    Dim cryp As New Simple3Des("123456")
    Dim seleccion2 As Integer = -1
    Dim microactive As Boolean = False
    Dim lectorPdf As iTextSharp.text.pdf.PdfReader
    Dim automplete As New AutocompleteItems()

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboInit.Fill(ComboBox2, variables.rutaPath + "Parroquia.txt")
        ComboInit.Fill(ComboBox1, variables.rutaPath + "LibroRegistral.txt")
        ComboInit.Fill(ComboBox4, variables.rutaPath + "DescBien.txt")
        ComboInit.Fill(ComboBox5, variables.rutaPath + "TipoContrato.txt")
        ComboInit.Fill(ComboBox6, variables.rutaPath + "Notarias.txt")
        ''ComboInit.Fill(ComboBox3, variables.rutaPath + "Comparecientes.txt")
        ''MsgBox(MainForm.Size.Width.ToString)
        Dim columna1 As Double = 0.4
        Dim columna2 As Double = 0.6
        Dim pantalla As Size
        pantalla = System.Windows.Forms.SystemInformation.PrimaryMonitorSize

        Me.Width = pantalla.Width + 10
        Me.Height = pantalla.Height - 50
        Panel1.Width = (pantalla.Width - 20) * columna2
        Panel1.Height = (pantalla.Height) - 90

        Panel2.Width = pantalla.Width * columna1
        Panel6.Width = pantalla.Width * columna1
        Panel4.Width = pantalla.Width * columna1
        Panel1.Location = New Point(Panel2.Location.X + Panel2.Width, 0)

        Dim sobrante As Integer = (pantalla.Height - 20) - (Panel4.Height + Panel6.Height + Panel2.Height)
        Panel6.Height = sobrante + 50

        DataGridView1.Width = Panel6.Width - 11
        DataGridView1.Height = Panel6.Height - 10 ''- 50
        'Button5.Width = Panel6.Width - 11
        'Button5.Height = 30
        'Button5.Location = New Point(4, DataGridView1.Height + 10)
        Panel2.Location = New Point(4, sobrante + 110)
        Panel7.Width = Panel2.Width - 10
        Panel5.Width = Panel2.Width - 10
        ''ORDENAR LOS ELEMENTOS DEL PANEL 5
        ''SE DEFINEN 2 COLUMNAS CON UN ESPACIO ENTRE CADA COLUMNA
        Dim espacioentre As Integer = 30
        Dim inicio As Integer = 10
        Dim anchoPanel As Integer = Panel5.Width - (inicio * 2)

        Dim tamanioColuma As Integer = anchoPanel / 2 - (espacioentre / 2)
        Label1.Location = New Point(inicio, 0)
        ComboBox3.Width = tamanioColuma
        ComboBox3.Location = New Point(inicio, Label1.Height + 3)


        Label8.Location = New Point((ComboBox3.Width + inicio) + espacioentre, 0)
        TextBox4.Width = tamanioColuma
        TextBox4.Location = New Point((ComboBox3.Width + inicio) + espacioentre, Label8.Height + 3)

        Label4.Location = New Point(inicio, (Label1.Height + 3) + ComboBox3.Height + 3)
        TextBox5.Width = tamanioColuma
        TextBox5.Location = New Point(Label4.Location.X, Label4.Location.Y + Label4.Height + 3)

        Label11.Location = New Point(TextBox5.Width + inicio + espacioentre, Label4.Location.Y)
        TextBox7.Width = tamanioColuma
        TextBox7.Location = New Point(Label8.Location.X, Label11.Location.Y + Label11.Height + 3)

        Button2.Width = tamanioColuma
        Button2.Location = New Point(inicio, TextBox5.Location.Y + TextBox5.Height + 3)

        Button4.Width = tamanioColuma
        Button4.Location = New Point(TextBox7.Location.X, TextBox7.Location.Y + TextBox7.Height + 3)

        DataGridView2.Width = Panel5.Width - 20
        DataGridView2.Location = New Point(inicio, Button2.Location.Y + Button2.Height + 5)

        TextBox2.Location = New Point(96 + 20, TextBox2.Location.Y)
        TextBox2.Width = (Panel7.Width / 2) - (96 + 30)

        TextBox6.Location = New Point(96 + 20, TextBox6.Location.Y)
        TextBox6.Width = (Panel7.Width / 2) - (96 + 30)

        ComboBox1.Location = New Point(96 + 20, ComboBox1.Location.Y)
        ComboBox1.Width = Panel7.Width - (96 + 30)

        ComboBox5.Location = New Point(96 + 20, ComboBox5.Location.Y)
        ComboBox5.Width = Panel7.Width - (96 + 30)

        DateTimePicker1.Location = New Point(96 + 20, DateTimePicker1.Location.Y)
        DateTimePicker1.Width = Panel7.Width - (96 + 30)

        ComboBox2.Location = New Point(96 + 20, ComboBox2.Location.Y)
        ComboBox2.Width = Panel7.Width - (96 + 30)

        ComboBox4.Location = New Point(96 + 20, ComboBox4.Location.Y)
        ComboBox4.Width = Panel7.Width - (96 + 30)

        TextBox1.Location = New Point(96 + 20, TextBox1.Location.Y)
        TextBox1.Width = Panel7.Width - (96 + 30)

        Button1.Location = New Point(Panel4.Width - (Button1.Width + 30), Button1.Location.Y)
        Button3.Location = New Point(Panel4.Width - (Button3.Width + Button1.Width + 60), Button1.Location.Y)
        Panel3.Location = New Point(Panel4.Width - (Panel3.Width + Button3.Width + Button1.Width + 90), Panel3.Location.Y)

        Label3.Location = New Point(10, Label3.Location.Y)
        Label5.Location = New Point(10, Label5.Location.Y)
        Label6.Location = New Point(TextBox2.Location.X + TextBox2.Width + 5, Label6.Location.Y)
        Label2.Location = New Point(10, Label2.Location.Y)
        Label9.Location = New Point(10, Label9.Location.Y)
        Label12.Location = New Point(10, Label12.Location.Y)
        Label13.Location = New Point(10, Label13.Location.Y)
        Label14.Location = New Point(10, Label14.Location.Y)
        Label15.Location = New Point(10, Label15.Location.Y)
        Label16.Location = New Point(TextBox6.Location.X + TextBox6.Width + 5, Label16.Location.Y)
        Label10.Location = New Point(10, Label10.Location.Y)

        TextBox3.Location = New Point(Label6.Location.X + Label6.Width + 5, TextBox3.Location.Y)
        TextBox3.Width = (Panel7.Width / 2) - (106)

        DateTimePicker2.Location = New Point(Label16.Location.X + Label16.Width + 5, DateTimePicker2.Location.Y)
        DateTimePicker2.Width = (Panel7.Width / 2) - (108)

        ComboBox6.Location = New Point(96 + 20, ComboBox6.Location.Y)
        ComboBox6.Width = Panel7.Width - (96 + 30)

        'microfono = New SpeechRecognitionEngine()
        'Dim comandos As String() = {"cedula", "nombres", "apellido", "compareciente", "repertorio", "libro", "inscripcion", "parroquia", "fecha", "siguiente", "agregar", "borrar", "actualizar", "cancelar"}
        'Dim VOCABULARIO As New GrammarBuilder
        'VOCABULARIO.Append(New Choices(comandos))
        'microfono.LoadGrammar(New Grammar(VOCABULARIO))

        'microfono.SetInputToDefaultAudioDevice()
        'microfono.RecognizeAsync(RecognizeMode.Multiple)
        'AddHandler microfono.SpeechRecognized, AddressOf RECONOCE
        'AddHandler microfono.SpeechRecognitionRejected, AddressOf NORECONOCE
        'AddHandler microfono.SpeechDetected, AddressOf DETECTA

        rutapdf = variables.ruta(0).ToString + "/pdf/"

        'For index = 0 To variables.columnas1.Count
        '    DataGridView1.Columns.Add(index, variables.columnas1(index).ToString)
        'Next

        ''SE CREAN LAS COLUMNAS DEL LIBRO 
        DataGridView1.Columns.Add(0, "id")
        DataGridView1.Columns.Add(1, "Repertorio")
        DataGridView1.Columns.Add(2, "Libro Registral")
        DataGridView1.Columns.Add(3, "Tipo Contrato")
        DataGridView1.Columns.Add(4, "Nº Inscripcion")
        DataGridView1.Columns.Add(5, "Fecha")
        DataGridView1.Columns.Add(6, "Parroquia")
        DataGridView1.Columns.Add(7, "Descripción del Bien")
        DataGridView1.Columns.Add(8, "Marginaciones")
        DataGridView1.Columns.Add(9, "Cuantía")
        DataGridView1.Columns.Add(10, "Fecha de escritura")
        DataGridView1.Columns.Add(11, "Notaría")

        ''DataGridView1.Columns(6).Visible = False

        DataGridView1.Columns(0).SortMode = SortOrder.Ascending


        ''SE CREAN LAS COLUMNAS DE LOS COMPARECIENTE
        DataGridView2.Columns.Add(0, "id")
        DataGridView2.Columns.Add(1, "Compareciente")
        DataGridView2.Columns.Add(2, "Cédula")
        DataGridView2.Columns.Add(3, "Nombres")
        DataGridView2.Columns.Add(4, "Apellidos")
        'Dim grid As DataGridView = variables.crearColumna(DataGridView1, variables.columnas1)
        Dim pdftxt As New StreamReader(variables.ruta(0) + variables.archivotext1)
        Dim linea As String
        Do
            linea = pdftxt.ReadLine()
            If linea IsNot Nothing Then
                Dim arrayPdfNombre As String() = linea.Split("|")
                Dim array(arrayPdfNombre.Count - 1) As String
                For index = 0 To array.Count - 2
                    If arrayPdfNombre(index).Trim() IsNot "" Then
                        array(index) = cryp.DecryptData(arrayPdfNombre(index))
                    End If
                Next
                DataGridView1.Rows.Add(array)
            End If
        Loop Until linea Is Nothing
        pdftxt.Close()
        'Dim leerArchivoGuia As New StreamReader(variables.ruta(0).ToArray + variables.archivoGuia)
        'MsgBox(rutapdf + DataGridView1.Rows(leerArchivoGuia.ReadLine().ToString).Cells(0).Value.ToString)
        'DataGridView1.Item(1, 0).Value = "hola"
        'DataGridView1.Rows(0).Selected = True

        DataGridView1.FirstDisplayedScrollingRowIndex = 0
        DataGridView1.CurrentCell = DataGridView1.Rows(variables.obtenerPosicionFila()).Cells(0)
        DataGridView1.Rows(variables.obtenerPosicionFila()).DefaultCellStyle.BackColor = Color.FromName("Highlight")
        lblFin.Text = DataGridView1.RowCount.ToString
        lblInicio.Text = (variables.obtenerPosicionFila() + 1).ToString
        automplete.FillControls(TextBox5, "comparecientes.txt", 3, "|")
        automplete.FillControls(TextBox7, "comparecientes.txt", 4, "|")
        loadMetadate()
    End Sub
    'Public Sub RECONOCE(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
    '    If microactive Then
    '        Dim resultado As RecognitionResult
    '        resultado = e.Result
    '        Dim palabra As String = resultado.Text

    '        Select Case palabra
    '            Case "fecha"
    '                DateTimePicker1.Focus()
    '                elemento.BackColor = Color.White
    '                Return
    '            Case "nombres"
    '                TextBox5.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox5
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "cedula"
    '                TextBox4.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox4
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "apellido"
    '                TextBox7.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox7
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "compareciente"
    '                ComboBox3.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox3
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "agregar"
    '                If Button2.Text = "Agregar" Then
    '                    agregar()
    '                    If elemento IsNot Nothing Then
    '                        elemento.BackColor = Color.White
    '                        Label1.Focus()
    '                    End If
    '                End If
    '            Case "siguiente"
    '                siguiente()
    '            Case "borrar"
    '                If Button4.Enabled Then
    '                    Button4_Click(Nothing, Nothing)
    '                End If
    '            Case "cancelar"
    '                If Button2.Text = "Cancelar" Then
    '                    Button2_Click_1(Nothing, Nothing)
    '                End If
    '            Case "actualizar"
    '                If Button2.Text = "Actualizar" Then
    '                    Button2_Click_1(Nothing, Nothing)
    '                End If
    '            Case "repertorio"
    '                TextBox2.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox2
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "libro"
    '                ComboBox1.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox1
    '                    ComboBox1.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "inscripcion"
    '                TextBox3.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox3
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If

    '            Case "parroquia"
    '                ComboBox2.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox2
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case Else
    '                Exit Select
    '        End Select
    '    End If

    'End Sub

    'Public Sub DETECTA()

    'End Sub

    'Public Sub NORECONOCE()

    'End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (String.IsNullOrWhiteSpace(TextBox2.Text) And String.IsNullOrWhiteSpace(TextBox3.Text) And
            String.IsNullOrWhiteSpace(TextBox1.Text) And String.IsNullOrWhiteSpace(TextBox6.Text)) Then
            MsgBox("NO SE PUEDEN DEJAR CAMPOS EN BLANCO. RELLENE CON 'N/A' EN CASO DE QUE NO CONTENGA INFORMACIÓN")
        Else
            siguiente()
        End If
    End Sub

    Public Function siguiente()
        If (ModoEdit) Then
            DataGridView1.Item(1, seleccion).Value = TextBox2.Text
            DataGridView1.Item(2, seleccion).Value = ComboBox1.SelectedItem.ToString
            DataGridView1.Item(3, seleccion).Value = ComboBox5.SelectedItem.ToString
            DataGridView1.Item(4, seleccion).Value = TextBox3.Text
            DataGridView1.Item(5, seleccion).Value = DateTimePicker1.Text.ToString
            DataGridView1.Item(6, seleccion).Value = ComboBox2.SelectedItem.ToString
            DataGridView1.Item(7, seleccion).Value = ComboBox4.SelectedItem.ToString
            DataGridView1.Item(8, seleccion).Value = TextBox1.Text
            DataGridView1.Item(9, seleccion).Value = TextBox6.Text
            DataGridView1.Item(10, seleccion).Value = DateTimePicker2.Text.ToString()
            DataGridView1.Item(11, seleccion).Value = ComboBox6.Text
            Button2.Text = "Agregar"
            Button4.Enabled = False

            'Funcion que permite renombrar los archivos PDF
            'por la nomenclatura dada por el cliente

            'Dim NombrePdf As String
            'Dim NombrePdfOld As String = DataGridView1(0, seleccion).Value
            'NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
            'File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
            'DataGridView1(0, seleccion).Value = NombrePdf + ".pdf"

            crearComparecientes()
            limpiar1()
            limpiar2()
            guardarMetadatosPdf()
            seleccion = seleccion + 1
            lblInicio.Text = (seleccion + 1).ToString
            If (seleccion < variables.obtenerPosicionFila()) Then
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
                ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
                ComboBox5.Text = DataGridView1(3, seleccion).Value.ToString
                TextBox3.Text = DataGridView1(4, seleccion).Value.ToString
                DateTimePicker1.Text = DataGridView1(5, seleccion).Value
                ComboBox2.Text = DataGridView1(6, seleccion).Value.ToString
                ComboBox4.Text = DataGridView1(7, seleccion).Value.ToString
                TextBox1.Text = DataGridView1(8, seleccion).Value.ToString
                TextBox6.Text = DataGridView1.Item(9, seleccion).Value.ToString
                DateTimePicker2.Text = DataGridView1.Item(10, seleccion).Value
                ComboBox6.Text = DataGridView1.Item(11, seleccion).Value.ToString

                FillComparecientes(DataGridView1(0, seleccion).Value.ToString)

                DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.Cyan
                DataGridView1.CurrentCell = DataGridView1.Rows(seleccion).Cells(0)
            Else
                DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                DataGridView1.CurrentCell = DataGridView1.Rows(seleccion).Cells(0)
                ModoEdit = False
                seleccion = -1
            End If
        Else

            If variables.obtenerPosicionFila() < DataGridView1.RowCount Then

                Dim validatedCamps As New ValidatedControlsForms()

                If validatedCamps.check() Then
                    MsgBox("Los campos Marcados son Obligatorios")
                Else
                    Button2.Text = "Agregar"
                    Button4.Enabled = False
                    Label5.BackColor = Color.Transparent
                    Label5.ForeColor = Color.Black
                    DataGridView1.Item(1, variables.obtenerPosicionFila()).Value = TextBox2.Text
                    DataGridView1.Item(2, variables.obtenerPosicionFila()).Value = ComboBox1.SelectedItem.ToString
                    DataGridView1.Item(3, variables.obtenerPosicionFila()).Value = ComboBox5.SelectedItem.ToString
                    DataGridView1.Item(4, variables.obtenerPosicionFila()).Value = TextBox3.Text
                    DataGridView1.Item(5, variables.obtenerPosicionFila()).Value = DateTimePicker1.Text.ToString
                    DataGridView1.Item(6, variables.obtenerPosicionFila()).Value = ComboBox2.SelectedItem.ToString
                    DataGridView1.Item(7, variables.obtenerPosicionFila()).Value = ComboBox4.SelectedItem.ToString
                    DataGridView1.Item(8, variables.obtenerPosicionFila()).Value = TextBox1.Text
                    DataGridView1.Item(9, variables.obtenerPosicionFila()).Value = TextBox6.Text
                    DataGridView1.Item(10, variables.obtenerPosicionFila()).Value = DateTimePicker2.Text.ToString()
                    DataGridView1.Item(11, variables.obtenerPosicionFila()).Value = ComboBox6.Text

                    'funcion que permite cambiar el nombre del pdf por un formato especifico
                    'dado por el cliente

                    'Dim NombrePdf As String
                    'Dim NombrePdfOld As String = DataGridView1(0, variables.obtenerPosicionFila()).Value
                    'NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
                    'File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
                    'DataGridView1(0, variables.obtenerPosicionFila()).Value = NombrePdf + ".pdf"

                    crearComparecientes()
                    guardarMetadatosPdf()
                    variables.MarcarFilaActual()
                    limpiar1()
                    limpiar2()
                    If (variables.obtenerPosicionFila() < DataGridView1.RowCount - 1) Then
                        variables.cambiarPosicion(1)
                        lblInicio.Text = (1 + variables.obtenerPosicionFila()).ToString
                        AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    Else
                        'AQUI SE DEBE MARCAR QUE YA SE FINALIZO TODO
                        MsgBox("NO HAY MAS PDF PARA INDEXAR")
                    End If
                End If
            Else
                MsgBox("No hay mas pdf")
            End If
        End If
        automplete.FillControls(TextBox5, "comparecientes.txt", 3, "|")
        automplete.FillControls(TextBox7, "comparecientes.txt", 4, "|")
    End Function



    Public Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        If elemento IsNot Nothing Then
            elemento.BackColor = Color.White
            Label1.Focus()
        End If
        agregar()
        'MsgBox(DataGridView1(0, variables.obtenerPosicionFila).Value.ToString)

    End Sub

    Public Function agregar()
        ''AxAcroPDF1.gotoLastPage()
        If Button2.Text = "Agregar" Then
            If (ModoEdit) Then
                If ComboBox3.SelectedIndex = -1 Then
                    Label1.BackColor = Color.Red
                    Label1.ForeColor = Color.White
                    DataGridView2.Rows.Add(DataGridView1(0, seleccion).Value.ToString, "", TextBox4.Text, TextBox5.Text, TextBox7.Text)
                Else
                    DataGridView2.Rows.Add(DataGridView1(0, seleccion).Value.ToString, ComboBox3.SelectedItem.ToString, TextBox4.Text, TextBox5.Text, TextBox7.Text)
                    Label1.BackColor = Color.Transparent
                    Label1.ForeColor = Color.Black
                End If
                limpiar2()
            Else
                If ComboBox3.SelectedIndex = -1 Then
                    Label1.BackColor = Color.Red
                    Label1.ForeColor = Color.White
                    DataGridView2.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, "", TextBox4.Text, TextBox5.Text, TextBox7.Text)
                Else
                    DataGridView2.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, ComboBox3.SelectedItem.ToString, TextBox4.Text, TextBox5.Text, TextBox7.Text)
                    Label1.BackColor = Color.Transparent
                    Label1.ForeColor = Color.Black
                End If
                limpiar2()
            End If
        ElseIf Button2.Text = "Cancelar" Then
            limpiar2()
            Button2.Text = "Agregar"
            seleccion2 = -1
            Button4.Enabled = False
        Else
            DataGridView2(1, seleccion2).Value = ComboBox3.SelectedItem.ToString
            DataGridView2(2, seleccion2).Value = TextBox4.Text
            DataGridView2(3, seleccion2).Value = TextBox5.Text
            DataGridView2(4, seleccion2).Value = TextBox7.Text
            seleccion2 = -1
            Button2.Text = "Agregar"
            Button4.Enabled = False
            limpiar2()
        End If
    End Function

    Public Function crearComparecientes()
        If (DataGridView2.Rows.Count > 0) Then
            If (Not System.IO.File.Exists(variables.ruta(0).ToString + variables.archivoCompareciente)) Then
                Dim creararchivo As FileStream
                creararchivo = File.Create(variables.ruta(0).ToString + variables.archivoCompareciente)
                creararchivo.Close()
            End If
            ''MsgBox(DataGridView2.Rows.Count.ToString)

            Dim lectorCompareciente As New StreamReader(variables.ruta(0).ToString + variables.archivoCompareciente)

            Dim archivoTemp As String = variables.ruta(0).ToString + "\comparecientestemp.txt"
            Dim archivotemporal As FileStream
            archivotemporal = File.Create(archivoTemp)
            archivotemporal.Close()

            Dim linea As String
            Dim escribirCompaTemp As New StreamWriter(archivoTemp)
            Do
                linea = lectorCompareciente.ReadLine()
                If linea IsNot Nothing Then
                    Dim vectorLinea As String() = linea.Split("|")

                    Dim datoID As String = cryp.DecryptData(vectorLinea(0))

                    If (datoID.Equals(DataGridView2(0, 0).Value.ToString)) Then
                        ''MsgBox(vectorLinea(0).ToString)
                    Else
                        escribirCompaTemp.WriteLine(linea)
                    End If
                    ''DataGridView1.Rows.Insert(0, New String() {linea})
                End If
            Loop Until linea Is Nothing

            lectorCompareciente.Close()

            For index = 0 To DataGridView2.Rows.Count - 1
                If ModoEdit Then
                    DataGridView2(0, index).Value = DataGridView1(0, seleccion).Value
                Else
                    DataGridView2(0, index).Value = DataGridView1(0, variables.obtenerPosicionFila()).Value
                End If
                escribirCompaTemp.WriteLine(cryp.EncryptData(DataGridView2(0, index).Value) + "|" + cryp.EncryptData(DataGridView2(1, index).Value) + "|" + cryp.EncryptData(DataGridView2(2, index).Value) + "|" + cryp.EncryptData(DataGridView2(3, index).Value) + "|" + cryp.EncryptData(DataGridView2(4, index).Value))
            Next
            escribirCompaTemp.Close()
            File.Delete(variables.ruta(0) + variables.archivoCompareciente)
            File.Move(archivoTemp, variables.ruta(0) + variables.archivoCompareciente)

        End If
    End Function

    Public Function limpiar1()
        Me.TextBox2.Clear()
        Me.TextBox3.Clear()
        Me.TextBox1.Clear()
        Me.TextBox6.Clear()
        Me.ComboBox1.SelectedIndex = -1
        Me.ComboBox2.SelectedIndex = -1
        Me.ComboBox4.SelectedIndex = -1
        Me.ComboBox5.SelectedIndex = -1
        Me.ComboBox6.SelectedIndex = -1
        Me.ComboBox3.Items.Clear()
        DateTimePicker1.Text = Now
        DateTimePicker2.Text = Now
        DataGridView2.Rows.Clear()
    End Function

    Public Function limpiar2()
        Me.ComboBox3.SelectedIndex = -1
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        Me.TextBox7.Clear()
    End Function

    Public Function guardarMetadatosPdf()
        Dim txtPDFTemp As String = variables.ruta(0).ToString + "\pdfTemp.txt"

        Dim creartxt As FileStream
        creartxt = File.Create(txtPDFTemp)
        creartxt.Close()
        Dim escritorPDFMetadatos As New StreamWriter(txtPDFTemp)
        For index = 0 To DataGridView1.Rows.Count - 1
            Dim linea As String = ""
            Dim cryp1 As New Simple3Des("123456")

            For index2 = 0 To DataGridView1.Columns.Count - 1
                Dim dato As String = DataGridView1(index2, index).Value
                If dato = "" Or dato = " " Then
                    linea = linea + "|"
                Else
                    linea = linea + cryp1.EncryptData(dato) + "|"
                End If
            Next
            escritorPDFMetadatos.WriteLine(linea)
        Next

        escritorPDFMetadatos.Close()

        File.Delete(variables.ruta(0) + variables.archivotext1)
        File.Move(txtPDFTemp, variables.ruta(0).ToString + variables.archivotext1)

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim opcion As Integer = MsgBox("Seguro que deseas borrar estos datos", MsgBoxStyle.OkCancel)

        If opcion = 1 Then
            DataGridView2.Rows.RemoveAt(seleccion2)
            limpiar2()
            seleccion2 = -1
            Button4.Enabled = False
            Button2.Text = "Agregar"
        Else

        End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick

        If e.RowIndex <> -1 Then
            lblInicio.Text = (e.RowIndex + 1).ToString
            If e.RowIndex < variables.obtenerPosicionFila Then
                If Not seleccion = -1 Then
                    DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
                End If
                seleccion = e.RowIndex
                ModoEdit = True
                DataGridView2.Rows.Clear()
                fillCambos()
                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(0)
                Dim id As String = DataGridView1(0, seleccion).Value.ToString
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + id

                FillComparecientes(id)

                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Cyan

            Else
                If e.RowIndex = variables.obtenerPosicionFila() Then
                    limpiar1()
                    limpiar2()
                    ModoEdit = False

                    If (seleccion >= 0) Then
                        DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
                    End If

                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                Else
                    MsgBox("No se puede mover a esta fila por que el puntero esta en la fila  " + (variables.obtenerPosicionFila + 1).ToString + ".Solamente puede moverve en filas anteriores")
                End If

            End If
        End If

    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        seleccion2 = e.RowIndex
        If e.RowIndex = -1 Then
        Else
            ComboBox3.SelectedItem = DataGridView2(1, e.RowIndex).Value
            TextBox4.Text = DataGridView2(2, e.RowIndex).Value.ToString
            TextBox5.Text = DataGridView2(3, e.RowIndex).Value.ToString
            TextBox7.Text = DataGridView2(4, e.RowIndex).Value.ToString
            Button4.Enabled = True
            Button2.Text = "Cancelar"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If microactive Then
            Button3.BackgroundImage = Image.FromFile(Application.StartupPath.ToString + "\microdisable.png")
            microactive = False
        Else
            Button3.BackgroundImage = Image.FromFile(Application.StartupPath.ToString + "\microactive.png")
            microactive = True
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        cambiarEstado()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        cambiarEstado()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        cambiarEstado()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        cambiarEstado()
    End Sub

    Function cambiarEstado()
        If seleccion2 > -1 Then
            If Button2.Text <> "Actualizar" Then
                Button2.Text = "Actualizar"
            End If
        End If
    End Function

    Private Sub AxAcroPDF1_Enter(sender As Object, e As EventArgs) Handles AxAcroPDF1.Enter
        AxAcroPDF1.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
        'lectorPdf = New iTextSharp.text.pdf.PdfReader(variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString)
        'MsgBox(lectorPdf.NumberOfPages)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ComboBox3.Items.Clear()
        ComboBox3.SelectedIndex = -1
        ComboBox3.Text = ""
        Dim indice As String = Convert.ToString(ComboBox5.SelectedIndex.ToString)
        Dim reader As New StreamReader(variables.rutaPath + "Comparecientes.txt")
        Dim linea As String
        Do
            linea = reader.ReadLine()
            If (Not linea Is Nothing) Then
                Dim vectorLinea As String() = linea.Split(",")
                If (vectorLinea(0).Equals(indice)) Then
                    ComboBox3.Enabled = True
                    ComboBox3.Items.Add(vectorLinea(1))
                End If
            End If
        Loop Until linea Is Nothing
        ComboBox3.Visible = False
        Task.Delay(200)
        ComboBox3.Visible = True
    End Sub

    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus
        Dim cedula As String = TextBox4.Text
        Dim catalogoRuta As String = variables.ruta(0) + "/catalogos/userData.txt"

        If (File.Exists(catalogoRuta)) Then
            If (Not String.IsNullOrWhiteSpace(cedula)) Then
                Dim reader As New StreamReader(catalogoRuta, Encoding.Default)
                Dim linea As String
                Do
                    linea = reader.ReadLine()
                    If (linea Is Nothing) Then
                        Exit Do
                    End If
                    Dim ArrayLine As String() = linea.Split(",")
                    If (ArrayLine(0).Equals(cedula)) Then
                        TextBox5.Text = ArrayLine(1)
                        TextBox7.Text = ArrayLine(2)
                        reader.Dispose()
                        Exit Do
                    End If
                Loop Until linea Is Nothing
                reader.Dispose()
            End If
        End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress

        TextBox5.Text = ""
        TextBox7.Text = ""
    End Sub

    Private Sub DateTimePicker1_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles DateTimePicker1.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not e.IsValidInput Then
            ToolTip1.ToolTipTitle = "Fecha invalida"
            ToolTip1.Show("La fecha ingresada es incorrecta", DateTimePicker1, 0, -DateTimePicker1.Location.Y + DateTimePicker1.Height - 10)
            DateTimePicker1.Focus()
        Else
            Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

            If userData > DateTime.Now Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", DateTimePicker1, 0, -DateTimePicker1.Location.Y + DateTimePicker1.Height - 10)
                DateTimePicker1.Focus()
            End If

        End If

    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        Dim valid As New validedDecimal()
        If Not valid.validated(e, TextBox6.Text) Then
            e.Handled = True
        End If
    End Sub

    Private Sub DateTimePicker2_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles DateTimePicker2.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not e.IsValidInput Then
            ToolTip1.ToolTipTitle = "Fecha invalida"
            ToolTip1.Show("La fecha ingresada es incorrecta", DateTimePicker2, 0, -73)
            DateTimePicker2.Focus()
        Else
            Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

            If userData > DateTime.Now Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", DateTimePicker2, 0, -73)
                DateTimePicker2.Focus()
            End If

        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        Me.ToolTip1.Hide(Me.DateTimePicker1)
    End Sub

    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        Me.ToolTip1.Hide(Me.DateTimePicker2)
    End Sub

    Private Sub loadMetadate()
        If DataGridView1.Rows.Count = variables.obtenerPosicionFila() + 1 Then
            seleccion = DataGridView1.Rows.Count - 1
            'fillCambos()

        End If
    End Sub
    Private Sub fillCambos()
        TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
        ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
        ComboBox5.Text = DataGridView1(3, seleccion).Value.ToString
        TextBox3.Text = DataGridView1(4, seleccion).Value.ToString
        DateTimePicker1.Text = DataGridView1(5, seleccion).Value
        ComboBox2.Text = DataGridView1(6, seleccion).Value.ToString
        ComboBox4.Text = DataGridView1(7, seleccion).Value.ToString
        TextBox1.Text = DataGridView1(8, seleccion).Value.ToString
        TextBox6.Text = DataGridView1.Item(9, seleccion).Value.ToString
        DateTimePicker2.Text = DataGridView1.Item(10, seleccion).Value
        ComboBox6.Text = DataGridView1.Item(11, seleccion).Value.ToString
    End Sub

    Private Sub FillComparecientes(ByVal id As String)
        If File.Exists(variables.ruta(0).ToString + variables.archivoCompareciente) Then
            Dim linea As String
            Dim lectorCompareciente As New StreamReader(variables.ruta(0).ToString + variables.archivoCompareciente)

            Do
                linea = lectorCompareciente.ReadLine()
                If linea IsNot Nothing Then

                    Dim arrayLinea As String() = linea.Split("|")
                    Dim datoId As String = cryp.DecryptData(arrayLinea(0))
                    If datoId.ToString.Equals(id) Then
                        Dim array(arrayLinea.Count - 1) As String
                        For index = 0 To array.Count - 1
                            If arrayLinea(index).Trim() IsNot "" Then
                                array(index) = cryp.DecryptData(arrayLinea(index))
                            End If
                        Next

                        DataGridView2.Rows.Add(array)
                    End If
                End If
            Loop Until linea Is Nothing
            lectorCompareciente.Close()
        End If
    End Sub

End Class