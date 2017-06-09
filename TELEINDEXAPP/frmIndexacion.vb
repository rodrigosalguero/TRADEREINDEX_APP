Imports System.Globalization
Imports System.IO
Imports System.Speech.Recognition
Imports System.Speech.Synthesis


Public Class frmIndexacion
    Private rutapdf As String
    Public variables As New VariablesGlobalesYfunciones()
    Public seleccion As Integer = -1
    Public ModoEdit As Boolean = False
    Dim elemento As New Control(Nothing)
    Dim microfono As New SpeechRecognitionEngine
    Dim digitosRepertorio As Integer = 4
    Dim cryp As New Simple3Des("123456")
    Dim seleccion2 As Integer = -1
    Dim microactive As Boolean = False
    Dim lectorPdf As iTextSharp.text.pdf.PdfReader

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        DataGridView1.Height = Panel6.Height - 50
        Button5.Width = Panel6.Width - 11
        Button5.Height = 30
        Button5.Location = New Point(4, DataGridView1.Height + 10)
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

        Label3.Location = New Point(10, Label3.Location.Y)
        Label5.Location = New Point(10, Label5.Location.Y)
        Label6.Location = New Point(10, Label6.Location.Y)
        Label2.Location = New Point(10, Label2.Location.Y)
        Label9.Location = New Point(10, Label9.Location.Y)

        TextBox2.Location = New Point(96 + 20, TextBox2.Location.Y)
        TextBox2.Width = Panel7.Width - (96 + 30)

        ComboBox1.Location = New Point(96 + 20, ComboBox1.Location.Y)
        ComboBox1.Width = Panel7.Width - (96 + 30)

        TextBox3.Location = New Point(96 + 20, TextBox3.Location.Y)
        TextBox3.Width = Panel7.Width - (96 + 30)

        DateTimePicker1.Location = New Point(96 + 20, DateTimePicker1.Location.Y)
        DateTimePicker1.Width = Panel7.Width - (96 + 30)

        ComboBox2.Location = New Point(96 + 20, ComboBox2.Location.Y)
        ComboBox2.Width = Panel7.Width - (96 + 30)

        Button1.Location = New Point(Panel4.Width - (Button1.Width + 30), Button1.Location.Y)
        Button3.Location = New Point(Panel4.Width - (Button3.Width + Button1.Width + 60), Button1.Location.Y)
        Panel3.Location = New Point(Panel4.Width - (Panel3.Width + Button3.Width + Button1.Width + 90), Panel3.Location.Y)


        microfono = New SpeechRecognitionEngine()
        Dim comandos As String() = {"cedula", "nombres", "apellido", "compareciente", "repertorio", "libro", "inscripcion", "parroquia", "fecha", "siguiente", "agregar", "borrar", "actualizar", "cancelar"}
        Dim VOCABULARIO As New GrammarBuilder
        VOCABULARIO.Append(New Choices(comandos))
        microfono.LoadGrammar(New Grammar(VOCABULARIO))

        microfono.SetInputToDefaultAudioDevice()
        microfono.RecognizeAsync(RecognizeMode.Multiple)
        AddHandler microfono.SpeechRecognized, AddressOf RECONOCE
        AddHandler microfono.SpeechRecognitionRejected, AddressOf NORECONOCE
        AddHandler microfono.SpeechDetected, AddressOf DETECTA


        rutapdf = variables.ruta(0).ToString + "/pdf/"

        'For index = 0 To variables.columnas1.Count
        '    DataGridView1.Columns.Add(index, variables.columnas1(index).ToString)
        'Next

        ''SE CREAN LAS COLUMNAS DEL LIBRO 
        DataGridView1.Columns.Add(0, "id")
        DataGridView1.Columns.Add(1, "Repertorio")
        DataGridView1.Columns.Add(2, "Libro Registral")
        DataGridView1.Columns.Add(3, "Nº Inscripcion")
        DataGridView1.Columns.Add(4, "Fecha")
        DataGridView1.Columns.Add(5, "Parroquia")

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
        'DataGridView1.FirstDisplayedScrollingRowIndex = 0
        DataGridView1.CurrentCell = DataGridView1.Rows(variables.obtenerPosicionFila()).Cells(0)
        DataGridView1.Rows(variables.obtenerPosicionFila()).DefaultCellStyle.BackColor = Color.FromName("Highlight")
        lblFin.Text = DataGridView1.RowCount.ToString
        lblInicio.Text = (variables.obtenerPosicionFila() + 1).ToString
    End Sub
    Public Sub RECONOCE(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
        If microactive Then
            Dim resultado As RecognitionResult
            resultado = e.Result
            Dim palabra As String = resultado.Text

            Select Case palabra
                Case "fecha"
                    DateTimePicker1.Focus()
                    elemento.BackColor = Color.White
                Case "nombres"
                    TextBox5.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = TextBox5
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
                Case "cedula"
                    TextBox4.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = TextBox4
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
                Case "apellido"
                    TextBox7.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = TextBox7
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
                Case "compareciente"
                    ComboBox3.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = ComboBox3
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
                Case "agregar"
                    If Button2.Text = "Agregar" Then
                        agregar()
                        If elemento IsNot Nothing Then
                            elemento.BackColor = Color.White
                            Label1.Focus()
                        End If
                    End If
                Case "siguiente"
                    siguiente()
                Case "borrar"
                    If Button4.Enabled Then
                        Button4_Click(Nothing, Nothing)
                    End If
                Case "cancelar"
                    If Button2.Text = "Cancelar" Then
                        Button2_Click_1(Nothing, Nothing)
                    End If
                Case "actualizar"
                    If Button2.Text = "Actualizar" Then
                        Button2_Click_1(Nothing, Nothing)
                    End If
                Case "repertorio"
                    TextBox2.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = TextBox2
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
                Case "libro"
                    ComboBox1.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = ComboBox1
                        ComboBox1.BackColor = Color.FromName("Highlight")
                    End If
                Case "inscripcion"
                    TextBox3.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = TextBox3
                        elemento.BackColor = Color.FromName("Highlight")
                    End If

                Case "parroquia"
                    ComboBox2.Focus()
                    If elemento IsNot Nothing Then
                        elemento.BackColor = Color.White
                        elemento = ComboBox2
                        elemento.BackColor = Color.FromName("Highlight")
                    End If
            End Select
        End If
    End Sub

    Public Sub DETECTA()

    End Sub

    Public Sub NORECONOCE()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        siguiente()
    End Sub

    Public Function siguiente()
        If (ModoEdit) Then
            DataGridView1.Item(1, seleccion).Value = TextBox2.Text
            DataGridView1.Item(2, seleccion).Value = ComboBox1.SelectedItem.ToString
            DataGridView1.Item(3, seleccion).Value = TextBox3.Text
            DataGridView1.Item(4, seleccion).Value = DateTimePicker1.Value.ToString
            DataGridView1.Item(5, seleccion).Value = ComboBox2.SelectedItem.ToString
            Button2.Text = "Agregar"
            Button4.Enabled = False
            Dim NombrePdf As String
            Dim NombrePdfOld As String = DataGridView1(0, seleccion).Value
            NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
            File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
            DataGridView1(0, seleccion).Value = NombrePdf + ".pdf"
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
                TextBox3.Text = DataGridView1(3, seleccion).Value.ToString
                DateTimePicker1.Value = DataGridView1(4, seleccion).Value
                ComboBox2.Text = DataGridView1(5, seleccion).Value.ToString
                If File.Exists(variables.ruta(0).ToString + variables.archivoCompareciente) Then
                    Dim linea As String
                    Dim lectorCompareciente As New StreamReader(variables.ruta(0).ToString + variables.archivoCompareciente)
                    Do
                        linea = lectorCompareciente.ReadLine()
                        If linea IsNot Nothing Then
                            Dim arrayLinea As String() = linea.Split("|")
                            Dim datoId As String = cryp.DecryptData(arrayLinea(0))
                            If datoId.ToString.Equals(DataGridView1(0, seleccion).Value.ToString) Then
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
            Dim errores As Boolean = False
            If variables.obtenerPosicionFila() < DataGridView1.RowCount - 1 Then
                If TextBox2.Text.Trim = "" Then
                    Label3.BackColor = Color.Red
                    Label3.ForeColor = Color.White
                    errores = True
                Else
                    Label3.BackColor = Color.Transparent
                    Label3.ForeColor = Color.Black
                    errores = False
                End If

                If TextBox3.Text.Trim = "" Then
                    Label6.BackColor = Color.Red
                    Label6.ForeColor = Color.White
                    errores = True
                Else
                    Label6.BackColor = Color.Transparent
                    Label6.ForeColor = Color.Black
                    errores = False
                End If

                If (ComboBox1.SelectedIndex = -1) Then
                    Label5.BackColor = Color.Red
                    Label5.ForeColor = Color.White
                    errores = True
                Else
                    Label5.BackColor = Color.Transparent
                    Label5.ForeColor = Color.Black
                    errores = False
                End If

                If (ComboBox2.SelectedIndex = -1) Then
                    Label9.BackColor = Color.Red
                    Label9.ForeColor = Color.White
                    errores = True
                Else
                    Label9.BackColor = Color.Transparent
                    Label9.ForeColor = Color.Black
                    errores = False
                End If

                If errores Then
                    MsgBox("Los campos Marcados son Obligatorios")
                Else
                    Button2.Text = "Agregar"
                    Button4.Enabled = False
                    Label5.BackColor = Color.Transparent
                    Label5.ForeColor = Color.Black
                    DataGridView1.Item(1, variables.obtenerPosicionFila()).Value = TextBox2.Text
                    DataGridView1.Item(2, variables.obtenerPosicionFila()).Value = ComboBox1.SelectedItem.ToString
                    DataGridView1.Item(3, variables.obtenerPosicionFila()).Value = TextBox3.Text
                    DataGridView1.Item(4, variables.obtenerPosicionFila()).Value = DateTimePicker1.Value.ToString
                    DataGridView1.Item(5, variables.obtenerPosicionFila()).Value = ComboBox2.SelectedItem.ToString
                    Dim NombrePdf As String
                    Dim NombrePdfOld As String = DataGridView1(0, variables.obtenerPosicionFila()).Value
                    NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
                    File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
                    DataGridView1(0, variables.obtenerPosicionFila()).Value = NombrePdf + ".pdf"
                    crearComparecientes()
                    guardarMetadatosPdf()
                    variables.MarcarFilaActual()
                    variables.cambiarPosicion(1)
                    lblInicio.Text = (1 + variables.obtenerPosicionFila()).ToString
                    limpiar1()
                    limpiar2()
                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                End If
            Else
                MsgBox("No hay mas pdf")
            End If
        End If
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
        Me.ComboBox1.SelectedIndex = -1
        Me.ComboBox2.SelectedIndex = -1
        DateTimePicker1.Value = Now
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
                TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
                ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
                TextBox3.Text = DataGridView1(3, seleccion).Value.ToString
                DateTimePicker1.Value = DataGridView1(4, seleccion).Value
                ComboBox2.Text = DataGridView1(5, seleccion).Value.ToString
                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(0)
                Dim id As String = DataGridView1(0, seleccion).Value.ToString
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + id
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

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

    End Sub

    Private Sub DateTimePicker1_LostFocus(sender As Object, e As EventArgs) Handles DateTimePicker1.LostFocus
        DateTimePicker1.Format = DateTimePickerFormat.Long

    End Sub

    Private Sub DateTimePicker1_GotFocus(sender As Object, e As EventArgs) Handles DateTimePicker1.GotFocus

        DateTimePicker1.Focus()
    End Sub

    Private Sub DateTimePicker1_Enter(sender As Object, e As EventArgs) Handles DateTimePicker1.Enter
        DateTimePicker1.Format = DateTimePickerFormat.Short
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub AxAcroPDF1_Enter(sender As Object, e As EventArgs) Handles AxAcroPDF1.Enter
        AxAcroPDF1.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
        ''lectorPdf = New iTextSharp.text.pdf.PdfReader(variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString)
        ''MsgBox(lectorPdf.NumberOfPages)
    End Sub
End Class