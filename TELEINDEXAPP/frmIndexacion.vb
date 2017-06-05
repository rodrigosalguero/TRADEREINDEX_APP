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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim cultura As New CultureInfo("es-ES")
        microfono = New SpeechRecognitionEngine(cultura)
        Dim comandos As String() = {"cedula", "nombres", "apellido", "compareciente", "repertorio", "libro registral", "numero de inscripcion", "inscripcion", "parroquia", "fecha", "siguiente", "anterior", "agregar", "borrar"}
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

        DataGridView1.Columns(0).SortMode = SortOrder.Descending

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
    End Sub


    Public Sub RECONOCE(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
        Dim resultado As RecognitionResult
        resultado = e.Result
        Dim palabra As String = resultado.Text

        Select Case palabra
            Case "fecha"
                DateTimePicker1.Focus()

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
                agregar()
                If elemento IsNot Nothing Then
                    elemento.BackColor = Color.White
                    Label1.Focus()
                End If
            Case "siguiente"
                siguiente()
            Case "repertorio"
                TextBox2.Focus()
                If elemento IsNot Nothing Then
                    elemento.BackColor = Color.White
                    elemento = TextBox2
                    elemento.BackColor = Color.FromName("Highlight")
                End If
            Case "libro registral"
                ComboBox1.Focus()
                If elemento IsNot Nothing Then
                    elemento.BackColor = Color.White
                    elemento = ComboBox1
                    ComboBox1.BackColor = Color.FromName("Highlight")
                End If
        End Select

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
                    Label5.BackColor = Color.Transparent
                    Label5.ForeColor = Color.Black

                    DataGridView1.Item(1, variables.obtenerPosicionFila()).Value = TextBox2.Text
                    DataGridView1.Item(2, variables.obtenerPosicionFila()).Value = ComboBox1.SelectedItem.ToString
                    DataGridView1.Item(3, variables.obtenerPosicionFila()).Value = TextBox3.Text
                    DataGridView1.Item(4, variables.obtenerPosicionFila()).Value = DateTimePicker1.Value.ToString
                    DataGridView1.Item(5, variables.obtenerPosicionFila()).Value = ComboBox2.SelectedItem.ToString

                    Dim NombrePdf As String
                    Dim NombrePdfOld As String = DataGridView1(0, variables.obtenerPosicionFila()).Value
                    NOmbrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))

                    File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
                    DataGridView1(0, variables.obtenerPosicionFila()).Value = NombrePdf + ".pdf"
                    crearComparecientes()
                    guardarMetadatosPdf()
                    variables.MarcarFilaActual()
                    variables.cambiarPosicion(1)
                    limpiar1()
                    limpiar2()
                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                End If

            Else
                MsgBox("No hay mas pdf")
            End If
        End If
    End Function

    Private Sub frmIndexacion_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        AxAcroPDF1.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
        'AxAcroPDF1.gotoLastPage()

    End Sub

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

        If (ModoEdit) Then
            If ComboBox3.SelectedIndex = -1 Then
                Label1.BackColor = Color.Red
                Label1.ForeColor = Color.White
                MsgBox("debe seleccionar un compareciente")
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
                MsgBox("debe seleccionar un compareciente")
            Else
                DataGridView2.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, ComboBox3.SelectedItem.ToString, TextBox4.Text, TextBox5.Text, TextBox7.Text)
                Label1.BackColor = Color.Transparent
                Label1.ForeColor = Color.Black
            End If
            limpiar2()
        End If
    End Function

    Public Function crearComparecientes()
        If (DataGridView2.Rows.Count > 1) Then
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

        ''Dim opcion As Integer = MsgBox("Seguro que deseas borrar estos datos", MsgBoxStyle.OkCancel)

        ''If opcion = 1 Then

        ''limpiar2()
        ''Else

        ''End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
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
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
            Else
                MsgBox("No se puede mover a esta fila por que el puntero esta en la fila  " + (variables.obtenerPosicionFila + 1).ToString + ".Solamente puede moverve en filas anteriores")
            End If

        End If
    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        'seleccion2 = e.RowIndex
        'If e.RowIndex = -1 Then
        'Else
        '    ComboBox3.SelectedItem = DataGridView2(1, e.RowIndex).Value
        '    TextBox4.Text = DataGridView2(2, e.RowIndex).Value.ToString
        '    TextBox5.Text = DataGridView2(3, e.RowIndex).Value.ToString
        '    TextBox7.Text = DataGridView2(4, e.RowIndex).Value.ToString
        '    Button4.Enabled = True
        'End If
    End Sub
End Class