Imports System.IO
Public Class frmIndexacion
    Private rutapdf As String
    Public variables As New VariablesGlobalesYfunciones()
    Public seleccion As Integer
    Public ModoEdit As Boolean = False

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load


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
                DataGridView1.Rows.Add(arrayPdfNombre)
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
            crearComparecientes()
            limpiar1()
            limpiar2()
            ModoEdit = False
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

                    crearComparecientes()
                    variables.MarcarFilaActual()
                    variables.cambiarPosicion(1)
                    limpiar1()
                    limpiar2()
                End If

            Else
                MsgBox("No hay mas pdf")
            End If
        End If

    End Function

    Private Sub frmIndexacion_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        AxAcroPDF1.src = "c:\pdf\20170105-02-0028.pdf"
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
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

        'MsgBox(DataGridView1(0, variables.obtenerPosicionFila).Value.ToString)

    End Sub

    Public Function crearComparecientes()
        If (DataGridView2.Rows.Count > 1) Then
            If (Not System.IO.File.Exists(variables.ruta(0).ToString + variables.archivoCompareciente)) Then
                Dim creararchivo As FileStream
                creararchivo = File.Create(variables.ruta(0).ToString + variables.archivoCompareciente)
                creararchivo.Close()
            End If
            MsgBox(DataGridView2.Rows.Count.ToString)

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
                    If (vectorLinea(0).Equals(DataGridView2(0, 0).Value.ToString)) Then
                        ''MsgBox(vectorLinea(0).ToString)
                    Else
                        escribirCompaTemp.WriteLine(linea)
                    End If
                    ''DataGridView1.Rows.Insert(0, New String() {linea})
                End If
            Loop Until linea Is Nothing

            lectorCompareciente.Close()

            For index = 0 To DataGridView2.Rows.Count - 2
                escribirCompaTemp.WriteLine(DataGridView2(0, index).Value.ToString + "|" + DataGridView2(1, index).Value.ToString + "|" + DataGridView2(2, index).Value.ToString + "|" + DataGridView2(3, index).Value.ToString + "|" + DataGridView2(4, index).Value.ToString)
            Next
            escribirCompaTemp.Close()

            File.Delete(variables.ruta(0) + variables.archivoCompareciente)
            File.Move(archivoTemp, variables.ruta(0) + variables.archivoCompareciente)

        End If


    End Function

    Public Function limpiar1()
        Me.TextBox2.Clear()
        Me.TextBox3.Clear()
        Me.ComboBox1.Text = ""
        Me.ComboBox2.Text = ""
        DateTimePicker1.Value = Now
        DataGridView2.Rows.Clear()
    End Function

    Public Function limpiar2()
        Me.ComboBox3.Text = ""
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        Me.TextBox7.Clear()
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        guardarMetadatosPdf()
    End Sub

    Public Function guardarMetadatosPdf()
        Dim txtPDFTemp As String = variables.ruta(0).ToString + "\pdfTemp.txt"

        Dim creartxt As FileStream
        creartxt = File.Create(txtPDFTemp)
        creartxt.Close()

        Dim escritorPDFMetadatos As New StreamWriter(txtPDFTemp)


        For index = 0 To DataGridView1.Rows.Count - 2
            Dim linea As String = DataGridView1(0, index).Value.ToString + "|" + DataGridView1(1, index).Value.ToString + "|" + DataGridView1(2, index).Value.ToString + "|" + DataGridView1(3, index).Value.ToString + "|" + DataGridView1(4, index).Value.ToString + "|" + DataGridView1(5, index).Value.ToString
            escritorPDFMetadatos.WriteLine(linea)
        Next

        escritorPDFMetadatos.Close()

        File.Delete(variables.ruta(0) + variables.archivotext1)
        File.Move(txtPDFTemp, variables.ruta(0).ToString + variables.archivotext1)

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        limpiar2()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex < variables.obtenerPosicionFila Then
            seleccion = e.RowIndex
            ModoEdit = True
            DataGridView2.Rows.Clear()
            TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
            ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
            TextBox3.Text = DataGridView1(3, seleccion).Value.ToString
            DateTimePicker1.Value = DataGridView1(4, seleccion).Value
            ComboBox2.Text = DataGridView1(5, seleccion).Value.ToString

            Dim id As String = DataGridView1(0, seleccion).Value.ToString

            If File.Exists(variables.ruta(0).ToString + variables.archivoCompareciente) Then
                Dim linea As String
                Dim lectorCompareciente As New StreamReader(variables.ruta(0).ToString + variables.archivoCompareciente)

                Do
                    linea = lectorCompareciente.ReadLine()
                    If linea IsNot Nothing Then

                        Dim arrayLinea As String() = linea.Split("|")
                        If arrayLinea(0).ToString.Equals(id) Then
                            DataGridView2.Rows.Add(arrayLinea)
                        End If
                    End If
                Loop Until linea Is Nothing

                lectorCompareciente.Close()

            End If

        Else
            If e.RowIndex = variables.obtenerPosicionFila() Then

            Else
                MsgBox("No se puede mover a esta fila por que el puntero esta en la fila  " + (variables.obtenerPosicionFila + 1).ToString + ".Solamente puede moverve en filas anteriores")
            End If

        End If
    End Sub
End Class