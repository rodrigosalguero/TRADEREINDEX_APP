Imports System.IO
Public Class frmIndexacion
    Private rutapdf As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim variables As New VariablesGlobalesYfunciones 'creamos un objeto de la clase
        variables.ruta(0) = "g:"

        rutapdf = variables.ruta(0).ToString + "/pdf/"

        'For index = 0 To variables.columnas1.Count
        '    DataGridView1.Columns.Add(index, variables.columnas1(index).ToString)
        'Next

        DataGridView1.Columns.Add(0, "id")
        DataGridView1.Columns.Add(1, "Repertorio")
        DataGridView1.Columns.Add(2, "Libro Registral")
        DataGridView1.Columns.Add(3, "Nº Inscripcion")
        DataGridView1.Columns.Add(4, "Fecha")
        DataGridView1.Columns.Add(5, "Parroquia")

        'Dim grid As DataGridView = variables.crearColumna(DataGridView1, variables.columnas1)
        Dim pdftxt As New StreamReader(variables.ruta(0) + variables.archivotext1)

        Dim linea As String

        Do
            linea = pdftxt.ReadLine()
            If linea IsNot Nothing Then
                DataGridView1.Rows.Insert(0, New String() {linea})
            End If
        Loop Until linea Is Nothing

        Dim leerArchivoGuia As New StreamReader(variables.ruta(0).ToArray + variables.archivoGuia)
        MsgBox(rutapdf + DataGridView1.Rows(leerArchivoGuia.ReadLine().ToString).Cells(0).Value.ToString)

        DataGridView1.Item(1, 0).Value = "hola"


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        siguiente()
    End Sub

    Public Function siguiente()

    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        AxAcroPDF1.src = "g:\pdf\20170105-02-0028.pdf"
    End Sub

    Private Sub frmIndexacion_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        AxAcroPDF1.src = "g:\pdf\20170105-02-0028.pdf"
    End Sub
End Class