Imports System.IO
Public Class VariablesGlobalesYfunciones
    Public archivotext1 As String = "\pdfNombres.txt"
    Public archivoGuia As String = "\posiciongrid.txt"
    Public archivoCompareciente As String = "\comparecientes.txt"
    Public archivoActos As String = "\metadatosActos.txt"

    '    Public ruta As String() = Application.StartupPath.ToString.Split("\")
    Public ruta As String() = Application.StartupPath.ToString.Split("\")
    Public columnas1 As String() = {"id", "Repertorio", "Libro Registral", "Nº Inscripcion", "Fecha", "Parroquia"}

    Public Function crearColumna(ByVal grid As DataGridView, ParamArray columnas() As String)
        For index = 0 To columnas.Count
            grid.Columns.Add(index, columnas(index).ToString)
        Next
        Return grid
    End Function

    Function obtenerPosicionFila() As Integer
        Dim lectorArchivo As New StreamReader(Me.ruta(0).ToString + Me.archivoGuia)
        Dim posicion As Integer = Convert.ToInt64(lectorArchivo.ReadLine())
        lectorArchivo.Close()
        Return posicion
    End Function



    Public Function MarcarFilaActual()
        If obtenerPosicionFila() < (frmIndexacion.DataGridView1.Rows.Count - 2) Then
            ''PERMITEN POSICIONAR EL FOCUS EN UNA POSICION DE LA FILA
            frmIndexacion.DataGridView1.CurrentCell = frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(0)
            ''CAMBIA EL COLOR DE TODA UNA FILA PARA DAR A CONOCER QUE ESTA SELECCIONADA
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).DefaultCellStyle.BackColor = Color.FromName("Highlight")

            'REGRESA A ESTADO NORNAL EL COLOR DE FONDO DE LA FILA
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).DefaultCellStyle.BackColor = Color.White

            'CAMBIA EL COLOR DE LAS LETRAS DE LA FILA ACTUAL
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(0).Style.ForeColor = Color.White
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(1).Style.ForeColor = Color.White
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(2).Style.ForeColor = Color.White
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(3).Style.ForeColor = Color.White
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(4).Style.ForeColor = Color.White
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila() + 1).Cells(5).Style.ForeColor = Color.White

            'RESTAURA EL COLOR
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(0).Style.ForeColor = Color.Black
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(1).Style.ForeColor = Color.Black
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(2).Style.ForeColor = Color.Black
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(3).Style.ForeColor = Color.Black
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(4).Style.ForeColor = Color.Black
            frmIndexacion.DataGridView1.Rows(obtenerPosicionFila()).Cells(5).Style.ForeColor = Color.Black
        End If
    End Function

    Function cambiarPosicion(ByVal valor As Integer)
        Dim posicion As Integer = obtenerPosicionFila()
        posicion = posicion + valor

        File.Delete(ruta(0).ToString + archivoGuia)

        Dim creartxt As FileStream
        creartxt = File.Create(ruta(0).ToString + archivoGuia)
        creartxt.Close()

        Dim escribirtxt As New StreamWriter(ruta(0).ToString + archivoGuia)

        escribirtxt.WriteLine(posicion)
        escribirtxt.Close()

    End Function

End Class
