Public Class VariablesGlobalesYfunciones
    Public archivotext1 As String = "\pdfNombres.txt"
    Public archivoGuia As String = "\posiciongrid.txt"
    '    Public ruta As String() = Application.StartupPath.ToString.Split("\")
    Public ruta As String() = Application.StartupPath.ToString.Split("\")
    Public columnas1 As String() = {"id", "Repertorio", "Libro Registral", "Nº Inscripcion", "Fecha", "Parroquia"}

    Public Function crearColumna(ByVal grid As DataGridView, ParamArray columnas() As String)
        For index = 0 To columnas.Count
            grid.Columns.Add(index, columnas(index).ToString)
        Next
        Return grid
    End Function
End Class
