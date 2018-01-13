Public Class FillDocsAcept
    Public Shared Sub Acept(ByVal grid As DataGridView)

        Dim dateNow As String = DateTime.Now().ToString()

        For index = 0 To grid.Rows.Count - 1
            grid.Item(20, index).Value = dateNow
            grid.Item(21, index).Value = "TRUE"
        Next

    End Sub
End Class
