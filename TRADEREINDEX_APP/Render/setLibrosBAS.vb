Public Class SetLibrosBAS
    Dim idLibros As List(Of Integer) = New List(Of Integer)(New Integer() {2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21})
    Dim nameLibros As List(Of String) = New List(Of String)(New String() {
            "DEMANDAS",
            "EMBARGOS",
            "ESPECIAL DE ADJUDICACIONES",
            "EXCLUSIÓN DE BIENES",
            "ESPECIAL DE FIANZA PERSONAL",
            "FUNDACIONES",
            "HIPOTECAS Y GRAVÁMENES",
            "INSTRUMENTO PÚBLICO",
            "INTERDICCIONES",
            "PROHIBICIONES JUDICIALES Y LEGALES",
            "MINERO",
            "NEGATIVAS",
            "ORGANIZACIONES RELIGIOSAS",
            "PROPIEDADES HORIZONTALES",
            "PROPIEDADES",
            "ESPECIAL DE SENTENCIAS",
            "SEPARACIONES",
            "PATRIMONIO FAMILIAR",
            "PROHIBICIONES VOLUNTARIAS"})























    Public Function getNameLibro(ByVal id As Integer) As String

        Dim indice As Integer = idLibros.IndexOf(id)

        Return nameLibros(indice)

    End Function


End Class
