Public Class setLibros
    Dim idLibros As List(Of Integer) = New List(Of Integer)(New Integer() {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 50})
    Dim nameLibros As List(Of String) = New List(Of String)(New String() {"PROPIEDADES",
            "HIPOTECAS",
            "PROHIBICIONES",
            "DEMANDAS",
            "EMBARGOS",
            "PATRIMONIO FAMILIAR",
            "PRENDA AGRICOLA",
            "FIANZAS PERSONALES",
            "GRAVAMENES",
            "FIDEICOMISO MERCANTIL",
            "PROPIEDAD HORIZONTAL",
            "ARRENDAMIENTOS",
            "ENCARGO FIDUCIARIO",
            "POSESIONES EFECTIVAS",
            "DISOLUCION CONYUGAL",
            "CESION DE DERECHOS DE AGUA",
            "PROMESA DE COMPRA VENTA",
            "MINERIA",
            "EXCLUSIONES",
            "ORGANIZACIONES RELIGIOSAS",
            "ORGANIZACIONES CATOLICAS",
            "MANDATO ESPECIAL DE MINERIA",
            "CONCURSO PREVENTIVO DE COMPAÑIAS",
            "PROHIBICIONES DE LA AGD",
            "EMANCIPACIONES",
            "INTERDICCIONES",
            "ANTICRESIS",
            "PROHIBICION INECEL",
            "ARRIENDO MERCANTIL",
            "CONTRATOS DE PRESTAMO",
            "MARGINACIONES",
            "AVALUOS TRANSFERENCIA DE DOMINO ARCHIVO",
            "CANCELACIONES",
            "CERTIFICADOS DE GRAVAMENES ARCHIVO",
            "CERTIFICADOS LIBERATORIOS ARCHIVO DE",
            "SENTENCIAS",
            "PROPIEDADES MAYOR CUANTIA",
            "EMANCIPACIONES PADRES-HIJOS",
            "DOCUMENTOS MUTUO",
            "REPERTORIO"})


    Public Function getNameLibro(ByVal id As Integer) As String

        Dim indice As Integer = idLibros.IndexOf(id)

        Return nameLibros(indice)

    End Function

End Class
