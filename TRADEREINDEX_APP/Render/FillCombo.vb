Imports System.IO
Imports System.Text
Imports System.Windows.Forms

Public Class FillCombo
    Public Sub Fill(ByVal el As ComboBox, ByVal UrlArchive As String)
        Dim ContentDatList As New StreamReader(UrlArchive, Encoding.UTF8)
        Dim linea As String
        Do
            linea = ContentDatList.ReadLine()
            If linea Is Nothing Then
                Exit Do
            End If
            el.Items.Add(linea)
        Loop Until linea Is Nothing
        ContentDatList.Dispose()
    End Sub
End Class
