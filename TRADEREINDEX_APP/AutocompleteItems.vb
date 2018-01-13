Imports System.Windows.Forms
Imports System.IO

Imports TradereIndex_APP.frmIndexacion


Public Class AutocompleteItems


    Private variables As New VariablesGlobalesYfunciones()
    Private cryp As New Simple3Des("123456")
    Public Sub FillControls(ByVal control As TextBox, ByVal pathItemsAutocomplete As String, ByVal column As Integer, ByVal separator As Char)
        control.AutoCompleteCustomSource.Clear()

        Dim ruta As String = variables.ruta(0) + "/" + pathItemsAutocomplete

        If File.Exists(ruta) Then
            Dim lector As New StreamReader(ruta)
            Dim linea As String = ""
            Do
                linea = lector.ReadLine()
                If (linea Is Nothing) Then
                Else
                    Dim vectorLinea As String() = linea.Split(separator)
                    If Not String.IsNullOrWhiteSpace(vectorLinea(column)) Then
                        control.AutoCompleteCustomSource.Add(cryp.DecryptData(vectorLinea(column)))
                        control.AutoCompleteMode = AutoCompleteMode.Suggest
                        control.AutoCompleteSource = AutoCompleteSource.CustomSource
                    End If
                End If
            Loop Until linea Is Nothing
            lector.Dispose()
        End If
    End Sub


End Class
