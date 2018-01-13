Public Class validedDecimal

    Public Function validated(ByVal e As KeyPressEventArgs, ByVal valor As String) As Boolean
        Dim simbolDecimal As Char = "."
        Dim digitos As Integer = 2
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Back) Then
            Return True
        End If

        If Not Char.IsNumber(e.KeyChar) Then
            If Not e.KeyChar = simbolDecimal Then
                Return False
            Else
                If valor.Contains(simbolDecimal) Then
                    Return False

                End If
            End If
        Else
            If valor.IndexOf(simbolDecimal) > -1 Then
                If (valor.IndexOf(simbolDecimal) + digitos) < (valor.Length) Then
                    Return False
                End If

            End If
        End If
        Return True
    End Function

    Public Function CompleteSimbolMil(ByVal Text As String) As String

        Dim array() As String = Text.Split(",")

        Dim newText As String = String.Join("", array)

        If newText.Length >= 4 Then
            Dim contador As Integer = 0
            Dim newString As String = ""
            For index = newText.Length - 1 To 0 Step -1
                If contador = 3 Then
                    contador = 1
                    newString = newString + "," + newText(index)
                Else
                    newString = newString + newText(index)
                    contador = contador + 1
                End If
            Next

            Dim stringFinal As String

            For index = newString.Length - 1 To 0 Step -1
                stringFinal = stringFinal + newString(index)
            Next

            frmIndexacion.TextBox6.Text = stringFinal
        End If
    End Function

End Class
