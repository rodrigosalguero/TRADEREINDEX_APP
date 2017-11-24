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

End Class
