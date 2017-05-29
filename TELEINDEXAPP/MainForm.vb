Imports System
Imports System.IO

Public Class MainForm

    Public Shared identificacion As String
    Public Shared apellidos As String
    Public Shared nombres As String

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '        frmUsersEntry.Show()
        '       frmUsersEntry.MdiParent = Me
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox("Desea Salir del Sistema Registral?", vbYesNo) = 6 Then
            Me.Close()
        End If
    End Sub

End Class
