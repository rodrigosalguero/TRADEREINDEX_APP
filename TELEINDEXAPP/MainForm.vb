Imports System
Imports System.IO

Public Class MainForm

    Public Shared identificacion As String
    Public Shared nombre As String

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        UserEntry.Show()
        UserEntry.MdiParent = Me
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        If MsgBox("Desea Salir del Sistema Registral?", vbYesNo) = 6 Then
            Me.Close()
        End If
    End Sub

    Private Sub IndexarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IndexarToolStripMenuItem.Click
        frmIndexacion.Show()
        frmIndexacion.MdiParent = Me
    End Sub
End Class
