<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class visorPDF
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(visorPDF))
        Me.visorPDFSecondScreen = New AxAcroPDFLib.AxAcroPDF()
        CType(Me.visorPDFSecondScreen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'visorPDFSecondScreen
        '
        Me.visorPDFSecondScreen.Dock = System.Windows.Forms.DockStyle.Fill
        Me.visorPDFSecondScreen.Enabled = True
        Me.visorPDFSecondScreen.Location = New System.Drawing.Point(0, 0)
        Me.visorPDFSecondScreen.Name = "visorPDFSecondScreen"
        Me.visorPDFSecondScreen.OcxState = CType(resources.GetObject("visorPDFSecondScreen.OcxState"), System.Windows.Forms.AxHost.State)
        Me.visorPDFSecondScreen.Size = New System.Drawing.Size(486, 290)
        Me.visorPDFSecondScreen.TabIndex = 0
        '
        'visorPDF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(486, 290)
        Me.Controls.Add(Me.visorPDFSecondScreen)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "visorPDF"
        Me.Text = "Visor PDF"
        CType(Me.visorPDFSecondScreen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents visorPDFSecondScreen As AxAcroPDFLib.AxAcroPDF
End Class
