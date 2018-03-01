Imports System.ComponentModel
Imports System.Globalization
Imports System.IO
'Imports System.Speech.Recognition
'Imports System.Speech.Synthesis
Imports System.Text
Imports System.Threading
Imports TradereIndex_APP.validedDecimal

Public Class frmIndexacion

    Private rutapdf As String
    Public variables As New VariablesGlobalesYfunciones()
    Public ComboInit As New FillCombo()
    Public seleccion As Integer = -1
    Public ModoEdit As Boolean = False
    Dim elemento As New Control(Nothing)
    'Dim microfono As New SpeechRecognitionEngine
    Dim digitosRepertorio As Integer = 4
    Dim cryp As New Simple3Des("123456")
    Dim seleccion2 As Integer = -1
    Dim microactive As Boolean = False
    Dim automplete As New AutocompleteItems()
    Dim validError As Boolean = False
    Dim sc() As Screen = Screen.AllScreens()
    Dim idpdf As String
    Dim dgv2(4, 100) As String
    Dim gm2(3, 100) As String
    Dim activeEditCC = False
    Dim waitActiveBtnCC = False
    Private docToCheck As Integer
    Private finishProcessCheck = False
    Private trd1 As Thread
    Private selectionMargin As Integer = -1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
        visorPDF.Close()
    End Sub

    Private Sub frmIndexacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If sc.Length > 1 Then
            visorPDF.Hide()
            visorPDF.StartPosition = FormStartPosition.Manual
            visorPDF.Location = New Point(sc(1).Bounds.X, 0)
            visorPDF.WindowState = FormWindowState.Maximized
            visorPDF.Show()
        End If
        CheckForIllegalCrossThreadCalls = False
        ComboInit.Fill(ComboBox2, variables.rutaPath + "Parroquia.txt")
        ComboInit.Fill(ComboBox1, variables.rutaPath + "LibroRegistral.txt")
        ComboInit.Fill(ComboBox4, variables.rutaPath + "DescBien.txt")
        ComboInit.Fill(ComboBox5, variables.rutaPath + "TipoContrato.txt")
        ComboInit.Fill(ComboBox8, variables.rutaPath + "Tipo_Entidad.txt")
        ComboInit.Fill(ComboBox9, variables.rutaPath + "Numero_Entidad.txt")
        ComboInit.Fill(ComboBox10, variables.rutaPath + "Provincias.txt")
        ComboInit.Fill(ComboBox11, variables.rutaPath + "Cantones.txt")
        ComboInit.Fill(ComboBox12, variables.rutaPath + "Rep_Entidades.txt")
        ''ComboInit.Fill(ComboBox3, variables.rutaPath + "Comparecientes.txt")
        ''MsgBox(MainForm.Size.Width.ToString)
        Dim columna1 As Double = 0.45
        Dim columna2 As Double = 0.55
        Dim pantalla As Size
        pantalla = System.Windows.Forms.SystemInformation.PrimaryMonitorSize

        Me.Width = pantalla.Width + 10
        Me.Height = pantalla.Height - 50

        Panel1.Width = (pantalla.Width - 20) * columna2
        Panel1.Height = (pantalla.Height) - 150

        If sc.Length > 1 Then
            AxAcroPDF1.Visible = False

            'CAMBIANDO POSICION Y DEFINIENTO TAMAÑAS DEL PANEL DE COMPARECIENTES

            Panel1.Controls.Add(Panel5)
            Panel5.Width = Panel1.Width - 15
            Panel5.Height = 300
            Panel5.Location = New Point(5, 5)
            DataGridView2.Height = 182
        End If

        PanelControls.Width = (pantalla.Width - 40) * columna2

        Panel2.Width = pantalla.Width * columna1
        If sc.Length > 1 Then
            Panel2.Height = Panel7.Height + 50
        End If


        Panel6.Width = pantalla.Width * columna1
        Panel4.Width = pantalla.Width * columna1
        Panel1.Location = New Point(Panel2.Location.X + Panel2.Width, 60)
        PanelControls.Location = New Point(Panel2.Location.X + Panel2.Width + 7, 5)

        Dim sobrante As Integer = (pantalla.Height - 20) - (Panel4.Height + Panel6.Height + Panel2.Height)
        Panel6.Height = sobrante + 50

        DataGridView1.Width = Panel6.Width - 11
        DataGridView1.Height = Panel6.Height - 10 ''- 50
        'Button5.Width = Panel6.Width - 11
        'Button5.Height = 30
        'Button5.Location = New Point(4, DataGridView1.Height + 10)
        Panel2.Location = New Point(4, sobrante + 110)
        Panel7.Width = Panel2.Width - 10
        If sc.Length <= 1 Then
            Panel5.Width = Panel2.Width - 10
        End If

        ''ORDENAR LOS ELEMENTOS DEL PANEL 5
        ''SE DEFINEN 2 COLUMNAS CON UN ESPACIO ENTRE CADA COLUMNA
        Dim espacioentre As Integer = 30
        Dim inicio As Integer = 10
        Dim anchoPanel As Integer = Panel5.Width - (inicio * 2)

        Dim tamanioColuma As Integer = anchoPanel / 2 - (espacioentre / 2)
        Label1.Location = New Point(inicio, 0)
        ComboBox3.Width = tamanioColuma
        ComboBox3.Location = New Point(inicio, Label1.Height + 3)


        Label8.Location = New Point((ComboBox3.Width + inicio) + espacioentre, 0)
        TextBox4.Width = tamanioColuma
        TextBox4.Location = New Point((ComboBox3.Width + inicio) + espacioentre, Label8.Height + 3)

        Label4.Location = New Point(inicio, (Label1.Height + 3) + ComboBox3.Height + 3)
        TextBox5.Width = tamanioColuma * 2 + espacioentre
        TextBox5.Location = New Point(Label4.Location.X, Label4.Location.Y + Label4.Height + 3)

        Label11.Location = New Point(TextBox5.Width + inicio + espacioentre, Label4.Location.Y)
        TextBox7.Width = tamanioColuma
        TextBox7.Location = New Point(Label8.Location.X, Label11.Location.Y + Label11.Height + 3)

        Button2.Width = tamanioColuma
        Button2.Location = New Point(inicio, TextBox5.Location.Y + TextBox5.Height + 3)

        Button4.Width = tamanioColuma
        Button4.Location = New Point(TextBox7.Location.X, TextBox7.Location.Y + TextBox7.Height + 3)

        DataGridView2.Width = Panel5.Width - 20
        DataGridView2.Location = New Point(inicio, Button2.Location.Y + Button2.Height + 5)

        TextBox2.Location = New Point(96 + 20, TextBox2.Location.Y)
        TextBox2.Width = (Panel7.Width / 2) - (96 + 30)

        TextBox6.Location = New Point(96 + 20, TextBox6.Location.Y)
        TextBox6.Width = (Panel7.Width / 2) - (96 + 30)

        ComboBox1.Location = New Point(96 + 20, ComboBox1.Location.Y)
        ComboBox1.Width = Panel7.Width - (96 + 30)

        ComboBox5.Location = New Point(96 + 20, ComboBox5.Location.Y)
        ComboBox5.Width = Panel7.Width - (96 + 30)

        DateTimePicker1.Location = New Point(96 + 20, DateTimePicker1.Location.Y)
        DateTimePicker1.Width = (Panel7.Width / 2) - (96 + 30)

        ComboBox2.Location = New Point(96 + 20, ComboBox2.Location.Y)
        ComboBox2.Width = (Panel7.Width - (96 + 30)) / 2

        TextBox8.Location = New Point(ComboBox2.Location.X + ComboBox2.Width + 8, TextBox8.Location.Y)
        TextBox8.Width = (Panel7.Width - (110 + 30)) / 2

        ComboBox4.Location = New Point(96 + 20, ComboBox4.Location.Y)
        ComboBox4.Width = Panel7.Width - (96 + 30)

        TextBox1.Location = New Point(96 + 20, TextBox1.Location.Y)
        TextBox1.Width = (Panel7.Width / 2) - (96 + 30)

        ComboBox8.Location = New Point(96 + 20, ComboBox8.Location.Y)
        ComboBox8.Width = (Panel7.Width - (96 + 20)) * 0.2 - 5

        ComboBox9.Location = New Point(ComboBox8.Location.X + ComboBox8.Width + 5, ComboBox9.Location.Y)
        ComboBox9.Width = (Panel7.Width - (96 + 20)) * 0.075 - 5

        ComboBox10.Location = New Point(ComboBox9.Location.X + ComboBox9.Width + 5, ComboBox10.Location.Y)
        ComboBox10.Width = (Panel7.Width - (96 + 20)) * 0.175 - 5

        ComboBox11.Location = New Point(ComboBox10.Location.X + ComboBox10.Width + 5, ComboBox11.Location.Y)
        ComboBox11.Width = (Panel7.Width - (96 + 20)) * 0.15 - 5

        ComboBox12.Location = New Point(ComboBox11.Location.X + ComboBox11.Width + 5, ComboBox12.Location.Y)
        ComboBox12.Width = (Panel7.Width - (96 + 20)) * 0.25 - 5

        Button8.Location = New Point(ComboBox12.Location.X + ComboBox12.Width + 5, Button8.Location.Y)
        Button8.Width = (Panel7.Width - (96 + 20)) * 0.075 - 5

        Button9.Location = New Point(Button8.Location.X + Button8.Width + 5, Button9.Location.Y)
        Button9.Width = (Panel7.Width - (96 + 20)) * 0.075 - 5

        MaskedTextBox1.Location = New Point(96 + 20, MaskedTextBox1.Location.Y)
        MaskedTextBox1.Width = (Panel7.Width / 2) - (96 + 30)

        gridMargin.Location = New Point(Label16.Location.X + Label16.Width + 5, gridMargin.Location.Y)
        gridMargin.Width = (Panel7.Width / 2) - 108

        Button1.Location = New Point(Panel4.Width - (Button1.Width + 30), Button1.Location.Y)
        Button3.Location = New Point(Panel4.Width - (Button3.Width + Button1.Width + 60), Button1.Location.Y)
        Panel3.Location = New Point(Panel4.Width - (Panel3.Width + Button3.Width + Button1.Width + 30), Panel3.Location.Y)

        Label3.Location = New Point(10, Label3.Location.Y)
        Label5.Location = New Point(10, Label5.Location.Y)
        Label6.Location = New Point(TextBox2.Location.X + TextBox2.Width + 5, Label6.Location.Y)
        Label2.Location = New Point(10, Label2.Location.Y)
        Label9.Location = New Point(10, Label9.Location.Y)
        Label12.Location = New Point(10, Label12.Location.Y)
        Label13.Location = New Point(10, Label13.Location.Y)
        Label14.Location = New Point(10, Label14.Location.Y)
        Label15.Location = New Point(10, Label15.Location.Y)
        Label16.Location = New Point(TextBox6.Location.X + TextBox6.Width + 5, Label16.Location.Y)
        Label10.Location = New Point(10, Label10.Location.Y)
        Label19.Location = New Point(Label6.Location.X, Label19.Location.Y)

        Button6.Width = Label16.Width
        Button6.Location = New Point(Label16.Location.X, Button6.Location.Y)

        gridMargin.Location = New Point(Label16.Location.X + Label16.Width + 5, gridMargin.Location.Y)
        gridMargin.Width = (Panel7.Width / 2) - 108

        TextBox3.Location = New Point(Label6.Location.X + Label6.Width + 5, TextBox3.Location.Y)
        TextBox3.Width = (Panel7.Width / 2) - (106)

        DateTimePicker2.Location = New Point(Label16.Location.X + Label16.Width + 5, DateTimePicker2.Location.Y)
        DateTimePicker2.Width = (Panel7.Width / 2) - (108)

        TextBox10.Location = New Point(96 + 20, TextBox10.Location.Y)
        TextBox10.Width = Panel7.Width - (96 + 30)

        TextBox9.Location = New Point(TextBox10.Location.X, TextBox9.Location.Y)
        TextBox9.Width = TextBox10.Width

        RadioButton1.Location = New Point(PanelControls.Width - RadioButton1.Width - 20 - btnResetIndex.Width, RadioButton1.Location.Y)
        RadioButton2.Location = New Point(PanelControls.Width - RadioButton1.Width * 2 - 30 - btnResetIndex.Width, RadioButton2.Location.Y)

        TextBox6.Width = (Panel7.Width / 2) - (96 + 30) - 40
        ComboBox7.Location = New Point(TextBox6.Location.X + TextBox6.Width + 5, ComboBox7.Location.Y)

        MaskedTextBox2.Location = New Point(TextBox3.Location.X, MaskedTextBox2.Location.Y)
        MaskedTextBox2.Width = (Panel7.Width / 2) - (106)


        btnCambiarModoControl.Location = New Point(PanelControls.Width - btnCambiarModoControl.Width - 10, btnCambiarModoControl.Location.Y)
        btnResetIndex.Location = New Point(PanelControls.Width - btnResetIndex.Width - 10, btnResetIndex.Location.Y)

        'microfono = New SpeechRecognitionEngine()
        'Dim comandos As String() = {"cedula", "nombres", "apellido", "compareciente", "repertorio", "libro", "inscripcion", "parroquia", "fecha", "siguiente", "agregar", "borrar", "actualizar", "cancelar"}
        'Dim VOCABULARIO As New GrammarBuilder
        'VOCABULARIO.Append(New Choices(comandos))
        'microfono.LoadGrammar(New Grammar(VOCABULARIO))

        'microfono.SetInputToDefaultAudioDevice()
        'microfono.RecognizeAsync(RecognizeMode.Multiple)
        'AddHandler microfono.SpeechRecognized, AddressOf RECONOCE
        'AddHandler microfono.SpeechRecognitionRejected, AddressOf NORECONOCE
        'AddHandler microfono.SpeechDetected, AddressOf DETECTA

        rutapdf = variables.ruta(0).ToString + "/pdf/"

        'For index = 0 To variables.columnas1.Count
        '    DataGridView1.Columns.Add(index, variables.columnas1(index).ToString)
        'Next

        ''SE CREAN LAS COLUMNAS DEL LIBRO 

        DataGridView1.Columns.Add(0, "id")
        DataGridView1.Columns.Add(1, "Repertorio")
        DataGridView1.Columns.Add(2, "Libro registral")
        DataGridView1.Columns.Add(3, "Tipo contrato")
        DataGridView1.Columns.Add(4, "Nº Inscripcion")
        DataGridView1.Columns.Add(5, "Fecha de incripción")
        DataGridView1.Columns.Add(6, "Fecha de repertorio")
        DataGridView1.Columns.Add(7, "Parroquia")
        DataGridView1.Columns.Add(8, "Descripción del bien")
        DataGridView1.Columns.Add(9, "Cuantía")
        DataGridView1.Columns.Add(10, "Moneda")
        DataGridView1.Columns.Add(11, "Fecha de otorgamiento")
        DataGridView1.Columns.Add(12, "Entidad")
        DataGridView1.Columns.Add(13, "# COMPARECIENTE")
        DataGridView1.Columns.Add(14, "FECHA CREADO")
        DataGridView1.Columns.Add(15, "CREADO")
        DataGridView1.Columns.Add(16, "FECHA REVISADO")
        DataGridView1.Columns.Add(17, "REVISADO")
        DataGridView1.Columns.Add(18, "FECHA CORREGIDO")
        DataGridView1.Columns.Add(19, "CORREGIDO")
        DataGridView1.Columns.Add(20, "FECHA ACEPTADO")
        DataGridView1.Columns.Add(21, "ACEPTADO")
        DataGridView1.Columns.Add(22, "IMAGEN INCORRECTA")
        DataGridView1.Columns.Add(23, "FECHA REG IMAGEN INCORRECTA")
        DataGridView1.Columns.Add(24, "Observaciones")

        DataGridView1.Columns(13).Visible = False
        DataGridView1.Columns(14).Visible = False
        DataGridView1.Columns(15).Visible = False
        DataGridView1.Columns(16).Visible = False
        DataGridView1.Columns(17).Visible = False
        DataGridView1.Columns(18).Visible = False
        DataGridView1.Columns(19).Visible = False
        DataGridView1.Columns(20).Visible = False
        DataGridView1.Columns(21).Visible = False
        DataGridView1.Columns(22).Visible = True
        DataGridView1.Columns(23).Visible = False

        ''DataGridView1.Columns(6).Visible = False

        ''SE CREAN LAS COLUMNAS DE LOS COMPARECIENTE
        DataGridView2.Columns.Add(0, "id")
        DataGridView2.Columns.Add(1, "Compareciente")
        DataGridView2.Columns.Add(2, "Cédula")
        DataGridView2.Columns.Add(3, "Nombres")
        'Dim grid As DataGridView = variables.crearColumna(DataGridView1, variables.columnas1)


        gridMargin.Columns.Add(0, "id")
        gridMargin.Columns(0).Visible = True
        gridMargin.Columns.Add(1, "Detalle")
        gridMargin.Columns.Add(2, "Fecha")

        Dim pdftxt As New StreamReader(variables.ruta(0) + variables.archivotext1)
        Dim linea As String
        Do
            linea = pdftxt.ReadLine()
            If linea IsNot Nothing Then
                Dim arrayPdfNombre As String() = linea.Split("|")
                Dim array(arrayPdfNombre.Count - 1) As String
                For index = 0 To array.Count - 2
                    If arrayPdfNombre(index).Trim() IsNot "" Then
                        array(index) = cryp.DecryptData(arrayPdfNombre(index))
                    End If
                Next
                DataGridView1.Rows.Add(array)
            End If
        Loop Until linea Is Nothing
        pdftxt.Close()
        'Dim leerArchivoGuia As New StreamReader(variables.ruta(0).ToArray + variables.archivoGuia)
        'MsgBox(rutapdf + DataGridView1.Rows(leerArchivoGuia.ReadLine().ToString).Cells(0).Value.ToString)
        'DataGridView1.Item(1, 0).Value = "hola"
        'DataGridView1.Rows(0).Selected = True

        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Ascending)


        DataGridView1.FirstDisplayedScrollingRowIndex = 0
        If MainForm.mode = 1 Then
            DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(0)
            'DataGridView1.Rows(0).DefaultCellStyle.BackColor = Color.Green
            'DataGridView1.Rows(0).DefaultCellStyle.ForeColor = Color.White
        Else
            DataGridView1.CurrentCell = DataGridView1.Rows(variables.obtenerPosicionFila()).Cells(0)
            DataGridView1.Rows(variables.obtenerPosicionFila()).DefaultCellStyle.BackColor = Color.Green
            DataGridView1.Rows(variables.obtenerPosicionFila()).DefaultCellStyle.ForeColor = Color.White
        End If

        lblFin.Text = DataGridView1.RowCount.ToString
        lblInicio.Text = (variables.obtenerPosicionFila() + 1).ToString
        '        MsgBox("hola entrada", vbOK)
        '        automplete.FillControls(TextBox5, "comparecientes.txt", 3, "|")
        '        MsgBox("chao entrada", vbOK)
        '        loadMetadate()

        If MainForm.mode = 1 Then

            Button5.Text = "C. CALIDAD Y ACTUALIZAR"
            Button5.Width = 160
            Button5.Enabled = False
            Button7.Enabled = False
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True
            RadioButton1.Visible = True
            RadioButton2.Visible = True

            lblFiltro.Visible = True
            cmbFiltro.Visible = True

            btnFiltrar.Location = New Point(180, btnFiltrar.Location.Y)
            btnFiltrar.Visible = True

            docToCheck = DataGridView1.Rows.Count * MainForm.porcentajeRevision

            lblInicio.Text = "Revisado    " + MainForm.listDocCheck.Count.ToString()
            Panel3.Width = 250
            Panel3.Location = New Point(Panel4.Width - (Panel3.Width + Button3.Width + Button1.Width + 5), Panel3.Location.Y)
            Label7.Location = New Point(140, Label7.Location.Y)
            lblFin.Location = New Point(180, lblFin.Location.Y)

            btnResetIndex.Visible = True

            lblFin.Text = docToCheck
        Else
            fillCambos(variables.obtenerPosicionFila())
            FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
            FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

            fillData(variables.obtenerPosicionFila())
        End If

        Button7.Location = New Point(Button5.Location.X + Button5.Width + 10, Button7.Location.Y)

        'If docToCheck = MainForm.listDocCheck.Count And MainForm.mode = 1 Then
        '    MsgBox("YA SE HA CUMPLIDO CON EL PORCENTAJE DE REVISIÓN")
        '    MainForm.IndexarToolStripMenuItem.Enabled = False
        '    Me.Close()
        'End If


        If sc.Length > 1 And MainForm.mode = 0 Then
            visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
        End If

        If MainForm.mode = 0 And (DataGridView1.Rows.Count - 1 = variables.obtenerPosicionFila()) Then
            Try
                If DataGridView1.Item(15, variables.obtenerPosicionFila()).Value.ToString() = "TRUE" Or DataGridView1.Item(22, variables.obtenerPosicionFila()).Value.ToString() = "TRUE" Then
                    btnCambiarModoControl.Visible = True
                End If
            Catch ex As Exception

            End Try
        End If

    End Sub

    Public Sub fillData(ByVal identy As Integer)

        Dim arrayStringDocName() As String = DataGridView1.Item(0, identy).Value.ToString().Split("-")

        If Len(DataGridView1.Item(0, identy).Value.ToString()) >= 22 Then
            DateTimePicker1.Text = arrayStringDocName(3) + "/" + arrayStringDocName(2) + "/" + arrayStringDocName(1)

            Dim a() As String = arrayStringDocName(4).Split(".")

            TextBox3.Text = a(0)
            Dim check As New setLibros()
            ComboBox1.Text = check.getNameLibro(Convert.ToInt32(arrayStringDocName(0)))
        Else
            Dim a() As String = arrayStringDocName(2).Split(".")
            TextBox3.Text = a(0)
            Dim check As New SetLibrosBAS()
            ComboBox1.Text = check.getNameLibro(Convert.ToInt32(arrayStringDocName(0)))
        End If
    End Sub
    'Public Sub RECONOCE(ByVal sender As Object, ByVal e As SpeechRecognizedEventArgs)
    '    If microactive Then
    '        Dim resultado As RecognitionResult
    '        resultado = e.Result
    '        Dim palabra As String = resultado.Text

    '        Select Case palabra
    '            Case "fecha"
    '                DateTimePicker1.Focus()
    '                elemento.BackColor = Color.White
    '                Return
    '            Case "nombres"
    '                TextBox5.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox5
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "cedula"
    '                TextBox4.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox4
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "apellido"
    '                TextBox7.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox7
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "compareciente"
    '                ComboBox3.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox3
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "agregar"
    '                If Button2.Text = "Agregar" Then
    '                    agregar()
    '                    If elemento IsNot Nothing Then
    '                        elemento.BackColor = Color.White
    '                        Label1.Focus()
    '                    End If
    '                End If
    '            Case "siguiente"
    '                siguiente()
    '            Case "borrar"
    '                If Button4.Enabled Then
    '                    Button4_Click(Nothing, Nothing)
    '                End If
    '            Case "cancelar"
    '                If Button2.Text = "Cancelar" Then
    '                    Button2_Click_1(Nothing, Nothing)
    '                End If
    '            Case "actualizar"
    '                If Button2.Text = "Actualizar" Then
    '                    Button2_Click_1(Nothing, Nothing)
    '                End If
    '            Case "repertorio"
    '                TextBox2.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox2
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "libro"
    '                ComboBox1.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox1
    '                    ComboBox1.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case "inscripcion"
    '                TextBox3.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = TextBox3
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If

    '            Case "parroquia"
    '                ComboBox2.Focus()
    '                If elemento IsNot Nothing Then
    '                    elemento.BackColor = Color.White
    '                    elemento = ComboBox2
    '                    elemento.BackColor = Color.FromName("Highlight")
    '                End If
    '            Case Else
    '                Exit Select
    '        End Select
    '    End If

    'End Sub

    'Public Sub DETECTA()

    'End Sub

    'Public Sub NORECONOCE()

    'End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        selectionMargin = -1
        Button6.Text = "Agregar"
        If Not (RadioButton1.Checked Or RadioButton2.Checked) And MainForm.mode = 1 Then
            RadioButton1.BackColor = Color.Red
            RadioButton1.ForeColor = Color.White
            RadioButton2.BackColor = Color.Red
            RadioButton2.ForeColor = Color.White
            MsgBox("Debe seleccionar una opción")
            Exit Sub
        Else
            RadioButton1.BackColor = Color.Transparent
            RadioButton1.ForeColor = Color.Black
            RadioButton2.BackColor = Color.Transparent
            RadioButton2.ForeColor = Color.Black
        End If

        Dim validatedCamps As New ValidatedControlsForms()

        Dim arrayControls As New List(Of Control)

        arrayControls = validatedCamps.check()

        If arrayControls.Count > 0 Then
            For index = 0 To arrayControls.Count - 1
                arrayControls(index).BackColor = Color.Red
                arrayControls(index).ForeColor = Color.White
            Next
            MsgBox("Los campos Marcados son Obligatorios, en caso que no contenga informacion rellene con N/D")
        Else
            If DataGridView2.Rows.Count > 0 Then
                siguiente()
            Else
                MsgBox("Necesitas por lo menos un compareciente")
            End If

        End If

        If MainForm.mode = 1 Then
            RadioButton1.Checked = False
            RadioButton2.Checked = False
            Button5.Enabled = False
            Button7.Enabled = False
            waitActiveBtnCC = False
        End If
    End Sub

    Public Function siguiente()
        If (ModoEdit) Then

            DataGridView1.Item(1, seleccion).Value = TextBox2.Text
            DataGridView1.Item(2, seleccion).Value = ComboBox1.SelectedItem.ToString
            DataGridView1.Item(3, seleccion).Value = ComboBox5.SelectedItem.ToString
            DataGridView1.Item(4, seleccion).Value = TextBox3.Text
            DataGridView1.Item(5, seleccion).Value = DateTimePicker1.Text.ToString
            DataGridView1.Item(6, seleccion).Value = MaskedTextBox2.Text
            DataGridView1.Item(7, seleccion).Value = TextBox8.Text
            DataGridView1.Item(8, seleccion).Value = ComboBox4.SelectedItem.ToString
            DataGridView1.Item(9, seleccion).Value = TextBox6.Text
            DataGridView1.Item(10, seleccion).Value = ComboBox7.Text
            DataGridView1.Item(11, seleccion).Value = DateTimePicker2.Text.ToString()
            DataGridView1.Item(12, seleccion).Value = TextBox10.Text
            DataGridView1.Item(13, seleccion).Value = DataGridView2.Rows.Count
            DataGridView1.Item(14, seleccion).Value = DateTime.Now()
            DataGridView1.Item(15, seleccion).Value = "TRUE"

            DataGridView1.Item(22, seleccion).Value = "FALSE"
            DataGridView1.Item(23, seleccion).Value = "NULL"

            DataGridView1.Item(24, seleccion).Value = TextBox9.Text

            If MainForm.mode = 1 Then
                DataGridView1.Item(16, seleccion).Value = "NULL"
                DataGridView1.Item(17, seleccion).Value = "FALSE"
                DataGridView1.Item(18, seleccion).Value = "NULL"
                DataGridView1.Item(19, seleccion).Value = "FALSE"
                DataGridView1.Item(20, seleccion).Value = "NULL"
                DataGridView1.Item(21, seleccion).Value = "FALSE"

                'DataGridView1.Columns.Add(14, "FECHA REVISADO")
                'DataGridView1.Columns.Add(15, "REVISADO")

                'DataGridView1.Columns.Add(16, "FECHA CORREGIDO")
                'DataGridView1.Columns.Add(17, "CORREGIDO")

                If RadioButton1.Checked Then
                    'SI LOS DATOS FUERON CORREGIDOS
                    DataGridView1.Item(18, seleccion).Value = DateTime.Now()
                    DataGridView1.Item(19, seleccion).Value = "TRUE"
                Else
                    'SI LOS DATOS SON CORRECTOS
                    DataGridView1.Item(16, seleccion).Value = DateTime.Now()
                    DataGridView1.Item(17, seleccion).Value = "TRUE"
                    'Action = "correcto"
                End If
                Button5.Enabled = False
                Button7.Enabled = False
            End If
            Button4.Enabled = False

            'Funcion que permite renombrar los archivos PDF
            'por la nomenclatura dada por el cliente

            'Dim NombrePdf As String
            'Dim NombrePdfOld As String = DataGridView1(0, seleccion).Value
            'NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
            'File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
            'DataGridView1(0, seleccion).Value = NombrePdf + ".pdf"

            If MainForm.mode = 1 Then

                If MainForm.listDocCheck.IndexOf(DataGridView1.Item(0, seleccion).Value.ToString()) = -1 Then
                    MainForm.listDocCheck.Add(DataGridView1.Item(0, seleccion).Value.ToString())
                End If

                If docToCheck = MainForm.listDocCheck.Count And MainForm.mode = 1 Then
                    MsgBox("YA SE HA CUMPLIDO CON EL PORCENTAJE DE REVISIÓN")
                    MsgBox("pruebaaa")

                    'MainForm.IndexarToolStripMenuItem.Enabled = False
                    FillDocsAcept.Acept(DataGridView1)

                End If

                If MainForm.listDocCheck.Count >= docToCheck And MainForm.mode = 1 Then
                    FillDocsAcept.Acept(DataGridView1)
                    '            guardarMetadatosPdf()
                End If

                generateReportControlCalidad()
            End If

            Button5.Enabled = False
            Button7.Enabled = False

            For i = 0 To DataGridView2.RowCount - 1
                For j = 0 To DataGridView2.ColumnCount - 1
                    dgv2(j, i) = DataGridView2(j, i).Value.ToString
                Next
            Next

            For i = 0 To gridMargin.RowCount - 1
                For j = 0 To gridMargin.ColumnCount - 1
                    gm2(j, i) = gridMargin(j, i).Value.ToString
                Next
            Next

            trd1 = New Thread(AddressOf guardarMetadatosPdf)
            trd1.IsBackground = True
            trd1.Start()

            limpiar1()
            limpiar2()

            If (Not MainForm.mode = 1) Then
                seleccion = seleccion + 1
                'lblInicio.Text = (seleccion + 1).ToString
                If (seleccion < variables.obtenerPosicionFila()) Then
                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                    If sc.Length > 1 Then
                        visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                    End If
                    'TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
                    'ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
                    'ComboBox5.Text = DataGridView1(3, seleccion).Value.ToString
                    'TextBox3.Text = DataGridView1(4, seleccion).Value.ToString
                    'DateTimePicker1.Text = DataGridView1(5, seleccion).Value
                    'TextBox8.Text = DataGridView1(6, seleccion).Value.ToString
                    'ComboBox4.Text = DataGridView1(7, seleccion).Value.ToString
                    'TextBox6.Text = DataGridView1.Item(8, seleccion).Value.ToString
                    'ComboBox7.Text = DataGridView1.Item(9, seleccion).Value.ToString
                    'DateTimePicker2.Text = DataGridView1.Item(10, seleccion).Value
                    'ComboBox6.Text = DataGridView1.Item(11, seleccion).Value.ToString

                    'Try
                    '    TextBox9.Text = DataGridView1.Item(23, seleccion).Value.ToString()
                    'Catch ex As Exception

                    'End Try

                    fillCambos(seleccion)

                    fillData(seleccion)

                    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                    DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.Cyan

                    If Not MainForm.mode = 1 Then
                        DataGridView1.CurrentCell = DataGridView1.Rows(seleccion).Cells(0)
                    End If
                Else
                    If (DataGridView1.Rows.Count - 1 = variables.obtenerPosicionFila()) And seleccion < DataGridView1.Rows.Count Then
                        AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                        If sc.Length > 1 Then
                            visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                        End If
                        fillCambos(variables.obtenerPosicionFila())
                        fillData(seleccion)
                        FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                        FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                        'Try
                        '    TextBox2.Text = DataGridView1(1, seleccion).Value.ToString
                        '    ComboBox1.Text = DataGridView1(2, seleccion).Value.ToString
                        '    ComboBox5.Text = DataGridView1(3, seleccion).Value.ToString
                        '    TextBox3.Text = DataGridView1(4, seleccion).Value.ToString
                        '    DateTimePicker1.Text = DataGridView1(5, seleccion).Value
                        '    TextBox8.Text = DataGridView1(6, seleccion).Value.ToString
                        '    ComboBox4.Text = DataGridView1(7, seleccion).Value.ToString
                        '    TextBox6.Text = DataGridView1.Item(8, seleccion).Value.ToString
                        '    ComboBox7.Text = DataGridView1.Item(9, seleccion).Value.ToString
                        '    DateTimePicker2.Text = DataGridView1.Item(10, seleccion).Value
                        '    ComboBox6.Text = DataGridView1.Item(11, seleccion).Value.ToString
                        '    TextBox9.Text = DataGridView1.Item(23, seleccion).Value.ToString()
                        '    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                        '    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                        'Catch ex As Exception

                        'End Try

                        fillCambos(seleccion)
                    Else

                        If MainForm.mode = 1 Then
                            Button5.Enabled = False
                        End If

                        If seleccion = DataGridView1.Rows.Count Then
                            limpiar1()
                            limpiar2()
                            DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                        Else
                            fillCambos(variables.obtenerPosicionFila())

                            FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                            FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)
                        End If
                    End If

                    If seleccion < DataGridView1.Rows.Count Then
                        DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                        AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                        If sc.Length > 1 Then
                            visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                        End If
                        fillData(seleccion)
                        DataGridView1.CurrentCell = DataGridView1.Rows(seleccion).Cells(0)
                    End If
                    ModoEdit = False
                    seleccion = -1
                End If
            Else
                If Not seleccion = -1 Then
                    DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
                End If
            End If

            'If finishProcessCheck Then
            '    Me.Close()
            'End If

        Else
            If variables.obtenerPosicionFila() < DataGridView1.RowCount Then

                Button2.Text = "Agregar"
                Button4.Enabled = False
                Label5.BackColor = Color.Transparent
                Label5.ForeColor = Color.Black
                DataGridView1.Item(1, variables.obtenerPosicionFila()).Value = TextBox2.Text
                DataGridView1.Item(2, variables.obtenerPosicionFila()).Value = ComboBox1.SelectedItem.ToString
                DataGridView1.Item(3, variables.obtenerPosicionFila()).Value = ComboBox5.SelectedItem.ToString
                DataGridView1.Item(4, variables.obtenerPosicionFila()).Value = TextBox3.Text
                DataGridView1.Item(5, variables.obtenerPosicionFila()).Value = DateTimePicker1.Text.ToString
                DataGridView1.Item(6, variables.obtenerPosicionFila()).Value = MaskedTextBox2.Text
                DataGridView1.Item(7, variables.obtenerPosicionFila()).Value = TextBox8.Text
                DataGridView1.Item(8, variables.obtenerPosicionFila()).Value = ComboBox4.SelectedItem.ToString
                DataGridView1.Item(9, variables.obtenerPosicionFila()).Value = TextBox6.Text
                DataGridView1.Item(10, variables.obtenerPosicionFila()).Value = ComboBox7.Text
                DataGridView1.Item(11, variables.obtenerPosicionFila()).Value = DateTimePicker2.Text.ToString()
                DataGridView1.Item(12, variables.obtenerPosicionFila()).Value = TextBox10.Text
                DataGridView1.Item(13, variables.obtenerPosicionFila()).Value = DataGridView2.Rows.Count
                DataGridView1.Item(14, variables.obtenerPosicionFila()).Value = DateTime.Now()
                DataGridView1.Item(15, variables.obtenerPosicionFila()).Value = "TRUE"
                DataGridView1.Item(16, variables.obtenerPosicionFila()).Value = "NULL"
                DataGridView1.Item(17, variables.obtenerPosicionFila()).Value = "FALSE"
                DataGridView1.Item(18, variables.obtenerPosicionFila()).Value = "NULL"
                DataGridView1.Item(19, variables.obtenerPosicionFila()).Value = "FALSE"
                DataGridView1.Item(20, variables.obtenerPosicionFila()).Value = "NULL"
                DataGridView1.Item(21, variables.obtenerPosicionFila()).Value = "FALSE"
                DataGridView1.Item(22, variables.obtenerPosicionFila()).Value = "FALSE"
                DataGridView1.Item(23, variables.obtenerPosicionFila()).Value = "NULL"
                DataGridView1.Item(24, variables.obtenerPosicionFila()).Value = TextBox9.Text

                'funcion que permite cambiar el nombre del pdf por un formato especifico
                'dado por el cliente

                'Dim NombrePdf As String
                'Dim NombrePdfOld As String = DataGridView1(0, variables.obtenerPosicionFila()).Value
                'NombrePdf = Format(DateTimePicker1.Value, "yyyyMMdd") + "-" + variables.retornarIdLibro(ComboBox1.SelectedItem.ToString) + "-" + variables.completarDigitos(Convert.ToInt64(TextBox2.Text))
                'File.Move(variables.ruta(0).ToString + "/pdf/" + NombrePdfOld, variables.ruta(0).ToString + "/pdf/" + NombrePdf + ".pdf")
                'DataGridView1(0, variables.obtenerPosicionFila()).Value = NombrePdf + ".pdf"

                Button5.Enabled = False
                Button7.Enabled = False

                '                trd2 = New Thread(Sub() Me.createChildrenMetadata1(DataGridView2, variables.ruta(0).ToString + variables.archivoCompareciente, "\comparecientestemp.txt"))
                '                trd2.IsBackground = True
                '                trd2.Start()

                '                trd3 = New Thread(Sub() Me.createChildrenMetadata2(gridMargin, variables.ruta(0).ToString + variables.archivoMarginaciones, "\marginacionestemp.txt"))
                '                trd3.IsBackground = True
                '                trd3.Start()

                For i = 0 To DataGridView2.RowCount - 1
                    For j = 0 To DataGridView2.ColumnCount - 1
                        dgv2(j, i) = DataGridView2(j, i).Value.ToString
                    Next
                Next

                For i = 0 To gridMargin.RowCount - 1
                    For j = 0 To gridMargin.ColumnCount - 1
                        gm2(j, i) = gridMargin(j, i).Value.ToString
                    Next
                Next

                trd1 = New Thread(AddressOf guardarMetadatosPdf)
                trd1.IsBackground = True
                trd1.Start()
                '                guardarMetadatosPdf()
                variables.MarcarFilaActual()
                limpiar1()
                limpiar2()
                If (variables.obtenerPosicionFila() < DataGridView1.RowCount - 1) Then
                    variables.cambiarPosicion(1)
                    fillCambos(variables.obtenerPosicionFila())
                    FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                    FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)
                    fillData(variables.obtenerPosicionFila())
                    lblInicio.Text = (1 + variables.obtenerPosicionFila()).ToString
                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    If sc.Length > 1 Then
                        visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    End If
                Else
                    'AQUI SE DEBE MARCAR QUE YA SE FINALIZO TODO

                    Button5.Enabled = False
                    Button7.Enabled = False
                    If MainForm.mode = 0 Then
                        MsgBox("NO HAY MAS PDF PARA INDEXAR")
                    End If
                    Dim opcion As String = MsgBox("DESEA CAMBIAR AL MODO CONTROL DE CALIDAD", vbOKCancel)
                    If opcion = 1 Then
                        cambiarModo()
                    End If
                    btnCambiarModoControl.Visible = True
                End If
            Else
                MsgBox("No hay mas pdf")
            End If
        End If
        'automplete.FillControls(TextBox5, "comparecientes.txt", 3, "|")
        'automplete.FillControls(TextBox7, "comparecientes.txt", 4, "|")
        Button2.Text = "Agregar"
    End Function

    Public Sub generateReportControlCalidad()
        CreateReportInspect()
        'If MainForm.listDocCheck.IndexOf(DataGridView1.Item(0, seleccion).Value.ToString()) = -1 Then
        '    MainForm.listDocCheck.Add(DataGridView1.Item(0, seleccion).Value.ToString())
        'End If

        lblInicio.Text = "Revisado    " + MainForm.listDocCheck.Count.ToString()

        'If docToCheck = MainForm.listDocCheck.Count And MainForm.mode = 1 Then
        '    MsgBox("YA SE HA CUMPLIDO CON EL PORCENTAJE DE REVISIÓN")
        '    'MainForm.IndexarToolStripMenuItem.Enabled = False
        '    FillDocsAcept.Acept(DataGridView1)

        '    '            guardarMetadatosPdf()
        '    For i = 0 To DataGridView2.RowCount - 1
        '        For j = 0 To DataGridView2.ColumnCount - 1
        '            dgv2(j, i) = DataGridView2(j, i).Value.ToString
        '        Next
        '    Next

        '    For i = 0 To gridMargin.RowCount - 1
        '        For j = 0 To gridMargin.ColumnCount - 1
        '            gm2(j, i) = gridMargin(j, i).Value.ToString
        '        Next
        '    Next

        '    trd1 = New Thread(AddressOf guardarMetadatosPdf)
        '    trd1.IsBackground = True
        '    trd1.Start()
        '    'finishProcessCheck = True
        'End If


        'If MainForm.listDocCheck.Count >= docToCheck And MainForm.mode = 1 Then
        '    FillDocsAcept.Acept(DataGridView1)
        '    '            guardarMetadatosPdf()
        '    For i = 0 To DataGridView2.RowCount - 1
        '        For j = 0 To DataGridView2.ColumnCount - 1
        '            dgv2(j, i) = DataGridView2(j, i).Value.ToString
        '        Next
        '    Next

        '    For i = 0 To gridMargin.RowCount - 1
        '        For j = 0 To gridMargin.ColumnCount - 1
        '            gm2(j, i) = gridMargin(j, i).Value.ToString
        '        Next
        '    Next

        '    trd1 = New Thread(AddressOf guardarMetadatosPdf)
        '    trd1.IsBackground = True
        '    trd1.Start()

        'End If
    End Sub

    Public Sub CreateReportInspect()
        If File.Exists(variables.ruta(0).ToString() + variables.archivoEstadisticas) Then
            Dim action As String
            Dim dateNow As String = DateTime.Now.ToString()
            Dim idDoc As String = DataGridView1(0, seleccion).Value.ToString()
            Dim operarioControl As String = "OPERARIO INDEXADOR"

            Dim writer As New StreamWriter(variables.ruta(0).ToString() + variables.archivoEstadisticas, True)
            If RadioButton1.Checked Then
                action = "corregido"
            Else
                action = "correcto"
            End If

            writer.WriteLine(cryp.EncryptData(idDoc) + "|" + cryp.EncryptData(action) + "|" + cryp.EncryptData(dateNow) + "|" + cryp.EncryptData(operarioControl) + "|" + cryp.EncryptData(MainForm.cedControlCalidad))
            writer.Close()
            writer.Dispose()

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        clearErrors()
        selectionMargin = -1
        Button6.Text = "Agregar"
        If MainForm.mode = 1 Then
            If (RadioButton1.Checked Or RadioButton2.Checked) Then

                RadioButton1.BackColor = Color.Transparent
                RadioButton1.ForeColor = Color.Black
                RadioButton2.BackColor = Color.Transparent
                RadioButton2.ForeColor = Color.Black

                DataGridView1.Item(16, seleccion).Value = "NULL"
                DataGridView1.Item(17, seleccion).Value = "FALSE"
                DataGridView1.Item(18, seleccion).Value = "NULL"
                DataGridView1.Item(19, seleccion).Value = "FALSE"
                DataGridView1.Item(20, seleccion).Value = "NULL"
                DataGridView1.Item(21, seleccion).Value = "FALSE"
                DataGridView1.Item(22, seleccion).Value = "TRUE"
                DataGridView1.Item(23, seleccion).Value = DateTime.Now()
                '                guardarMetadatosPdf()

                For i = 0 To DataGridView2.RowCount - 1
                    For j = 0 To DataGridView2.ColumnCount - 1
                        dgv2(j, i) = DataGridView2(j, i).Value.ToString
                    Next
                Next

                For i = 0 To gridMargin.RowCount - 1
                    For j = 0 To gridMargin.ColumnCount - 1
                        gm2(j, i) = gridMargin(j, i).Value.ToString
                    Next
                Next

                trd1 = New Thread(AddressOf guardarMetadatosPdf)
                trd1.IsBackground = True
                trd1.Start()

                Button5.Enabled = False
                Button7.Enabled = False
                ModoEdit = False
                generateReportControlCalidad()
                limpiar1()
                limpiar2()
                DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
                seleccion = -1
                RadioButton1.Checked = False
                RadioButton2.Checked = False
                Exit Sub
            Else
                RadioButton1.BackColor = Color.Red
                RadioButton1.ForeColor = Color.White
                RadioButton2.BackColor = Color.Red
                RadioButton2.ForeColor = Color.White
                MsgBox("Debe seleccionar una opción")
                Exit Sub
            End If
        End If

        If ModoEdit Then

            DataGridView1.Item(14, variables.obtenerPosicionFila()).Value = DateTime.Now()
            DataGridView1.Item(15, variables.obtenerPosicionFila()).Value = "TRUE"
            DataGridView1.Item(16, seleccion).Value = "NULL"
            DataGridView1.Item(17, seleccion).Value = "FALSE"
            DataGridView1.Item(18, seleccion).Value = "NULL"
            DataGridView1.Item(19, seleccion).Value = "FALSE"
            DataGridView1.Item(20, seleccion).Value = "NULL"
            DataGridView1.Item(21, seleccion).Value = "FALSE"
            DataGridView1.Item(22, seleccion).Value = "TRUE"
            DataGridView1.Item(23, seleccion).Value = DateTime.Now()
            '            guardarMetadatosPdf()

            For i = 0 To DataGridView2.RowCount - 1
                For j = 0 To DataGridView2.ColumnCount - 1
                    dgv2(j, i) = DataGridView2(j, i).Value.ToString
                Next
            Next

            For i = 0 To gridMargin.RowCount - 1
                For j = 0 To gridMargin.ColumnCount - 1
                    gm2(j, i) = gridMargin(j, i).Value.ToString
                Next
            Next

            trd1 = New Thread(AddressOf guardarMetadatosPdf)
            trd1.IsBackground = True
            trd1.Start()

            limpiar1()
            limpiar2()
            seleccion = seleccion + 1

            If (seleccion < variables.obtenerPosicionFila()) Then

                fillCambos(seleccion)
                fillData(seleccion)
                FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                If sc.Length > 1 Then
                    visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, seleccion).Value.ToString
                End If
                DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.Cyan
            Else

                If seleccion = DataGridView1.Rows.Count Then
                    limpiar1()
                    limpiar2()
                    DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                    seleccion = -1
                    ModoEdit = False
                Else
                    fillCambos(seleccion)
                    fillData(seleccion)
                    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                    FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)
                    ModoEdit = False
                    DataGridView1.Rows(seleccion - 1).DefaultCellStyle.BackColor = Color.White
                    DataGridView1.CurrentCell = DataGridView1.Rows(seleccion).Cells(0)
                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    If sc.Length > 1 Then
                        visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    End If
                    seleccion = -1
                End If


            End If
        Else

            DataGridView1.Item(14, variables.obtenerPosicionFila()).Value = DateTime.Now()
            DataGridView1.Item(15, variables.obtenerPosicionFila()).Value = "TRUE"
            DataGridView1.Item(16, variables.obtenerPosicionFila()).Value = "NULL"
            DataGridView1.Item(17, variables.obtenerPosicionFila()).Value = "FALSE"
            DataGridView1.Item(18, variables.obtenerPosicionFila()).Value = "NULL"
            DataGridView1.Item(19, variables.obtenerPosicionFila()).Value = "FALSE"
            DataGridView1.Item(20, variables.obtenerPosicionFila()).Value = "NULL"
            DataGridView1.Item(21, variables.obtenerPosicionFila()).Value = "FALSE"
            DataGridView1.Item(22, variables.obtenerPosicionFila()).Value = "TRUE"
            DataGridView1.Item(23, variables.obtenerPosicionFila()).Value = DateTime.Now()

            '            guardarMetadatosPdf()
            For i = 0 To DataGridView2.RowCount - 1
                For j = 0 To DataGridView2.ColumnCount - 1
                    dgv2(j, i) = DataGridView2(j, i).Value.ToString
                Next
            Next

            For i = 0 To gridMargin.RowCount - 1
                For j = 0 To gridMargin.ColumnCount - 1
                    gm2(j, i) = gridMargin(j, i).Value.ToString
                Next
            Next

            trd1 = New Thread(AddressOf guardarMetadatosPdf)
            trd1.IsBackground = True
            trd1.Start()

            variables.MarcarFilaActual()
            limpiar1()
            limpiar2()

            If (variables.obtenerPosicionFila() < DataGridView1.RowCount - 1) Then
                variables.cambiarPosicion(1)
                fillCambos(variables.obtenerPosicionFila())
                FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)
                fillData(variables.obtenerPosicionFila())
                lblInicio.Text = (1 + variables.obtenerPosicionFila()).ToString
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                If sc.Length > 1 Then
                    visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                End If
            Else
                'AQUI SE DEBE MARCAR QUE YA SE FINALIZO TODO
                Button5.Enabled = False
                Button7.Enabled = False
                If MainForm.mode = 0 Then
                    MsgBox("NO HAY MAS PDF PARA INDEXAR")
                    Dim opcion As String = MsgBox("DESEA CAMBIAR AL MODO CONTROL DE CALIDAD", vbOKCancel)
                    If opcion = 1 Then
                        cambiarModo()
                    End If
                    btnCambiarModoControl.Visible = True
                End If
            End If
        End If
        Button2.Text = "Agregar"
    End Sub

    Public Sub cambiarModo()
        If File.Exists(variables.ruta(0).ToString + variables.archivoEstadisticastemp) Then
            File.Move(variables.ruta(0).ToString + variables.archivoEstadisticastemp, variables.ruta(0).ToString + variables.archivoEstadisticas)
        Else
            If Not File.Exists(variables.ruta(0).ToString + variables.archivoEstadisticas) Then
                File.Create(variables.ruta(0).ToString + variables.archivoEstadisticas).Close()
            End If
        End If
        MsgBox("SE HA CAMBIADO A MODO CONTROL DE CALIDAD")
        MainForm.Close()
    End Sub

    Public Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        If elemento IsNot Nothing Then
            elemento.BackColor = Color.White
            Label1.Focus()
        End If
        agregar()
        'MsgBox(DataGridView1(0, variables.obtenerPosicionFila).Value.ToString)

    End Sub

    Public Function agregar()
        ''AxAcroPDF1.gotoLastPage()
        If ComboBox3.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(TextBox4.Text) Or String.IsNullOrWhiteSpace(TextBox5.Text) Then
            MsgBox("NO SE PUEDEN DEJAR CAMPOS VACIOS. RELLENE CON N/D")
            Exit Function
        End If

        If Button2.Text = "Agregar" Then
            If (ModoEdit) Then
                If ComboBox3.SelectedIndex = -1 Then
                    Label1.BackColor = Color.Red
                    Label1.ForeColor = Color.White
                    DataGridView2.Rows.Add(DataGridView1(0, seleccion).Value.ToString, "", TextBox4.Text, TextBox5.Text)
                Else
                    DataGridView2.Rows.Add(DataGridView1(0, seleccion).Value.ToString, ComboBox3.SelectedItem.ToString, TextBox4.Text, TextBox5.Text)
                    Label1.BackColor = Color.Transparent
                    Label1.ForeColor = Color.Black
                End If
                'limpiar2()
            Else
                If ComboBox3.SelectedIndex = -1 Or String.IsNullOrWhiteSpace(TextBox4.Text) Or String.IsNullOrWhiteSpace(TextBox5.Text) Then
                    MsgBox("NO SE PUEDEN DEJAR CAMPOS VACIOS. RELLENE CON N/D")
                Else
                    If ComboBox3.SelectedIndex = -1 Then
                        DataGridView2.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, "", TextBox4.Text, TextBox5.Text)
                    Else
                        DataGridView2.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, ComboBox3.SelectedItem.ToString, TextBox4.Text, TextBox5.Text)
                        Label1.BackColor = Color.Transparent
                        Label1.ForeColor = Color.Black
                    End If
                    'limpiar2()
                End If
            End If
        ElseIf Button2.Text = "Cancelar" Then
            'limpiar2()
            Button2.Text = "Agregar"
            seleccion2 = -1
            Button4.Enabled = False
        Else
            DataGridView2(1, seleccion2).Value = ComboBox3.SelectedItem.ToString
            DataGridView2(2, seleccion2).Value = TextBox4.Text
            DataGridView2(3, seleccion2).Value = TextBox5.Text
            seleccion2 = -1
            Button2.Text = "Agregar"
            Button4.Enabled = False
            'limpiar2()
        End If
    End Function

    Private Sub createChildrenMetadata1(ByVal grid As Array, ByVal path As String, ByVal archiveTemp As String)
        If (grid.GetLength(1) > 0) Then
            If (Not System.IO.File.Exists(path)) Then
                Dim creararchivo As FileStream
                creararchivo = File.Create(path)
                creararchivo.Close()
            End If

            Dim lectorCompareciente As New StreamReader(path)

            Dim archivoTemp As String = variables.ruta(0).ToString + archiveTemp
            Dim archivotemporal As FileStream
            archivotemporal = File.Create(archivoTemp)
            archivotemporal.Close()

            Dim linea As String
            Dim escribirCompaTemp As New StreamWriter(archivoTemp)
            Do
                linea = lectorCompareciente.ReadLine()
                If linea IsNot Nothing Then
                    Dim vectorLinea As String() = linea.Split("|")

                    Dim datoID As String = cryp.DecryptData(vectorLinea(0))

                    If (datoID.Equals(grid(0, 0))) Then
                        ''MsgBox(vectorLinea(0).ToString)
                    Else
                        escribirCompaTemp.WriteLine(linea)
                    End If
                    ''DataGridView1.Rows.Insert(0, New String() {linea})
                End If
            Loop Until linea Is Nothing

            lectorCompareciente.Close()

            Dim dataLine As String = ""
            For index = 0 To grid.GetLength(1) - 1
                If String.IsNullOrWhiteSpace(grid(0, index)) Then
                    Exit For
                End If
                If index > 0 Then
                    dataLine = dataLine + vbCrLf
                End If
                For i = 0 To grid.GetLength(0) - 1
                    Try
                        If Not i = grid.GetLength(0) Then
                            dataLine = dataLine + cryp.EncryptData(grid(i, index).ToString) + "|"
                        Else
                            dataLine = dataLine + cryp.EncryptData(IIf(String.IsNullOrEmpty(grid(i, index).ToString), DateTime.Now.ToShortDateString, grid(i, index).ToString))
                        End If
                    Catch ex As Exception
                        dataLine = dataLine + "|"
                    End Try
                Next
            Next
            If Not String.IsNullOrWhiteSpace(dataLine) Then
                escribirCompaTemp.WriteLine(dataLine)
            End If
            escribirCompaTemp.Close()
            Do While True
                Try
                    File.Delete(path)
                    Exit Do
                Catch ex As Exception
                    Thread.Sleep(500)
                End Try
            Loop
            File.Move(archivoTemp, path)
        End If
        Array.Clear(dgv2, 0, dgv2.Length)
    End Sub

    Private Sub createChildrenMetadata2(ByVal grid As Array, ByVal path As String, ByVal archiveTemp As String)
        '        If (grid.GetLength(1) > 0) Then
        If (Not System.IO.File.Exists(path)) Then
            Dim creararchivo As FileStream
            creararchivo = File.Create(path)
            creararchivo.Close()
        End If

        Dim lectorCompareciente As New StreamReader(path)

        Dim archivoTemp As String = variables.ruta(0).ToString + archiveTemp
        Dim archivotemporal As FileStream
        archivotemporal = File.Create(archivoTemp)
        archivotemporal.Close()

        Dim linea As String
        Dim escribirCompaTemp As New StreamWriter(archivoTemp)
        Do
            linea = lectorCompareciente.ReadLine()
            If linea IsNot Nothing Then
                Dim vectorLinea As String() = linea.Split("|")

                Dim datoID As String = cryp.DecryptData(vectorLinea(0))

                If (datoID.Equals(idpdf)) Then
                    ''MsgBox(vectorLinea(0).ToString)
                Else
                    escribirCompaTemp.WriteLine(linea)
                End If
                ''DataGridView1.Rows.Insert(0, New String() {linea})
            End If
        Loop Until linea Is Nothing

        lectorCompareciente.Close()

        Dim dataLine As String = ""
        For index = 0 To grid.GetLength(1) - 1
            If String.IsNullOrWhiteSpace(grid(0, index)) Then
                Exit For
            End If
            If index > 0 Then
                dataLine = dataLine + vbCrLf
            End If
            For i = 0 To grid.GetLength(0) - 1
                Try
                    If Not i = grid.GetLength(0) Then
                        dataLine = dataLine + cryp.EncryptData(grid(i, index).ToString) + "|"
                    Else
                        dataLine = dataLine + cryp.EncryptData(grid(i, index).ToString)
                    End If
                Catch ex As Exception
                    dataLine = dataLine + "|"
                End Try
            Next
        Next
        If Not String.IsNullOrWhiteSpace(dataLine) Then
            escribirCompaTemp.WriteLine(dataLine)
        End If
        escribirCompaTemp.Close()
        Do While True
            Try
                File.Delete(path)
                Exit Do
            Catch ex As Exception
                Thread.Sleep(500)
            End Try
        Loop
        File.Move(archivoTemp, path)
        '        End If
        Array.Clear(gm2, 0, gm2.Length)
    End Sub

    Public Function limpiar1()
        Me.TextBox2.Clear()
        Me.TextBox3.Clear()
        Me.TextBox1.Clear()
        Me.TextBox6.Clear()
        Me.TextBox9.Clear()
        Me.ComboBox1.SelectedIndex = -1
        Me.MaskedTextBox2.Text = ""
        Me.TextBox8.Clear()
        Me.ComboBox4.SelectedIndex = -1
        Me.ComboBox5.SelectedIndex = -1
        Me.TextBox10.Clear()
        Me.ComboBox7.SelectedIndex = -1
        Me.ComboBox3.Items.Clear()
        DateTimePicker1.Text = ""
        DateTimePicker2.Text = ""
        DataGridView2.Rows.Clear()
        gridMargin.Rows.Clear()
    End Function

    Public Function limpiar2()
        Me.ComboBox3.SelectedIndex = -1
        Me.TextBox4.Clear()
        Me.TextBox5.Clear()
        Me.TextBox7.Clear()
    End Function

    Private Sub guardarMetadatosPdf()
        activeEditCC = True
        idpdf = dgv2(0, 0)
        createChildrenMetadata1(dgv2, variables.ruta(0).ToString + variables.archivoCompareciente, "\comparecientestemp.txt")
        createChildrenMetadata2(gm2, variables.ruta(0).ToString + variables.archivoMarginaciones, "\marginacionestemp.txt")

        Dim txtPDFTemp As String = variables.ruta(0).ToString + "\pdfTemp.txt"
        Dim creartxt As FileStream
        creartxt = File.Create(txtPDFTemp)
        creartxt.Close()
        Dim escritorPDFMetadatos As New StreamWriter(txtPDFTemp)
        Dim linea As String = ""
        Dim cryp1 As New Simple3Des("123456")
        For index = 0 To DataGridView1.Rows.Count - 1
            If index > 0 Then
                linea = linea + vbCrLf
            End If
            For index2 = 0 To DataGridView1.Columns.Count - 1
                Dim dato As String = DataGridView1(index2, index).Value
                If dato = "" Or dato = " " Then
                    linea = linea + "|"
                Else
                    linea = linea + cryp1.EncryptData(dato) + "|"
                End If
            Next
            '            escritorPDFMetadatos.WriteLine(linea)
        Next
        escritorPDFMetadatos.WriteLine(linea)

        escritorPDFMetadatos.Close()

        File.Delete(variables.ruta(0) + variables.archivotext1)
        File.Move(txtPDFTemp, variables.ruta(0).ToString + variables.archivotext1)

        If MainForm.mode = 1 Then
            If Not Button5.Enabled And waitActiveBtnCC Then
                Button5.Enabled = True
                Button7.Enabled = True
            End If
        Else
            Button5.Enabled = True
            Button7.Enabled = True
        End If
        activeEditCC = False
        trd1.Abort()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim opcion As Integer = MsgBox("Seguro que deseas borrar estos datos", MsgBoxStyle.OkCancel)

        If opcion = 1 Then
            DataGridView2.Rows.RemoveAt(seleccion2)
            limpiar2()
            seleccion2 = -1
            Button4.Enabled = False
            Button2.Text = "Agregar"
        Else

        End If

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        If e.RowIndex <> -1 Then
            waitActiveBtnCC = True
            selectionMargin = -1
            Button6.Text = "Agregar"
            'Try
            '    If DataGridView1(21, e.RowIndex).Value.ToString() = "TRUE" And Not MainForm.mode = 1 Then
            '        MsgBox("ESTE REGISTRO HA SIDO MARCADO COMO ERRONEO. NO SE PUEDE EDITAR")
            '        Exit Sub
            '    End If
            'Catch ex As Exception
            'End Try



            Button7.Enabled = True
            seleccion2 = -1
            limpiar2()
            If MainForm.mode = 1 Then
                RadioButton1.BackColor = Color.Transparent
                RadioButton1.ForeColor = Color.Black
                RadioButton2.BackColor = Color.Transparent
                RadioButton2.ForeColor = Color.Black
            Else
                Button5.Enabled = True
            End If

            clearErrors()
            'If Not MainForm.mode = 1 Then
            '    lblInicio.Text = (e.RowIndex + 1).ToString
            'End If
            If Not seleccion = -1 Then
                DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
            End If

            If e.RowIndex < variables.obtenerPosicionFila Then
                limpiar1()
                limpiar2()
                Button2.Text = "Agregar"
                seleccion = e.RowIndex
                ModoEdit = True
                fillCambos(seleccion)
                fillData(seleccion)
                DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(0)

                Dim id As String = DataGridView1(0, seleccion).Value.ToString
                AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + id
                If sc.Length > 1 Then
                    visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + id
                End If
                FillComparecientes(id, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                FillComparecientes(id, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Cyan

            Else
                If e.RowIndex = variables.obtenerPosicionFila() Then

                    If DataGridView1.Rows.Count - 1 = variables.obtenerPosicionFila() Then
                        seleccion = e.RowIndex
                        ModoEdit = True
                        fillCambos(seleccion)
                        fillData(seleccion)
                        DataGridView1.CurrentCell = DataGridView1.Rows(e.RowIndex).Cells(0)
                        seleccion = e.RowIndex
                        FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                        FillComparecientes(DataGridView1(0, seleccion).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)
                        DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Cyan
                    Else


                        fillCambos(variables.obtenerPosicionFila())
                        FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoCompareciente, DataGridView2)
                        FillComparecientes(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, variables.ruta(0).ToString + variables.archivoMarginaciones, gridMargin)

                        fillData(variables.obtenerPosicionFila())
                        Button2.Text = "Agregar"
                        ModoEdit = False
                        seleccion = -1
                    End If

                    AxAcroPDF1.src = variables.ruta(0).ToString + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    If sc.Length > 1 Then
                        visorPDF.visorPDFSecondScreen.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
                    End If
                Else
                    If Not seleccion = -1 Then
                        DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.Cyan
                    End If

                    MsgBox("No se puede mover a esta fila por que el puntero esta en la fila  " + (variables.obtenerPosicionFila + 1).ToString + ".Solamente puede moverve en filas anteriores")
                End If

            End If
            If Not activeEditCC And MainForm.mode = 1 Then
                Button5.Enabled = True

            End If
        End If

    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        seleccion2 = e.RowIndex
        If Not e.RowIndex = -1 Then
            Try
                ComboBox3.SelectedItem = DataGridView2(1, e.RowIndex).Value
            Catch ex As Exception
                ComboBox3.SelectedItem = ""
            End Try
            Try
                TextBox4.Text = DataGridView2(2, e.RowIndex).Value.ToString
            Catch ex As Exception
                TextBox4.Text = ""
            End Try
            Try
                TextBox5.Text = DataGridView2(3, e.RowIndex).Value.ToString
            Catch ex As Exception
                TextBox5.Text = ""
            End Try

            Button4.Enabled = True
            Button2.Text = "Cancelar"
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If microactive Then
            Button3.BackgroundImage = Image.FromFile(Application.StartupPath.ToString + "\microdisable.png")
            microactive = False
        Else
            Button3.BackgroundImage = Image.FromFile(Application.StartupPath.ToString + "\microactive.png")
            microactive = True
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        cambiarEstado()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        cambiarEstado()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        cambiarEstado()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        cambiarEstado()
    End Sub

    Function cambiarEstado()
        If seleccion2 > -1 Then
            If Button2.Text <> "Actualizar" Then
                Button2.Text = "Actualizar"
            End If
        End If
    End Function

    Private Sub AxAcroPDF1_Enter(sender As Object, e As EventArgs) Handles AxAcroPDF1.Enter
        AxAcroPDF1.src = variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString
        'lectorPdf = New iTextSharp.text.pdf.PdfReader(variables.ruta(0) + "\pdf\" + DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString)
        'MsgBox(lectorPdf.NumberOfPages)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        ComboBox3.Items.Clear()
        ComboBox3.SelectedIndex = -1
        ComboBox3.Text = ""
        If Not ComboBox5.SelectedIndex = -1 Then
            Dim indice As String = Convert.ToString(ComboBox5.SelectedItem.ToString)
            Dim reader As New StreamReader(variables.rutaPath + "Comparecientes.txt")
            Dim linea As String

            Dim findResult As Boolean = False

            Do
                linea = reader.ReadLine()
                If (Not linea Is Nothing) Then
                    Dim vectorLinea As String() = linea.Split(",")
                    If (vectorLinea(0).Equals(indice)) Then
                        ComboBox3.Enabled = True
                        ComboBox3.Items.Add(vectorLinea(1))
                        findResult = True
                    End If
                End If
            Loop Until linea Is Nothing

            If Not findResult Then
                ComboBox3.Items.Add("ACTOR")
                ComboBox3.Items.Add("RECEPTOR")
            End If
            reader.Close()
            'combobox3.visible = false
            'task.delay(200)
            'combobox3.visible = true
        End If
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged

        If ComboBox10.SelectedIndex > -1 Then
            ComboBox11.Items.Clear()
            ComboBox11.SelectedIndex = -1
            ComboBox11.Text = ""
            If Not ComboBox10.SelectedIndex = -1 Then
                Dim indice As String = Convert.ToString(ComboBox10.SelectedItem.ToString)
                Dim reader As New StreamReader(variables.rutaPath + "Cantones.txt")
                Dim linea As String

                Dim findResult As Boolean = False

                Do
                    linea = reader.ReadLine()
                    If (Not linea Is Nothing) Then
                        Dim vectorLinea As String() = linea.Split(",")
                        If (vectorLinea(1).Equals(indice)) Then
                            ComboBox11.Enabled = True
                            ComboBox11.Items.Add(vectorLinea(0))
                            findResult = True
                        End If
                    End If
                Loop Until linea Is Nothing

                reader.Close()
            End If

            ComboBox11.Enabled = True
        Else
            ComboBox11.Items.Clear()
            ComboBox11.Enabled = False
        End If
    End Sub



    Private Sub TextBox4_LostFocus(sender As Object, e As EventArgs) Handles TextBox4.LostFocus

        'Dim cedula As String = TextBox4.Text
        'Dim catalogoRuta As String = variables.ruta(0) + "/catalogos/userData.txt"

        'If (File.Exists(catalogoRuta)) Then
        '    If (Not String.IsNullOrWhiteSpace(cedula)) Then
        '        Dim reader As New StreamReader(catalogoRuta, Encoding.Default)
        '        Dim linea As String
        '        Do
        '            linea = reader.ReadLine()
        '            If (linea Is Nothing) Then
        '                Exit Do
        '            End If
        '            Dim ArrayLine As String() = linea.Split(",")
        '            If (ArrayLine(0).Equals(cedula)) Then
        '                TextBox5.Text = ArrayLine(1)
        '                TextBox7.Text = ArrayLine(2)
        '                reader.Dispose()
        '                Exit Do
        '            End If
        '        Loop Until linea Is Nothing
        '        reader.Dispose()
        '    End If
        'End If
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        'TextBox5.Text = ""
        'TextBox7.Text = ""
    End Sub

    Private Sub DateTimePicker1_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles DateTimePicker1.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not e.IsValidInput Then
            ToolTip1.ToolTipTitle = "Fecha invalida"
            ToolTip1.Show("La fecha ingresada es incorrecta", DateTimePicker1, 0, -DateTimePicker1.Location.Y + DateTimePicker1.Height - 10)
            DateTimePicker1.Focus()
        Else
            Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

            If userData > DateTime.Now Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", DateTimePicker1, 0, -DateTimePicker1.Location.Y + DateTimePicker1.Height - 10)
                DateTimePicker1.Focus()
            End If

        End If
    End Sub
    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        Dim valid As New validedDecimal()
        If Not valid.validated(e, TextBox6.Text) Then
            e.Handled = True
        End If
    End Sub

    Private Sub DateTimePicker2_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles DateTimePicker2.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not e.IsValidInput Then
            ToolTip1.ToolTipTitle = "Fecha invalida"
            ToolTip1.Show("La fecha ingresada es incorrecta", DateTimePicker2, 0, -73)
            DateTimePicker2.Focus()
        Else
            Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

            If userData > DateTime.Now Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", DateTimePicker2, 0, -73)
                DateTimePicker2.Focus()
            End If

        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        Me.ToolTip1.Hide(Me.DateTimePicker1)
    End Sub

    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        Me.ToolTip1.Hide(Me.DateTimePicker2)
    End Sub

    Private Sub loadMetadate()
        If DataGridView1.Rows.Count = variables.obtenerPosicionFila() + 1 Then
            'seleccion = DataGridView1.Rows.Count - 1
            'fillCambos()

        End If
    End Sub

    Private Sub fillCambos(ByVal rowsView As Integer)
        Try
            TextBox2.Text = DataGridView1(1, rowsView).Value.ToString
        Catch ex As Exception
            TextBox2.Text = ""
        End Try

        Try
            ComboBox1.Text = DataGridView1(2, rowsView).Value.ToString
        Catch ex As Exception
            ComboBox1.SelectedIndex = -1
        End Try

        Try
            ComboBox5.Text = DataGridView1(3, rowsView).Value.ToString
        Catch ex As Exception
            ComboBox5.SelectedIndex = -1
        End Try

        Try
            TextBox3.Text = DataGridView1(4, rowsView).Value.ToString
        Catch ex As Exception
            TextBox3.Text = ""
        End Try

        Try
            DateTimePicker1.Text = DataGridView1(5, rowsView).Value
        Catch ex As Exception
            DateTimePicker1.Text = ""
        End Try

        Try
            MaskedTextBox2.Text = DataGridView1(6, rowsView).Value.ToString
        Catch ex As Exception
            MaskedTextBox2.Text = ""
        End Try

        Try
            TextBox8.Text = DataGridView1(7, rowsView).Value.ToString
        Catch ex As Exception
            TextBox8.Text = ""
        End Try

        Try
            ComboBox4.Text = DataGridView1(8, rowsView).Value.ToString
        Catch ex As Exception
            ComboBox4.SelectedIndex = -1
        End Try

        Try
            TextBox6.Text = DataGridView1.Item(9, rowsView).Value.ToString
        Catch ex As Exception
            TextBox6.Text = ""
        End Try

        Try
            ComboBox7.Text = DataGridView1.Item(10, rowsView).Value
        Catch ex As Exception
            ComboBox7.SelectedIndex = -1
        End Try

        Try
            DateTimePicker2.Text = DataGridView1.Item(11, rowsView).Value
        Catch ex As Exception
            DateTimePicker2.Text = ""
        End Try

        Try
            TextBox10.Text = DataGridView1.Item(12, rowsView).Value.ToString
        Catch ex As Exception
            TextBox10.Clear()
        End Try

        Try
            TextBox9.Text = DataGridView1.Item(24, rowsView).Value.ToString
        Catch ex As Exception
            TextBox9.Text = ""
        End Try

        Try

        Catch ex As Exception

        End Try

    End Sub

    Private Sub FillComparecientes(ByVal id As String, ByVal ruta As String, ByVal grid1 As DataGridView)


        grid1.Rows.Clear()
        If File.Exists(ruta) Then
            Dim linea As String
            Dim lectorCompareciente As New StreamReader(ruta)

            Do
                linea = lectorCompareciente.ReadLine()
                If linea IsNot Nothing Then

                    Dim arrayLinea As String() = linea.Split("|")
                    Dim datoId As String = cryp.DecryptData(arrayLinea(0))
                    If datoId.ToString.Equals(id) Then
                        Dim array(arrayLinea.Count - 1) As String
                        For index = 0 To array.Count - 1
                            If arrayLinea(index).Trim() IsNot "" Then
                                array(index) = cryp.DecryptData(arrayLinea(index))
                            End If
                        Next

                        grid1.Rows.Add(array)
                    End If
                End If
            Loop Until linea Is Nothing
            lectorCompareciente.Close()
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Button6.Text = "Editar" Or Button6.Text = "Agregar" Then
            If Not TextBox1.Text.Trim() = "" And Not MaskedTextBox1.Text.Trim() = "/  /" Then
                Label13.BackColor = Color.Transparent
                Label13.ForeColor = Color.Black
                If Not validError Then
                    Label17.BackColor = Color.Transparent
                    Label17.ForeColor = Color.Black
                    If selectionMargin <> -1 Then
                        gridMargin.Item(1, selectionMargin).Value = TextBox1.Text
                        gridMargin.Item(2, selectionMargin).Value = MaskedTextBox1.Text
                        selectionMargin = -1
                        Button6.Text = "Agregar"
                    Else
                        If seleccion <> -1 Then
                            gridMargin.Rows.Add(DataGridView1(0, seleccion).Value.ToString, TextBox1.Text, MaskedTextBox1.Text)
                        Else
                            gridMargin.Rows.Add(DataGridView1(0, variables.obtenerPosicionFila()).Value.ToString, TextBox1.Text, MaskedTextBox1.Text)
                        End If

                    End If

                    TextBox1.Clear()
                    MaskedTextBox1.Clear()
                    TextBox1.Focus()
                Else
                    Label17.BackColor = Color.Red
                    Label17.ForeColor = Color.White
                    MsgBox("DEBE INGRESAR UNA FECHA CORRECTA")
                End If
            Else
                Label13.BackColor = Color.Red
                Label13.ForeColor = Color.White
                Label17.BackColor = Color.Red
                Label17.ForeColor = Color.White
                MsgBox("COMPOS OBLIGATORIOS")
            End If
        ElseIf (Button6.Text = "Borrar") Then
            Dim opcion As Integer = MsgBox("Seguro que deseas borrar estos datos", MsgBoxStyle.OkCancel)

            If opcion = 1 Then
                gridMargin.Rows.RemoveAt(selectionMargin)
                Button6.Text = "Agregar"
                selectionMargin = -1
                TextBox1.Clear()
                MaskedTextBox1.Clear()
            End If
        End If

    End Sub

    Private Sub MaskedTextBox1_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles MaskedTextBox1.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not MaskedTextBox1.Text.Trim().Equals("/  /") Then
            If Not e.IsValidInput Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es incorrecta", MaskedTextBox1, 0, -73)
                MaskedTextBox1.Focus()
                validError = True
            Else
                Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

                If userData > DateTime.Now Then
                    ToolTip1.ToolTipTitle = "Fecha invalida"
                    ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", MaskedTextBox1, 0, -73)
                    MaskedTextBox1.Focus()
                    validError = True
                Else
                    validError = False
                End If

            End If
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox1.KeyDown
        Me.ToolTip1.Hide(Me.MaskedTextBox1)
    End Sub

    Private Sub clearErrors()
        Dim labelArray As New List(Of Label)
        labelArray.Add(Label3)
        labelArray.Add(Label6)
        labelArray.Add(Label5)
        labelArray.Add(Label14)
        labelArray.Add(Label2)
        labelArray.Add(Label9)
        labelArray.Add(Label12)
        labelArray.Add(Label13)
        labelArray.Add(Label17)
        labelArray.Add(Label15)
        labelArray.Add(Label16)
        labelArray.Add(Label10)
        labelArray.Add(Label19)

        For index = 0 To labelArray.Count - 1
            labelArray(index).BackColor = Color.Transparent
            labelArray(index).ForeColor = Color.Black
        Next

    End Sub



    Private Sub cmbFiltro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFiltro.SelectedIndexChanged
        If cmbFiltro.SelectedIndex = 1 Then
            panelfiltro.Visible = True
            btnFiltrar.Location = New Point(350, btnFiltrar.Location.Y)
        Else
            panelfiltro.Visible = False
            btnFiltrar.Location = New Point(180, btnFiltrar.Location.Y)
        End If
    End Sub

    Private Sub btnFiltrar_Click(sender As Object, e As EventArgs) Handles btnFiltrar.Click
        limpiar1()
        limpiar2()
        Button5.Enabled = False
        If Not seleccion = -1 Then
            DataGridView1.Rows(seleccion).DefaultCellStyle.BackColor = Color.White
        End If
        showRowsAll()
        Dim selectedindex As Integer = cmbFiltro.SelectedIndex

        If selectedindex = -1 Then
            MsgBox("Debes seleccionar un tipo de filtro")
        ElseIf (selectedindex = 0) Then
            showRowsAll()
        ElseIf (selectedindex = 1) Then
            filterCompareciente()
        ElseIf (selectedindex = 2) Then
            FilterDataEmpty()
        ElseIf (selectedindex = 3) Then
            FilterActsIncorrect()
        End If
    End Sub

    Public Sub showRowsAll()
        For index = 0 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(index).Visible = True
        Next
    End Sub

    Private Function filterCompareciente()
        Try
            Dim number As Integer = Convert.ToInt32(txtNumber.Text)

            For index = 0 To DataGridView1.Rows.Count - 1

                Dim numberCompareciente As Integer = Convert.ToInt32(DataGridView1(13, index).Value.ToString())

                If RadioButton3.Checked Then
                    If Not numberCompareciente > number Then
                        DataGridView1.Rows(index).Visible = False
                    End If
                ElseIf RadioButton4.Checked Then
                    If Not numberCompareciente < number Then
                        DataGridView1.Rows(index).Visible = False
                    End If
                ElseIf RadioButton5.Checked Then
                    If Not numberCompareciente = number Then
                        DataGridView1.Rows(index).Visible = False
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox("INGRESE SOLO NUMEROS ENTEROS Y QUE SEAN POSITIVOS")
        End Try


    End Function


    Private Sub FilterDataEmpty()
        Dim ColumnasValidatedData() As Integer = {1, 4, 8}
        Dim ColumnasValidatedComparecientes() As Integer = {3, 4}
        Dim ColumnasValidatedMargination() As Integer = {1}

        Dim filter As String = "N/D"

        For fila = 0 To DataGridView1.Rows.Count - 1

            For columm = 0 To ColumnasValidatedData.Length - 1

                Try
                    If Not DataGridView1(ColumnasValidatedData(columm), fila).Value.ToString().ToUpper().Equals(filter) Then
                        DataGridView1.Rows(fila).Visible = False
                    Else
                        Exit For
                    End If
                Catch ex As Exception
                    DataGridView1.Rows(fila).Visible = False
                End Try
            Next

        Next

    End Sub

    Private Sub FilterActsIncorrect()
        Dim filter As String = "TRUE"
        For index = 0 To DataGridView1.Rows.Count - 1
            If Not DataGridView1(22, index).Value.ToString.ToUpper.Equals(filter) Then
                DataGridView1.Rows(index).Visible = False
            End If
        Next
    End Sub

    Private Sub gridMargin_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles gridMargin.CellDoubleClick
        If e.RowIndex <> -1 Then
            selectionMargin = e.RowIndex
            Try
                TextBox1.Text = gridMargin.Item(1, e.RowIndex).Value.ToString()
            Catch ex As Exception
                TextBox1.Text = ""
            End Try
            Try
                MaskedTextBox1.Text = gridMargin.Item(2, e.RowIndex).Value.ToString()
            Catch ex As Exception
                MaskedTextBox1.Text = ""
            End Try

            Button6.Text = "Borrar"
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        If selectionMargin <> -1 Then
            Button6.Text = "Editar"
        End If
    End Sub

    Private Sub MaskedTextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MaskedTextBox1.KeyPress
        If selectionMargin <> -1 Then
            Button6.Text = "Editar"
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        'If Not ComboBox2.SelectedIndex = -1 Then
        '    If String.IsNullOrWhiteSpace(TextBox8.Text) Then
        '        TextBox8.Text = ComboBox2.Text
        '    Else
        '        TextBox8.Text = TextBox8.Text + "," + ComboBox2.Text
        '    End If
        '    ComboBox2.SelectedIndex = -1
        'End If
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If e.KeyChar = ChrW(Keys.Back) Then
            If Not String.IsNullOrWhiteSpace(TextBox8.Text) Then
                Dim arrayParroquias As String() = TextBox8.Text.Split(",")
                TextBox8.Clear()
                For index = 0 To arrayParroquias.Length - 2
                    If index = arrayParroquias.Length - 2 Then
                        TextBox8.Text = TextBox8.Text + arrayParroquias(index)
                    Else
                        TextBox8.Text = TextBox8.Text + arrayParroquias(index) + ","
                    End If
                Next
                TextBox8.Select(TextBox8.TextLength, 0)
                e.Handled = True
            End If
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox8_Click(sender As Object, e As EventArgs) Handles TextBox8.Click
        TextBox8.Select(TextBox8.TextLength, 0)
    End Sub

    Private Sub txtNumber_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles txtNumber.MaskInputRejected

    End Sub

    Private Sub ComboBox2_TextChanged(sender As Object, e As EventArgs) Handles ComboBox2.TextChanged
        If Not ComboBox2.SelectedIndex = -1 Then
            If String.IsNullOrWhiteSpace(TextBox8.Text) Then
                TextBox8.Text = ComboBox2.Text
            Else
                TextBox8.Text = TextBox8.Text + "," + ComboBox2.Text
            End If
            ComboBox2.SelectedIndex = -1
        End If
    End Sub

    Private Sub btnCambiarModoControl_Click(sender As Object, e As EventArgs) Handles btnCambiarModoControl.Click
        cambiarModo()
    End Sub

    Private Sub btnResetIndex_Click(sender As Object, e As EventArgs) Handles btnResetIndex.Click
        Dim result As DialogResult = MsgBox("ESTA SEGURO QUE DESEA CAMBIAR DE MODO INDEXADOR", MsgBoxStyle.YesNo)
        If result = DialogResult.Yes Then
            If File.Exists(variables.ruta(0).ToString + variables.archivoEstadisticas) Then
                File.Move(variables.ruta(0).ToString + variables.archivoEstadisticas, variables.ruta(0).ToString + variables.archivoEstadisticastemp)
                'File.Delete(variables.ruta(0).ToString + variables.archivoGuia)
                MsgBox("SE HA CAMBIADO A MODO INDEXADOR. ES NECESARIO CERRAR LA APLICACIÓN")
                MainForm.Close()
            End If
        End If

    End Sub

    Private Sub MaskedTextBox2_TypeValidationCompleted(sender As Object, e As TypeValidationEventArgs) Handles MaskedTextBox2.TypeValidationCompleted
        ToolTip1.RemoveAll()
        If Not e.IsValidInput Then
            ToolTip1.ToolTipTitle = "Fecha invalida"
            ToolTip1.Show("La fecha ingresada es incorrecta", MaskedTextBox2, 0, -MaskedTextBox2.Location.Y + MaskedTextBox2.Height - 10)
            MaskedTextBox2.Focus()
        Else
            Dim userData As DateTime = DirectCast(e.ReturnValue, DateTime)

            If userData > DateTime.Now Then
                ToolTip1.ToolTipTitle = "Fecha invalida"
                ToolTip1.Show("La fecha ingresada es mayor a la fecha actual", MaskedTextBox2, 0, -MaskedTextBox2.Location.Y + MaskedTextBox2.Height - 10)
                MaskedTextBox2.Focus()
            End If

        End If
    End Sub

    Private Sub MaskedTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBox2.KeyDown
        Me.ToolTip1.Hide(Me.MaskedTextBox2)
    End Sub

    Private Sub DataGridView1_ColumnAdded(sender As Object, e As DataGridViewColumnEventArgs) Handles DataGridView1.ColumnAdded
        e.Column.SortMode = DataGridViewColumnSortMode.NotSortable
    End Sub

    Private Sub DateTimePicker1_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles DateTimePicker1.MaskInputRejected

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox10.Text = ComboBox8.Text.Trim + "-" + ComboBox9.Text.Trim + "-" + ComboBox10.Text.Trim + "-" + ComboBox11.Text.Trim + "-" + ComboBox12.Text.Trim
        ComboBox8.SelectedIndex = -1
        ComboBox9.SelectedIndex = -1
        ComboBox10.SelectedIndex = -1
        ComboBox11.SelectedIndex = -1
        ComboBox12.SelectedIndex = -1
        ComboBox12.Text = ""
        ComboBox9.Enabled = False
        ComboBox10.Enabled = False
        ComboBox11.Enabled = False
        ComboBox12.Enabled = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox10.Clear()
        ComboBox8.SelectedIndex = -1
        ComboBox9.SelectedIndex = -1
        ComboBox10.SelectedIndex = -1
        ComboBox11.SelectedIndex = -1
        ComboBox12.SelectedIndex = -1
        ComboBox12.Text = ""
        ComboBox9.Enabled = False
        ComboBox10.Enabled = False
        ComboBox11.Enabled = False
        ComboBox12.Enabled = False
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        If ComboBox8.SelectedIndex > -1 Then
            ComboBox9.Enabled = True
            ComboBox10.Enabled = True
            ComboBox12.Enabled = True
            Button8.Enabled = True
        Else
            ComboBox9.Enabled = False
            ComboBox10.Enabled = False
            ComboBox12.Enabled = False
            Button8.Enabled = False
        End If
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged

    End Sub
End Class