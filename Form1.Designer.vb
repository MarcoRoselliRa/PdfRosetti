<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        txtPdfPath = New TextBox()
        dgvReport = New DataGridView()
        picPreview = New PictureBox()
        prbLoad = New ProgressBarConTesto()
        dgvTotali = New DataGridView()
        btnAggiungiBookmark = New Button()
        btnInviaStampante1 = New Button()
        btnSfogliaIndice = New Button()
        txtIndicePath = New TextBox()
        GroupBox1 = New GroupBox()
        chkPuntaDopoDiv = New CheckBox()
        chkMostraVociSenzaDivisorio = New CheckBox()
        btnSalvaImpostazioni = New Button()
        Label3 = New Label()
        txtCartaDivisorio = New TextBox()
        Label2 = New Label()
        txtCartaA3 = New TextBox()
        Label1 = New Label()
        txtCartaA4 = New TextBox()
        btnSfogliaStampante1 = New Button()
        btnSfogliaStampante2 = New Button()
        txtStampante1 = New TextBox()
        txtStampante2 = New TextBox()
        btnInviaStampante2 = New Button()
        Label4 = New Label()
        Label5 = New Label()
        Label6 = New Label()
        GroupBox2 = New GroupBox()
        btnCaricaImpostazioniFile = New Button()
        btnSalvaImpostazioniFile = New Button()
        chbSettaggi = New CheckBox()
        nudCopie = New NumericUpDown()
        Label7 = New Label()
        btnReset = New Button()
        CType(dgvReport, ComponentModel.ISupportInitialize).BeginInit()
        CType(picPreview, ComponentModel.ISupportInitialize).BeginInit()
        CType(dgvTotali, ComponentModel.ISupportInitialize).BeginInit()
        GroupBox1.SuspendLayout()
        GroupBox2.SuspendLayout()
        CType(nudCopie, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' txtPdfPath
        ' 
        txtPdfPath.Location = New Point(28, 32)
        txtPdfPath.Name = "txtPdfPath"
        txtPdfPath.ReadOnly = True
        txtPdfPath.Size = New Size(890, 39)
        txtPdfPath.TabIndex = 0
        ' 
        ' dgvReport
        ' 
        dgvReport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvReport.Location = New Point(29, 114)
        dgvReport.Name = "dgvReport"
        dgvReport.RowHeadersWidth = 82
        dgvReport.Size = New Size(1162, 1023)
        dgvReport.TabIndex = 1
        ' 
        ' picPreview
        ' 
        picPreview.BorderStyle = BorderStyle.FixedSingle
        picPreview.Location = New Point(1220, 114)
        picPreview.Name = "picPreview"
        picPreview.Size = New Size(577, 701)
        picPreview.SizeMode = PictureBoxSizeMode.Zoom
        picPreview.TabIndex = 2
        picPreview.TabStop = False
        ' 
        ' prbLoad
        ' 
        prbLoad.Location = New Point(1220, 32)
        prbLoad.Name = "prbLoad"
        prbLoad.Size = New Size(577, 42)
        prbLoad.TabIndex = 4
        prbLoad.Testo = ""
        ' 
        ' dgvTotali
        ' 
        dgvTotali.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvTotali.Location = New Point(1220, 844)
        dgvTotali.Name = "dgvTotali"
        dgvTotali.RowHeadersWidth = 82
        dgvTotali.Size = New Size(577, 293)
        dgvTotali.TabIndex = 5
        ' 
        ' btnAggiungiBookmark
        ' 
        btnAggiungiBookmark.AutoSize = True
        btnAggiungiBookmark.Location = New Point(28, 1167)
        btnAggiungiBookmark.Name = "btnAggiungiBookmark"
        btnAggiungiBookmark.Size = New Size(419, 71)
        btnAggiungiBookmark.TabIndex = 6
        btnAggiungiBookmark.Text = "Salvo Pdf con Divisori"
        btnAggiungiBookmark.UseVisualStyleBackColor = True
        ' 
        ' btnInviaStampante1
        ' 
        btnInviaStampante1.Location = New Point(477, 1167)
        btnInviaStampante1.Name = "btnInviaStampante1"
        btnInviaStampante1.Size = New Size(419, 71)
        btnInviaStampante1.TabIndex = 7
        btnInviaStampante1.Text = "Xerox C9281A (Dispari)"
        btnInviaStampante1.UseVisualStyleBackColor = True
        ' 
        ' btnSfogliaIndice
        ' 
        btnSfogliaIndice.Location = New Point(29, 83)
        btnSfogliaIndice.Name = "btnSfogliaIndice"
        btnSfogliaIndice.Size = New Size(114, 42)
        btnSfogliaIndice.TabIndex = 8
        btnSfogliaIndice.Text = "Sfoglia"
        btnSfogliaIndice.UseVisualStyleBackColor = True
        ' 
        ' txtIndicePath
        ' 
        txtIndicePath.Location = New Point(159, 83)
        txtIndicePath.Name = "txtIndicePath"
        txtIndicePath.Size = New Size(660, 39)
        txtIndicePath.TabIndex = 9
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(chkPuntaDopoDiv)
        GroupBox1.Controls.Add(chkMostraVociSenzaDivisorio)
        GroupBox1.Controls.Add(btnSalvaImpostazioni)
        GroupBox1.Controls.Add(Label3)
        GroupBox1.Controls.Add(txtCartaDivisorio)
        GroupBox1.Controls.Add(Label2)
        GroupBox1.Controls.Add(txtCartaA3)
        GroupBox1.Controls.Add(Label1)
        GroupBox1.Controls.Add(txtCartaA4)
        GroupBox1.Location = New Point(64, 732)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(935, 332)
        GroupBox1.TabIndex = 10
        GroupBox1.TabStop = False
        GroupBox1.Text = "Nomi Carte"
        GroupBox1.Visible = False
        ' 
        ' chkPuntaDopoDiv
        ' 
        chkPuntaDopoDiv.AutoSize = True
        chkPuntaDopoDiv.Location = New Point(587, 146)
        chkPuntaDopoDiv.Name = "chkPuntaDopoDiv"
        chkPuntaDopoDiv.Size = New Size(289, 36)
        chkPuntaDopoDiv.TabIndex = 8
        chkPuntaDopoDiv.Text = "Punta dopo il Divisorio"
        chkPuntaDopoDiv.UseVisualStyleBackColor = True
        ' 
        ' chkMostraVociSenzaDivisorio
        ' 
        chkMostraVociSenzaDivisorio.AutoSize = True
        chkMostraVociSenzaDivisorio.Location = New Point(587, 104)
        chkMostraVociSenzaDivisorio.Name = "chkMostraVociSenzaDivisorio"
        chkMostraVociSenzaDivisorio.Size = New Size(312, 36)
        chkMostraVociSenzaDivisorio.TabIndex = 7
        chkMostraVociSenzaDivisorio.Text = "Mostra Divisori Mancanti"
        chkMostraVociSenzaDivisorio.UseVisualStyleBackColor = True
        ' 
        ' btnSalvaImpostazioni
        ' 
        btnSalvaImpostazioni.Location = New Point(587, 230)
        btnSalvaImpostazioni.Name = "btnSalvaImpostazioni"
        btnSalvaImpostazioni.Size = New Size(167, 67)
        btnSalvaImpostazioni.TabIndex = 6
        btnSalvaImpostazioni.Text = "Salva carte"
        btnSalvaImpostazioni.UseVisualStyleBackColor = True
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(33, 224)
        Label3.Name = "Label3"
        Label3.Size = New Size(93, 32)
        Label3.TabIndex = 5
        Label3.Text = "Divisori"
        ' 
        ' txtCartaDivisorio
        ' 
        txtCartaDivisorio.Location = New Point(31, 260)
        txtCartaDivisorio.Name = "txtCartaDivisorio"
        txtCartaDivisorio.Size = New Size(493, 39)
        txtCartaDivisorio.TabIndex = 4
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(33, 146)
        Label2.Name = "Label2"
        Label2.Size = New Size(138, 32)
        Label2.TabIndex = 3
        Label2.Text = "Formato A3"
        ' 
        ' txtCartaA3
        ' 
        txtCartaA3.Location = New Point(31, 182)
        txtCartaA3.Name = "txtCartaA3"
        txtCartaA3.Size = New Size(493, 39)
        txtCartaA3.TabIndex = 2
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(33, 68)
        Label1.Name = "Label1"
        Label1.Size = New Size(138, 32)
        Label1.TabIndex = 1
        Label1.Text = "Formato A4"
        ' 
        ' txtCartaA4
        ' 
        txtCartaA4.Location = New Point(31, 104)
        txtCartaA4.Name = "txtCartaA4"
        txtCartaA4.Size = New Size(493, 39)
        txtCartaA4.TabIndex = 0
        ' 
        ' btnSfogliaStampante1
        ' 
        btnSfogliaStampante1.Location = New Point(29, 160)
        btnSfogliaStampante1.Name = "btnSfogliaStampante1"
        btnSfogliaStampante1.Size = New Size(114, 42)
        btnSfogliaStampante1.TabIndex = 11
        btnSfogliaStampante1.Text = "Sfoglia"
        btnSfogliaStampante1.UseVisualStyleBackColor = True
        ' 
        ' btnSfogliaStampante2
        ' 
        btnSfogliaStampante2.Location = New Point(29, 237)
        btnSfogliaStampante2.Name = "btnSfogliaStampante2"
        btnSfogliaStampante2.Size = New Size(114, 42)
        btnSfogliaStampante2.TabIndex = 12
        btnSfogliaStampante2.Text = "Sfoglia"
        btnSfogliaStampante2.UseVisualStyleBackColor = True
        ' 
        ' txtStampante1
        ' 
        txtStampante1.Location = New Point(159, 160)
        txtStampante1.Name = "txtStampante1"
        txtStampante1.Size = New Size(660, 39)
        txtStampante1.TabIndex = 13
        ' 
        ' txtStampante2
        ' 
        txtStampante2.Location = New Point(159, 237)
        txtStampante2.Name = "txtStampante2"
        txtStampante2.Size = New Size(660, 39)
        txtStampante2.TabIndex = 14
        ' 
        ' btnInviaStampante2
        ' 
        btnInviaStampante2.Location = New Point(927, 1167)
        btnInviaStampante2.Name = "btnInviaStampante2"
        btnInviaStampante2.Size = New Size(419, 71)
        btnInviaStampante2.TabIndex = 15
        btnInviaStampante2.Text = "Xerox C9281B (Pari)"
        btnInviaStampante2.UseVisualStyleBackColor = True
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(159, 48)
        Label4.Name = "Label4"
        Label4.Size = New Size(167, 32)
        Label4.TabIndex = 16
        Label4.Text = "File testi Indici"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(159, 125)
        Label5.Name = "Label5"
        Label5.Size = New Size(238, 32)
        Label5.TabIndex = 17
        Label5.Text = "Percorso Xerox9281A"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(159, 202)
        Label6.Name = "Label6"
        Label6.Size = New Size(237, 32)
        Label6.TabIndex = 18
        Label6.Text = "Percorso Xerox9281B"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(btnCaricaImpostazioniFile)
        GroupBox2.Controls.Add(btnSalvaImpostazioniFile)
        GroupBox2.Controls.Add(Label6)
        GroupBox2.Controls.Add(Label5)
        GroupBox2.Controls.Add(Label4)
        GroupBox2.Controls.Add(txtStampante2)
        GroupBox2.Controls.Add(txtStampante1)
        GroupBox2.Controls.Add(btnSfogliaStampante2)
        GroupBox2.Controls.Add(btnSfogliaStampante1)
        GroupBox2.Controls.Add(txtIndicePath)
        GroupBox2.Controls.Add(btnSfogliaIndice)
        GroupBox2.Location = New Point(64, 341)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(934, 373)
        GroupBox2.TabIndex = 19
        GroupBox2.TabStop = False
        GroupBox2.Text = "Percorsi"
        GroupBox2.Visible = False
        ' 
        ' btnCaricaImpostazioniFile
        ' 
        btnCaricaImpostazioniFile.Location = New Point(489, 303)
        btnCaricaImpostazioniFile.Name = "btnCaricaImpostazioniFile"
        btnCaricaImpostazioniFile.Size = New Size(287, 39)
        btnCaricaImpostazioniFile.TabIndex = 20
        btnCaricaImpostazioniFile.Text = "Importa Impostazioni"
        btnCaricaImpostazioniFile.UseVisualStyleBackColor = True
        ' 
        ' btnSalvaImpostazioniFile
        ' 
        btnSalvaImpostazioniFile.Location = New Point(159, 303)
        btnSalvaImpostazioniFile.Name = "btnSalvaImpostazioniFile"
        btnSalvaImpostazioniFile.Size = New Size(287, 39)
        btnSalvaImpostazioniFile.TabIndex = 19
        btnSalvaImpostazioniFile.Text = "Esporta Impostazioni"
        btnSalvaImpostazioniFile.UseVisualStyleBackColor = True
        ' 
        ' chbSettaggi
        ' 
        chbSettaggi.Appearance = Appearance.Button
        chbSettaggi.AutoSize = True
        chbSettaggi.Location = New Point(1079, 32)
        chbSettaggi.Name = "chbSettaggi"
        chbSettaggi.Size = New Size(112, 42)
        chbSettaggi.TabIndex = 20
        chbSettaggi.Text = "Settaggi"
        chbSettaggi.UseVisualStyleBackColor = True
        ' 
        ' nudCopie
        ' 
        nudCopie.Location = New Point(948, 35)
        nudCopie.Maximum = New Decimal(New Integer() {3, 0, 0, 0})
        nudCopie.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        nudCopie.Name = "nudCopie"
        nudCopie.Size = New Size(93, 39)
        nudCopie.TabIndex = 21
        nudCopie.Value = New Decimal(New Integer() {3, 0, 0, 0})
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(948, 0)
        Label7.Name = "Label7"
        Label7.Size = New Size(76, 32)
        Label7.TabIndex = 22
        Label7.Text = "Copie"
        ' 
        ' btnReset
        ' 
        btnReset.Location = New Point(1376, 1167)
        btnReset.Name = "btnReset"
        btnReset.Size = New Size(419, 71)
        btnReset.TabIndex = 23
        btnReset.Text = "Reset"
        btnReset.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AllowDrop = True
        AutoScaleDimensions = New SizeF(13F, 32F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1835, 1267)
        Controls.Add(btnReset)
        Controls.Add(Label7)
        Controls.Add(nudCopie)
        Controls.Add(chbSettaggi)
        Controls.Add(GroupBox2)
        Controls.Add(btnInviaStampante2)
        Controls.Add(GroupBox1)
        Controls.Add(btnInviaStampante1)
        Controls.Add(btnAggiungiBookmark)
        Controls.Add(dgvTotali)
        Controls.Add(prbLoad)
        Controls.Add(picPreview)
        Controls.Add(dgvReport)
        Controls.Add(txtPdfPath)
        Name = "Form1"
        Text = "PdfConteggio Rosetti"
        CType(dgvReport, ComponentModel.ISupportInitialize).EndInit()
        CType(picPreview, ComponentModel.ISupportInitialize).EndInit()
        CType(dgvTotali, ComponentModel.ISupportInitialize).EndInit()
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        CType(nudCopie, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtPdfPath As TextBox
    Friend WithEvents dgvReport As DataGridView
    Friend WithEvents picPreview As PictureBox
    Friend WithEvents prbLoad As ProgressBarConTesto
    Friend WithEvents dgvTotali As DataGridView
    Friend WithEvents btnAggiungiBookmark As Button
    Friend WithEvents btnInviaStampante1 As Button
    Friend WithEvents btnSfogliaIndice As Button
    Friend WithEvents txtIndicePath As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label3 As Label
    Friend WithEvents txtCartaDivisorio As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents txtCartaA3 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents txtCartaA4 As TextBox
    Friend WithEvents btnSalvaImpostazioni As Button
    Friend WithEvents btnSfogliaStampante1 As Button
    Friend WithEvents btnSfogliaStampante2 As Button
    Friend WithEvents txtStampante1 As TextBox
    Friend WithEvents txtStampante2 As TextBox
    Friend WithEvents btnInviaStampante2 As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents chbSettaggi As CheckBox
    Friend WithEvents chkPuntaDopoDiv As CheckBox
    Friend WithEvents chkMostraVociSenzaDivisorio As CheckBox
    Friend WithEvents nudCopie As NumericUpDown
    Friend WithEvents Label7 As Label
    Friend WithEvents btnCaricaImpostazioniFile As Button
    Friend WithEvents btnSalvaImpostazioniFile As Button
    Friend WithEvents btnReset As Button

End Class
