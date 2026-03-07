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
        CType(dgvReport, ComponentModel.ISupportInitialize).BeginInit()
        CType(picPreview, ComponentModel.ISupportInitialize).BeginInit()
        CType(dgvTotali, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' txtPdfPath
        ' 
        txtPdfPath.Location = New Point(28, 32)
        txtPdfPath.Name = "txtPdfPath"
        txtPdfPath.ReadOnly = True
        txtPdfPath.Size = New Size(1163, 39)
        txtPdfPath.TabIndex = 0
        ' 
        ' dgvReport
        ' 
        dgvReport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvReport.Location = New Point(29, 114)
        dgvReport.Name = "dgvReport"
        dgvReport.RowHeadersWidth = 82
        dgvReport.Size = New Size(1162, 1124)
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
        btnAggiungiBookmark.Location = New Point(1220, 1167)
        btnAggiungiBookmark.Name = "btnAggiungiBookmark"
        btnAggiungiBookmark.Size = New Size(312, 71)
        btnAggiungiBookmark.TabIndex = 6
        btnAggiungiBookmark.Text = "Salvo Pdf con Divisori"
        btnAggiungiBookmark.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AllowDrop = True
        AutoScaleDimensions = New SizeF(13.0F, 32.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1828, 1267)
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
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtPdfPath As TextBox
    Friend WithEvents dgvReport As DataGridView
    Friend WithEvents picPreview As PictureBox
    Friend WithEvents prbLoad As ProgressBarConTesto
    Friend WithEvents dgvTotali As DataGridView
    Friend WithEvents btnAggiungiBookmark As Button

End Class
