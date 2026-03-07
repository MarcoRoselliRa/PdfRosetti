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
        picOcr = New PictureBox()
        CType(dgvReport, ComponentModel.ISupportInitialize).BeginInit()
        CType(picPreview, ComponentModel.ISupportInitialize).BeginInit()
        CType(picOcr, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' txtPdfPath
        ' 
        txtPdfPath.Location = New Point(66, 22)
        txtPdfPath.Name = "txtPdfPath"
        txtPdfPath.ReadOnly = True
        txtPdfPath.Size = New Size(1144, 39)
        txtPdfPath.TabIndex = 0
        ' 
        ' dgvReport
        ' 
        dgvReport.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvReport.Location = New Point(67, 104)
        dgvReport.Name = "dgvReport"
        dgvReport.RowHeadersWidth = 82
        dgvReport.Size = New Size(1162, 1124)
        dgvReport.TabIndex = 1
        ' 
        ' picPreview
        ' 
        picPreview.BorderStyle = BorderStyle.FixedSingle
        picPreview.Location = New Point(1258, 104)
        picPreview.Name = "picPreview"
        picPreview.Size = New Size(709, 556)
        picPreview.SizeMode = PictureBoxSizeMode.Zoom
        picPreview.TabIndex = 2
        picPreview.TabStop = False
        ' 
        ' picOcr
        ' 
        picOcr.BorderStyle = BorderStyle.FixedSingle
        picOcr.Location = New Point(1258, 672)
        picOcr.Name = "picOcr"
        picOcr.Size = New Size(709, 556)
        picOcr.SizeMode = PictureBoxSizeMode.Zoom
        picOcr.TabIndex = 3
        picOcr.TabStop = False
        ' 
        ' Form1
        ' 
        AllowDrop = True
        AutoScaleDimensions = New SizeF(13F, 32F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1979, 1250)
        Controls.Add(picOcr)
        Controls.Add(picPreview)
        Controls.Add(dgvReport)
        Controls.Add(txtPdfPath)
        Name = "Form1"
        Text = "Form1"
        CType(dgvReport, ComponentModel.ISupportInitialize).EndInit()
        CType(picPreview, ComponentModel.ISupportInitialize).EndInit()
        CType(picOcr, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents txtPdfPath As TextBox
    Friend WithEvents dgvReport As DataGridView
    Friend WithEvents picPreview As PictureBox
    Friend WithEvents picOcr As PictureBox

End Class
