Imports System.IO
Imports System.Drawing.Imaging
Imports PdfiumViewer
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Geom
Imports iText.Kernel.Font
Imports iText.IO.Font.Constants
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties
Imports ZXing
Imports ZXing.Common

Imports PdfiumDoc = PdfiumViewer.PdfDocument
Imports iTextRect = iText.Kernel.Geom.Rectangle

Public Class Form1

    Private PdfPath As String = ""
    Private CurrentPdf As PdfiumDoc = Nothing
    Private CurrentPreviewBmp As Bitmap = Nothing
    Private _indiceBookmark As List(Of (Numero As String, Descrizione As String))

    ' Tolleranza in mm per il riconoscimento formato pagina
    Private Const FORMAT_TOLERANCE_MM As Double = 5.0

    ' Formati standard ISO (larghezza x altezza in mm, sempre Portrait)
    Private Shared ReadOnly IsoFormats As (Name As String, W As Double, H As Double)() = {
        ("A0", 841, 1189),
        ("A1", 594, 841),
        ("A2", 420, 594),
        ("A3", 297, 420),
        ("A4", 210, 297),
        ("A5", 148, 210)
    }

    ' ══════════════════════════════════════════════════════════
    '  INIZIALIZZAZIONE
    ' ══════════════════════════════════════════════════════════

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        If Not String.IsNullOrWhiteSpace(My.Settings.IndicePath) AndAlso File.Exists(My.Settings.IndicePath) Then
            txtIndicePath.Text = My.Settings.IndicePath
            _indiceBookmark = CaricaIndice(My.Settings.IndicePath)
        End If
        SetupGrid()
        AggiornaTesto(0, 0)
        CaricaImpostazioniCarta()
        ' Ripristina percorsi stampanti
        If Not String.IsNullOrWhiteSpace(My.Settings.PercorsoStampante1) Then
            txtStampante1.Text = My.Settings.PercorsoStampante1
        End If
        If Not String.IsNullOrWhiteSpace(My.Settings.PercorsoStampante2) Then
            txtStampante2.Text = My.Settings.PercorsoStampante2
        End If
        ' Ripristina checkbox bookmark
        chkMostraVociSenzaDivisorio.Checked = My.Settings.MostraVociSenzaDivisorio
        chkPuntaDopoDiv.Checked = My.Settings.PuntaDopoDiv
    End Sub

    Private Sub SetupGrid()
        dgvReport.AllowUserToAddRows = False
        dgvReport.AllowUserToDeleteRows = False
        dgvReport.ReadOnly = True
        dgvReport.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvReport.MultiSelect = False
        dgvReport.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvReport.Columns.Clear()
        dgvReport.Columns.Add("colPagina", "Pagina")
        dgvReport.Columns.Add("colDimensione", "Dimensione")
        dgvReport.Columns.Add("colBNColore", "BN/Colore")
        dgvReport.Columns.Add("colOrientamento", "Orientamento")
        dgvReport.Columns.Add("colDivisorio", "Divisorio")

        Dim centrata As New DataGridViewCellStyle()
        centrata.Alignment = DataGridViewContentAlignment.MiddleCenter
        For Each col As DataGridViewColumn In dgvReport.Columns
            If col.Name <> "colPagina" Then col.DefaultCellStyle = centrata
        Next

        SetupGridTotali()
    End Sub

    Private Sub SetupGridTotali()
        dgvTotali.AllowUserToAddRows = False
        dgvTotali.AllowUserToDeleteRows = False
        dgvTotali.ReadOnly = True
        dgvTotali.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvTotali.MultiSelect = False
        dgvTotali.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvTotali.Columns.Clear()
        dgvTotali.Columns.Add("colDescrizione", "Descrizione")
        dgvTotali.Columns.Add("colQuantita", "Quantità")
        dgvTotali.Columns("colQuantita").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  PROGRESSBAR CON TESTO CENTRALE
    ' ══════════════════════════════════════════════════════════

    Private Sub AggiornaTesto(pagina As Integer, totale As Integer)
        If totale = 0 Then
            prbLoad.Value = 0
            prbLoad.Testo = ""
        Else
            prbLoad.Maximum = totale
            prbLoad.Value = Math.Min(pagina, totale)
            prbLoad.Testo = $"Pagina {pagina} di {totale}"
        End If
        prbLoad.Invalidate()
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  DRAG & DROP
    ' ══════════════════════════════════════════════════════════

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length > 0 AndAlso String.Equals(System.IO.Path.GetExtension(files(0)), ".pdf", StringComparison.OrdinalIgnoreCase) Then
                e.Effect = DragDropEffects.Copy
            End If
        End If
    End Sub

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length = 0 Then Exit Sub
        CaricaPdf(files(0))
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  CARICAMENTO PDF
    ' ══════════════════════════════════════════════════════════

    Private Sub CaricaPdf(path As String)
        If Not File.Exists(path) Then Exit Sub

        ' Resetta tutto e disabilita i controlli durante l'elaborazione
        ResetTutto()
        AbilitaControlli(False)

        PdfPath = path
        txtPdfPath.Text = System.IO.Path.GetFileName(PdfPath)

        If CurrentPdf IsNot Nothing Then
            CurrentPdf.Dispose()
            CurrentPdf = Nothing
        End If

        CurrentPdf = PdfiumDoc.Load(PdfPath)
        CaricaPagineInGriglia()
        AnalizzaTuttoAsync()
    End Sub

    Private Sub CaricaPagineInGriglia()
        dgvReport.Rows.Clear()

        For i As Integer = 0 To CurrentPdf.PageCount - 1
            Dim sizePt = CurrentPdf.PageSizes(i)
            Dim wMm As Double = Math.Round(PtToMm(sizePt.Width), 1)
            Dim hMm As Double = Math.Round(PtToMm(sizePt.Height), 1)
            Dim orientamento As String = If(hMm >= wMm, "Portrait", "Landscape")
            Dim dimensione As String = RiconosciFomato(wMm, hMm)
            dgvReport.Rows.Add((i + 1).ToString(), dimensione, "", orientamento, "")
        Next
    End Sub

    Private Function PtToMm(pt As Double) As Double
        Return pt * 25.4 / 72.0
    End Function

    Private Function RiconosciFomato(wMm As Double, hMm As Double) As String
        Dim shortSide As Double = Math.Min(wMm, hMm)
        Dim longSide As Double = Math.Max(wMm, hMm)

        For Each fmt In IsoFormats
            Dim fmtShort As Double = Math.Min(fmt.W, fmt.H)
            Dim fmtLong As Double = Math.Max(fmt.W, fmt.H)
            If Math.Abs(shortSide - fmtShort) <= FORMAT_TOLERANCE_MM AndAlso
               Math.Abs(longSide - fmtLong) <= FORMAT_TOLERANCE_MM Then
                Return fmt.Name
            End If
        Next

        Return $"{wMm} x {hMm} mm"
    End Function

    ' ══════════════════════════════════════════════════════════
    '  ANALISI AUTOMATICA PAGINA PER PAGINA
    ' ══════════════════════════════════════════════════════════

    Private Async Sub AnalizzaTuttoAsync()

        Dim totale As Integer = CurrentPdf.PageCount
        AggiornaTesto(0, totale)

        Await Task.Run(Sub()

                           Using pdf As PdfiumDoc = PdfiumDoc.Load(PdfPath)

                               For i As Integer = 0 To pdf.PageCount - 1
                                   Dim rowIndex As Integer = i
                                   Dim risultatoColore As String = "BN"
                                   Dim risultatoDivisorio As String = ""

                                   Try
                                       Dim sz = pdf.PageSizes(rowIndex)
                                       Dim maxSidePx As Integer = 1200
                                       Dim scale As Double = maxSidePx / CDbl(Math.Max(sz.Width, sz.Height))
                                       Dim w As Integer = Math.Max(1, CInt(sz.Width * scale))
                                       Dim h As Integer = Math.Max(1, CInt(sz.Height * scale))

                                       Using bmp As Bitmap = pdf.Render(rowIndex, w, h, 120, 120, PdfRenderFlags.Annotations)
                                           ' Analisi colore
                                           risultatoColore = If(IsColorPixel(bmp), "Colore", "BN")

                                           ' Lettura barcode Code128
                                           risultatoDivisorio = LeggiBarcode(bmp)

                                       End Using





                                   Catch ex As Exception
                                       risultatoColore = "ERR"
                                       Debug.WriteLine($"Errore pagina {rowIndex}: {ex.Message}")
                                   End Try

                                   Me.BeginInvoke(Sub()
                                                      If rowIndex < dgvReport.Rows.Count Then
                                                          dgvReport.Rows(rowIndex).Cells("colBNColore").Value = risultatoColore
                                                          dgvReport.Rows(rowIndex).Cells("colDivisorio").Value = risultatoDivisorio
                                                          AggiornaTesto(rowIndex + 1, totale)
                                                      End If
                                                  End Sub)

                               Next

                           End Using

                       End Sub)

        If dgvReport.Rows.Count > 0 Then
            dgvReport.ClearSelection()
            dgvReport.Rows(0).Selected = True
            dgvReport.FirstDisplayedScrollingRowIndex = 0
        End If
        AggiornaTesto(0, 0)
        AggiornaTotali()
        AbilitaControlli(True)

    End Sub

    ' ══════════════════════════════════════════════════════════
    '  LETTURA BARCODE
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Legge il barcode Code128 dalla pagina renderizzata.
    ''' Prova prima sulla pagina intera, poi ritaglia la fascia
    ''' sinistra in basso dove è posizionato il barcode del divisorio.
    ''' Restituisce il testo del barcode o stringa vuota se non trovato.
    ''' </summary>
    Private Function LeggiBarcode(bmp As Bitmap) As String
        Try
            Dim reader As New ZXing.BarcodeReaderGeneric()
            reader.Options = New DecodingOptions()
            reader.Options.PossibleFormats = New List(Of BarcodeFormat) From {BarcodeFormat.CODE_128}
            reader.Options.TryHarder = True

            ' Prova sulla pagina intera
            Dim bytes = BitmapToRgbBytes(bmp)
            Dim source As New ZXing.RGBLuminanceSource(bytes, bmp.Width, bmp.Height,
                                                    ZXing.RGBLuminanceSource.BitmapFormat.RGB24)
            Dim result = reader.Decode(source)
            If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.Text) Then
                Return result.Text.Trim()
            End If

            ' Ritaglia fascia sinistra in basso con margine bianco
            Dim cropW As Integer = CInt(bmp.Width * 0.65)
            Dim cropH As Integer = CInt(bmp.Height * 0.12)
            Dim cropY As Integer = bmp.Height - cropH - CInt(bmp.Height * 0.01)
            Dim pad As Integer = 20

            Using padded As New Bitmap(cropW + pad * 2, cropH + pad * 2)
                Using g As Graphics = Graphics.FromImage(padded)
                    g.Clear(Color.White)
                    g.DrawImage(bmp, New System.Drawing.Rectangle(pad, pad, cropW, cropH),
                            New System.Drawing.Rectangle(0, cropY, cropW, cropH), GraphicsUnit.Pixel)
                End Using

                Dim bytesCrop = BitmapToRgbBytes(padded)
                Dim sourceCrop As New ZXing.RGBLuminanceSource(bytesCrop, padded.Width, padded.Height,
                                                            ZXing.RGBLuminanceSource.BitmapFormat.RGB24)
                result = reader.Decode(sourceCrop)
                If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.Text) Then
                    Return result.Text.Trim()
                End If
            End Using

        Catch ex As Exception
            Debug.WriteLine($"Errore lettura barcode: {ex.Message}")
        End Try

        Return ""
    End Function

    Private Function BitmapToRgbBytes(bmp As Bitmap) As Byte()
        Dim bmp24 As Bitmap
        Dim created As Boolean = False

        If bmp.PixelFormat <> Imaging.PixelFormat.Format24bppRgb Then
            bmp24 = New Bitmap(bmp.Width, bmp.Height, Imaging.PixelFormat.Format24bppRgb)
            Using g As Graphics = Graphics.FromImage(bmp24)
                g.DrawImage(bmp, 0, 0)
            End Using
            created = True
        Else
            bmp24 = bmp
        End If

        Try
            Dim data = bmp24.LockBits(New System.Drawing.Rectangle(0, 0, bmp24.Width, bmp24.Height),
                                   Imaging.ImageLockMode.ReadOnly,
                                   Imaging.PixelFormat.Format24bppRgb)
            Dim bytes(data.Stride * bmp24.Height - 1) As Byte
            System.Runtime.InteropServices.Marshal.Copy(data.Scan0, bytes, 0, bytes.Length)
            bmp24.UnlockBits(data)
            Return bytes
        Finally
            If created Then bmp24.Dispose()
        End Try
    End Function

    ' ══════════════════════════════════════════════════════════
    '  TOTALI
    ' ══════════════════════════════════════════════════════════

    Private Sub AggiornaTotali()
        dgvTotali.Rows.Clear()

        Dim conteggi As New Dictionary(Of String, Integer)
        Dim divisori As Integer = 0
        Dim fuoriFormato As Integer = 0

        Dim formatiNoti As New HashSet(Of String)({"A0", "A1", "A2", "A3", "A4", "A5"})

        For Each row As DataGridViewRow In dgvReport.Rows
            Dim dimensione As String = If(row.Cells("colDimensione").Value?.ToString(), "")
            Dim colore As String = If(row.Cells("colBNColore").Value?.ToString(), "")
            Dim divisorio As String = If(row.Cells("colDivisorio").Value?.ToString(), "")

            If divisorio <> "" Then
                divisori += 1
                Continue For
            End If

            If Not formatiNoti.Contains(dimensione) Then
                fuoriFormato += 1
                Continue For
            End If

            If colore = "BN" OrElse colore = "Colore" Then
                Dim chiave As String = $"{dimensione} {colore}"
                If conteggi.ContainsKey(chiave) Then
                    conteggi(chiave) += 1
                Else
                    conteggi(chiave) = 1
                End If
            End If
        Next

        Dim ordineFormati() As String = {"A0", "A1", "A2", "A3", "A4", "A5"}
        Dim ordineColore() As String = {"Colore", "BN"}

        For Each fmt In ordineFormati
            For Each col In ordineColore
                Dim chiave As String = $"{fmt} {col}"
                If conteggi.ContainsKey(chiave) AndAlso conteggi(chiave) > 0 Then
                    dgvTotali.Rows.Add(chiave, conteggi(chiave))
                End If
            Next
        Next

        If divisori > 0 Then dgvTotali.Rows.Add("Divisori", divisori)
        If fuoriFormato > 0 Then dgvTotali.Rows.Add("Fuori formato", fuoriFormato)
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  ANALISI COLORE
    ' ══════════════════════════════════════════════════════════

    Private Function IsColorPixel(bmp As Bitmap) As Boolean
        For y As Integer = 0 To bmp.Height - 1 Step 4
            For x As Integer = 0 To bmp.Width - 1 Step 4
                Dim c As Color = bmp.GetPixel(x, y)
                Dim diff As Integer = Math.Abs(CInt(c.R) - CInt(c.G)) +
                                      Math.Abs(CInt(c.R) - CInt(c.B)) +
                                      Math.Abs(CInt(c.G) - CInt(c.B))
                If diff > 10 Then Return True
            Next
        Next
        Return False
    End Function

    ' ══════════════════════════════════════════════════════════
    '  PREVIEW PAGINA
    ' ══════════════════════════════════════════════════════════

    Private Sub dgvReport_SelectionChanged(sender As Object, e As EventArgs) Handles dgvReport.SelectionChanged
        If CurrentPdf Is Nothing Then Exit Sub
        If dgvReport.SelectedRows.Count = 0 Then Exit Sub
        Dim pageIndex As Integer = dgvReport.SelectedRows(0).Index
        AggiornaAnteprimaPagina(pageIndex)
    End Sub

    Private Sub AggiornaAnteprimaPagina(pageIndex As Integer)
        If CurrentPreviewBmp IsNot Nothing Then
            picPreview.Image = Nothing
            CurrentPreviewBmp.Dispose()
            CurrentPreviewBmp = Nothing
        End If

        Dim maxSidePx As Integer = 1200
        Dim sz = CurrentPdf.PageSizes(pageIndex)
        Dim w As Integer = Math.Max(1, CInt(sz.Width))
        Dim h As Integer = Math.Max(1, CInt(sz.Height))
        Dim scale As Double = maxSidePx / CDbl(Math.Max(w, h))
        Dim outW As Integer = Math.Max(1, CInt(w * scale))
        Dim outH As Integer = Math.Max(1, CInt(h * scale))

        CurrentPreviewBmp = CurrentPdf.Render(pageIndex, outW, outH, 120, 120, PdfRenderFlags.Annotations)
        picPreview.Image = CurrentPreviewBmp
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  CHIUSURA FORM
    ' ══════════════════════════════════════════════════════════

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If CurrentPreviewBmp IsNot Nothing Then CurrentPreviewBmp.Dispose()
        If CurrentPdf IsNot Nothing Then CurrentPdf.Dispose()
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  PERCORSI STAMPANTI
    ' ══════════════════════════════════════════════════════════

    Private Sub btnSfogliaStampante1_Click(sender As Object, e As EventArgs) Handles btnSfogliaStampante1.Click
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Seleziona cartella Stampante 1"
            If Not String.IsNullOrWhiteSpace(My.Settings.PercorsoStampante1) Then
                fbd.SelectedPath = My.Settings.PercorsoStampante1
            End If
            If fbd.ShowDialog() = DialogResult.OK Then
                txtStampante1.Text = fbd.SelectedPath
                My.Settings.PercorsoStampante1 = fbd.SelectedPath
                My.Settings.Save()
            End If
        End Using
    End Sub

    Private Sub btnSfogliaStampante2_Click(sender As Object, e As EventArgs) Handles btnSfogliaStampante2.Click
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Seleziona cartella Stampante 2"
            If Not String.IsNullOrWhiteSpace(My.Settings.PercorsoStampante2) Then
                fbd.SelectedPath = My.Settings.PercorsoStampante2
            End If
            If fbd.ShowDialog() = DialogResult.OK Then
                txtStampante2.Text = fbd.SelectedPath
                My.Settings.PercorsoStampante2 = fbd.SelectedPath
                My.Settings.Save()
            End If
        End Using
    End Sub

    Private Sub btnInviaStampante1_Click(sender As Object, e As EventArgs) Handles btnInviaStampante1.Click
        InviaAStampante(txtStampante1.Text.Trim(), "Stampante 1")
    End Sub

    Private Sub btnInviaStampante2_Click(sender As Object, e As EventArgs) Handles btnInviaStampante2.Click
        InviaAStampante(txtStampante2.Text.Trim(), "Stampante 2")
    End Sub

    Private Sub InviaAStampante(percorso As String, nomePrinter As String)
        If String.IsNullOrWhiteSpace(PdfPath) Then
            MessageBox.Show("Carica prima un PDF.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrWhiteSpace(percorso) Then
            MessageBox.Show($"Percorso {nomePrinter} non impostato. Usa il pulsante Sfoglia.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If Not Directory.Exists(percorso) Then
            MessageBox.Show($"La cartella {nomePrinter} non esiste:{Environment.NewLine}{percorso}", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Genera PDF di stampa e XPF
        If Not GeneraFileDiStampa() Then Return

        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(PdfPath) & "_print"
        Dim printPdf As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(PdfPath), baseName & ".pdf")
        Dim printXpf As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(PdfPath), baseName & ".pdf.xpf")

        Try
            Dim destPdf As String = System.IO.Path.Combine(percorso, System.IO.Path.GetFileName(printPdf))
            Dim destXpf As String = System.IO.Path.Combine(percorso, System.IO.Path.GetFileName(printXpf))

            ' Copia solo se origine e destinazione sono diverse
            If Not String.Equals(printPdf, destPdf, StringComparison.OrdinalIgnoreCase) Then
                File.Copy(printPdf, destPdf, True)
            End If
            If Not String.Equals(printXpf, destXpf, StringComparison.OrdinalIgnoreCase) Then
                File.Copy(printXpf, destXpf, True)
            End If

            ' Cancella i file temporanei dalla cartella originale
            If File.Exists(printPdf) Then File.Delete(printPdf)
            If File.Exists(printXpf) Then File.Delete(printXpf)

            MessageBox.Show($"File inviati a {nomePrinter}:{Environment.NewLine}{System.IO.Path.GetFileName(printPdf)}{Environment.NewLine}{System.IO.Path.GetFileName(printXpf)}",
                            "Completato", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"Errore durante l'invio a {nomePrinter}:{Environment.NewLine}{ex.Message}",
                            "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  CREAZIONE PAGINA DI COPERTINA
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Crea un PDF temporaneo con la pagina di copertina (nome file, data/ora, totali).
    ''' </summary>
    Private Function CreaCopertura(cartellaDest As String) As String

        Dim coverPath As String = System.IO.Path.Combine(cartellaDest, "_copertina_tmp.pdf")

        Using writer As New PdfWriter(coverPath)
            Using pdfDoc As New iText.Kernel.Pdf.PdfDocument(writer)
                Using doc As New Document(pdfDoc, PageSize.A4)

                    doc.SetMargins(60, 50, 60, 50)

                    Dim fontBold = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD)
                    Dim fontNormal = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)

                    ' ── Nome file ──────────────────────────────
                    Dim titolo As New Paragraph(System.IO.Path.GetFileName(PdfPath))
                    titolo.SetFont(fontBold).SetFontSize(20)
                    titolo.SetTextAlignment(TextAlignment.CENTER)
                    titolo.SetMarginBottom(10)
                    doc.Add(titolo)

                    ' ── Data e ora ─────────────────────────────
                    Dim dataOra As New Paragraph(DateTime.Now.ToString("dd/MM/yyyy HH:mm"))
                    dataOra.SetFont(fontNormal).SetFontSize(12)
                    dataOra.SetTextAlignment(TextAlignment.CENTER)
                    dataOra.SetMarginBottom(40)
                    doc.Add(dataOra)

                    ' ── Separatore ─────────────────────────────
                    Dim sep As New Paragraph("─────────────────────────────────────────")
                    sep.SetFont(fontNormal).SetFontSize(10)
                    sep.SetTextAlignment(TextAlignment.CENTER)
                    sep.SetMarginBottom(30)
                    doc.Add(sep)

                    ' ── Titolo tabella totali ──────────────────
                    Dim titoloTot As New Paragraph("RIEPILOGO FOGLI")
                    titoloTot.SetFont(fontBold).SetFontSize(14)
                    titoloTot.SetTextAlignment(TextAlignment.CENTER)
                    titoloTot.SetMarginBottom(20)
                    doc.Add(titoloTot)

                    ' ── Tabella totali da dgvTotali ────────────
                    Dim table As New Table(2)
                    table.SetWidth(UnitValue.CreatePercentValue(60))
                    table.SetHorizontalAlignment(HorizontalAlignment.CENTER)

                    ' Intestazione tabella
                    Dim hDesc As New Cell()
                    hDesc.Add(New Paragraph("Descrizione").SetFont(fontBold).SetFontSize(11))
                    hDesc.SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
                    hDesc.SetPadding(6)
                    table.AddHeaderCell(hDesc)

                    Dim hQta As New Cell()
                    hQta.Add(New Paragraph("Quantità").SetFont(fontBold).SetFontSize(11))
                    hQta.SetBackgroundColor(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY)
                    hQta.SetPadding(6)
                    hQta.SetTextAlignment(TextAlignment.RIGHT)
                    table.AddHeaderCell(hQta)

                    ' Righe dalla griglia totali
                    For Each row As DataGridViewRow In dgvTotali.Rows
                        Dim desc As String = If(row.Cells("colDescrizione").Value?.ToString(), "")
                        Dim qta As String = If(row.Cells("colQuantita").Value?.ToString(), "")

                        Dim cDesc As New Cell()
                        cDesc.Add(New Paragraph(desc).SetFont(fontNormal).SetFontSize(11))
                        cDesc.SetPadding(5)
                        table.AddCell(cDesc)

                        Dim cQta As New Cell()
                        cQta.Add(New Paragraph(qta).SetFont(fontNormal).SetFontSize(11))
                        cQta.SetPadding(5)
                        cQta.SetTextAlignment(TextAlignment.RIGHT)
                        table.AddCell(cQta)
                    Next

                    ' Riga totale pagine
                    Dim totPagine As Integer = dgvReport.Rows.Count
                    Dim cTotDesc As New Cell()
                    cTotDesc.Add(New Paragraph("TOTALE PAGINE").SetFont(fontBold).SetFontSize(11))
                    cTotDesc.SetPadding(5)
                    table.AddCell(cTotDesc)

                    Dim cTotQta As New Cell()
                    cTotQta.Add(New Paragraph(totPagine.ToString()).SetFont(fontBold).SetFontSize(11))
                    cTotQta.SetPadding(5)
                    cTotQta.SetTextAlignment(TextAlignment.RIGHT)
                    table.AddCell(cTotQta)

                    doc.Add(table)

                End Using
            End Using
        End Using

        Return coverPath
    End Function

    ' ══════════════════════════════════════════════════════════
    '  CREAZIONE PDF DI STAMPA (divisori allargati a 223mm)
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Crea un nuovo PDF dove le pagine divisorio vengono allargate da 210mm a 223mm
    ''' aggiungendo un margine bianco a sinistra, così l'immagine risulta allineata a destra.
    ''' Le altre pagine vengono copiate invariate.
    ''' Restituisce il path del nuovo PDF creato.
    ''' </summary>
    Private Function CreaPdfStampa(pagineDivisorio As HashSet(Of Integer)) As String

        Dim outPdfPath As String = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(PdfPath),
            System.IO.Path.GetFileNameWithoutExtension(PdfPath) & "_print.pdf")

        Dim cartella As String = System.IO.Path.GetDirectoryName(PdfPath)

        ' Larghezza extra in punti: 15mm * 72/25.4
        Dim extraPt As Single = CSng(15.0 * 72.0 / 25.4)

        ' Crea prima la copertina come PDF temporaneo
        Dim coverPath As String = CreaCopertura(cartella)

        Try
            Using writer As New PdfWriter(outPdfPath)
                Using dstDoc As New iText.Kernel.Pdf.PdfDocument(writer)

                    ' ── Pagina 1: copertina ───────────────────
                    Using coverReader As New PdfReader(coverPath)
                        Using coverDoc As New iText.Kernel.Pdf.PdfDocument(coverReader)
                            coverDoc.CopyPagesTo(1, 1, dstDoc)
                        End Using
                    End Using

                    ' ── Pagine successive: PDF originale ──────
                    Using srcReader As New PdfReader(PdfPath)
                        Using srcDoc As New iText.Kernel.Pdf.PdfDocument(srcReader)
                            For i As Integer = 1 To srcDoc.GetNumberOfPages()
                                Dim srcPage = srcDoc.GetPage(i)
                                Dim srcSize = srcPage.GetPageSize()

                                If pagineDivisorio.Contains(i) Then
                                    ' Pagina divisorio: allarga a sinistra
                                    Dim nuovaLarghezza As Single = srcSize.GetWidth() + extraPt
                                    Dim nuovaSize As New iText.Kernel.Geom.Rectangle(nuovaLarghezza, srcSize.GetHeight())
                                    Dim dstPage = dstDoc.AddNewPage(New PageSize(nuovaSize))
                                    Dim canvas As New PdfCanvas(dstPage)
                                    Dim xobj = srcPage.CopyAsFormXObject(dstDoc)
                                    canvas.AddXObjectAt(xobj, extraPt, 0)
                                    canvas.Release()
                                Else
                                    srcDoc.CopyPagesTo(i, i, dstDoc)
                                End If
                            Next
                        End Using
                    End Using

                End Using
            End Using

        Finally
            ' Elimina il PDF copertina temporaneo
            If File.Exists(coverPath) Then File.Delete(coverPath)
        End Try

        Return outPdfPath

    End Function

    ''' <summary>
    ''' Carica i nomi carta salvati nei settings nelle TextBox del pannello.
    ''' Le TextBox devono chiamarsi: txtCartaA4, txtCartaA3, txtCartaDivisorio
    ''' </summary>
    Private Sub CaricaImpostazioniCarta()
        txtCartaA4.Text = If(String.IsNullOrWhiteSpace(My.Settings.NomeCartaA4),
                             "Rosetti 70gr A4 LEF Bianco", My.Settings.NomeCartaA4)
        txtCartaA3.Text = If(String.IsNullOrWhiteSpace(My.Settings.NomeCartaA3),
                             "Rosetti 70gr A3 Bianco", My.Settings.NomeCartaA3)
        txtCartaDivisorio.Text = If(String.IsNullOrWhiteSpace(My.Settings.NomeCartaDivisorio),
                                    "Rosetti 140gr 223x297mm LEF Verde", My.Settings.NomeCartaDivisorio)
    End Sub

    Private Sub btnSalvaImpostazioni_Click(sender As Object, e As EventArgs) Handles btnSalvaImpostazioni.Click
        If String.IsNullOrWhiteSpace(txtCartaA4.Text) OrElse
           String.IsNullOrWhiteSpace(txtCartaA3.Text) OrElse
           String.IsNullOrWhiteSpace(txtCartaDivisorio.Text) Then
            MessageBox.Show("I nomi carta non possono essere vuoti.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        My.Settings.NomeCartaA4 = txtCartaA4.Text.Trim()
        My.Settings.NomeCartaA3 = txtCartaA3.Text.Trim()
        My.Settings.NomeCartaDivisorio = txtCartaDivisorio.Text.Trim()
        My.Settings.Save()
        MessageBox.Show("Impostazioni salvate.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  GENERAZIONE XPIF JOB TICKET
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Classifica il tipo di pagina in base a dimensione, colore e presenza barcode divisorio.
    ''' </summary>
    Private Enum TipoPagina
        A4Colore    ' default assoluto — nessuna eccezione necessaria
        A4BN        ' eccezione: solo colore → monochrome-grayscale
        A3Colore    ' eccezione: carta A3, color-effects-type vuoto (eredita default)
        A3BN        ' eccezione: carta A3 + monochrome-grayscale
        Divisorio   ' eccezione: carta divisorio, color-effects-type vuoto (eredita)
    End Enum

    Private Function ClassificaPagina(dimensione As String, colore As String, isDivisorio As Boolean) As TipoPagina
        If isDivisorio Then Return TipoPagina.Divisorio
        Dim isGrande As Boolean = (dimensione = "A3" OrElse dimensione = "A2" OrElse
                                   dimensione = "A1" OrElse dimensione = "A0")
        If isGrande Then
            Return If(colore = "BN", TipoPagina.A3BN, TipoPagina.A3Colore)
        End If
        Return If(colore = "BN", TipoPagina.A4BN, TipoPagina.A4Colore)
    End Function

    ''' <summary>
    ''' Scrive il blocco media-col completo nel formato FreeFlow Core.
    ''' </summary>
    Private Sub AppendMediaCol(sb As System.Text.StringBuilder,
                               indent As String,
                               mediaKey As String,
                               mediaColor As String,
                               xDim As Integer,
                               yDim As Integer)
        sb.AppendLine($"{indent}<media-col syntax=""collection"">")
        sb.AppendLine($"{indent}  <media-color syntax=""keyword"">{mediaColor}</media-color>")
        sb.AppendLine($"{indent}  <media-key syntax=""name"" xml:space=""preserve"">{mediaKey}</media-key>")
        sb.AppendLine($"{indent}  <media-order-count syntax=""integer"">1</media-order-count>")
        sb.AppendLine($"{indent}  <media-size syntax=""collection"">")
        sb.AppendLine($"{indent}    <x-dimension syntax=""integer"">{xDim}</x-dimension>")
        sb.AppendLine($"{indent}    <y-dimension syntax=""integer"">{yDim}</y-dimension>")
        sb.AppendLine($"{indent}  </media-size>")
        sb.AppendLine($"{indent}  <media-type syntax=""keyword"">stationery</media-type>")
        sb.AppendLine($"{indent}</media-col>")
    End Sub

    ''' <summary>
    ''' Genera il PDF di stampa e il file XPF. Restituisce True se tutto ok.
    ''' </summary>
    Private Function GeneraFileDiStampa() As Boolean

        If String.IsNullOrEmpty(PdfPath) Then
            MessageBox.Show("Carica prima un PDF.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("Nessuna pagina in griglia.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        Dim cartaA4 As String = txtCartaA4.Text.Trim()
        Dim cartaA3 As String = txtCartaA3.Text.Trim()
        Dim cartaDiv As String = txtCartaDivisorio.Text.Trim()

        If String.IsNullOrWhiteSpace(cartaA4) OrElse
           String.IsNullOrWhiteSpace(cartaA3) OrElse
           String.IsNullOrWhiteSpace(cartaDiv) Then
            MessageBox.Show("Verifica i nomi carta nelle impostazioni prima di generare l'XPIF.",
                            "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If

        ' ── Raccoglie le pagine divisorio ─────────────────────
        Dim pagineDivisorio As New HashSet(Of Integer)
        Dim eccezioni As New List(Of (DaP As Integer, AP As Integer, Tipo As TipoPagina))

        For Each row As DataGridViewRow In dgvReport.Rows
            Dim pageNum As Integer = row.Index + 2  ' +1 per la copertina in testa
            Dim dimensione As String = If(row.Cells("colDimensione").Value?.ToString(), "A4")
            Dim colore As String = If(row.Cells("colBNColore").Value?.ToString(), "Colore")
            Dim divisorio As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            Dim isDivisorio As Boolean = (divisorio <> "" AndAlso divisorio <> "ERR")

            If isDivisorio Then pagineDivisorio.Add(row.Index + 1)  ' nel PDF originale parte da 1

            Dim tipo As TipoPagina = ClassificaPagina(dimensione, colore, isDivisorio)
            If tipo = TipoPagina.A4Colore Then Continue For

            If eccezioni.Count > 0 Then
                Dim ultimo = eccezioni(eccezioni.Count - 1)
                If ultimo.Tipo = tipo AndAlso ultimo.AP = pageNum - 1 Then
                    eccezioni(eccezioni.Count - 1) = (ultimo.DaP, pageNum, tipo)
                    Continue For
                End If
            End If
            eccezioni.Add((pageNum, pageNum, tipo))
        Next

        Try
            ' ── Step 1: crea il PDF di stampa con divisori allargati ─
            Dim printPdfPath As String = CreaPdfStampa(pagineDivisorio)
            Dim printPdfName As String = System.IO.Path.GetFileName(printPdfPath)

            ' ── Step 2: genera il file XPF ───────────────────────────
            Dim outPath As String = System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(PdfPath),
                System.IO.Path.GetFileNameWithoutExtension(PdfPath) & "_print.pdf.xpf")

            Dim sb As New System.Text.StringBuilder()
            Dim jobName As String = printPdfName
            Dim totalPages As Integer = dgvReport.Rows.Count + 1  ' +1 per la copertina

            sb.AppendLine("<?xml version=""1.0"" encoding=""utf-8""?>")
            sb.AppendLine("<!DOCTYPE xpif SYSTEM ""xpif-v02080.dtd"">")
            sb.AppendLine("<xpif version=""1.0"" cpss-version=""2.07"" xml:lang=""en-US"">")
            sb.AppendLine("  <xpif-operation-attributes>")
            sb.AppendLine("    <document-format syntax=""mimeMediaType"">application/pdf</document-format>")
            sb.AppendLine($"    <job-name syntax=""name"" xml:space=""preserve"">{jobName}</job-name>")
            sb.AppendLine($"    <job-pages syntax=""integer"">{totalPages}</job-pages>")
            sb.AppendLine("    <requesting-user-name syntax=""name"" xml:space=""preserve"">FreeFlow Core</requesting-user-name>")
            sb.AppendLine("  </xpif-operation-attributes>")
            sb.AppendLine("  <job-template-attributes>")
            sb.AppendLine($"    <copies syntax=""integer"">{nudCopie.Value}</copies>")
            sb.AppendLine("    <finishings syntax=""1setOf"">")
            sb.AppendLine("      <value syntax=""enum"">0</value>")
            sb.AppendLine("      <value syntax=""enum"">92</value>")
            sb.AppendLine("      <value syntax=""enum"">93</value>")
            sb.AppendLine("      <value syntax=""enum"">1011</value>")
            sb.AppendLine("    </finishings>")
            sb.AppendLine("    <imposition-source-orientation syntax=""keyword"">portrait</imposition-source-orientation>")
            sb.AppendLine("    <job-offset syntax=""1setOf"">")
            sb.AppendLine("      <value syntax=""keyword"">none</value>")
            sb.AppendLine("    </job-offset>")
            AppendMediaCol(sb, "    ", cartaA4, "white", 21000, 29700)
            sb.AppendLine("    <orientation-requested syntax=""enum"">3</orientation-requested>")
            sb.AppendLine("    <output-bin syntax=""keyword"">stacker-2</output-bin>")

            If eccezioni.Count > 0 Then
                sb.AppendLine("    <page-overrides syntax=""1setOf"">")
                For Each exc In eccezioni
                    Dim nomeCarta As String
                    Dim coloreCarta As String
                    Dim xDim As Integer
                    Dim yDim As Integer
                    Dim colorEffect As String

                    Select Case exc.Tipo
                        Case TipoPagina.A3Colore
                            nomeCarta = cartaA3 : coloreCarta = "white"
                            xDim = 29700 : yDim = 42000
                            colorEffect = ""
                        Case TipoPagina.A3BN
                            nomeCarta = cartaA3 : coloreCarta = "white"
                            xDim = 29700 : yDim = 42000
                            colorEffect = "monochrome-grayscale"
                        Case TipoPagina.Divisorio
                            nomeCarta = cartaDiv : coloreCarta = "green"
                            xDim = 22300 : yDim = 29700
                            colorEffect = ""
                        Case Else ' A4BN
                            nomeCarta = cartaA4 : coloreCarta = "white"
                            xDim = 21000 : yDim = 29700
                            colorEffect = "monochrome-grayscale"
                    End Select

                    sb.AppendLine("      <value syntax=""collection"">")
                    sb.AppendLine($"        <color-effects-type syntax=""keyword"">{colorEffect}</color-effects-type>")
                    sb.AppendLine("        <input-documents syntax=""1setOf"">")
                    sb.AppendLine("          <value syntax=""rangeOfInteger"">")
                    sb.AppendLine("            <lower-bound syntax=""integer"">1</lower-bound>")
                    sb.AppendLine("            <upper-bound syntax=""integer"">1</upper-bound>")
                    sb.AppendLine("          </value>")
                    sb.AppendLine("        </input-documents>")
                    AppendMediaCol(sb, "        ", nomeCarta, coloreCarta, xDim, yDim)
                    sb.AppendLine("        <pages syntax=""1setOf"">")
                    sb.AppendLine("          <value syntax=""rangeOfInteger"">")
                    sb.AppendLine($"            <lower-bound syntax=""integer"">{exc.DaP}</lower-bound>")
                    sb.AppendLine($"            <upper-bound syntax=""integer"">{exc.AP}</upper-bound>")
                    sb.AppendLine("          </value>")
                    sb.AppendLine("        </pages>")
                    sb.AppendLine("        <sides syntax=""keyword""></sides>")
                    sb.AppendLine("      </value>")
                Next
                sb.AppendLine("    </page-overrides>")
            End If

            sb.AppendLine("  </job-template-attributes>")
            sb.AppendLine("</xpif>")

            File.WriteAllText(outPath, sb.ToString(), New System.Text.UTF8Encoding(False))
            Return True

        Catch ex As Exception
            MessageBox.Show("Errore durante la generazione:" &
                            Environment.NewLine & ex.ToString(),
                            "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    ' ══════════════════════════════════════════════════════════
    '  BOOKMARK PDF
    ' ══════════════════════════════════════════════════════════

    Private Sub btnAggiungiBookmark_Click(sender As Object, e As EventArgs) Handles btnAggiungiBookmark.Click

        If String.IsNullOrEmpty(PdfPath) Then
            MessageBox.Show("Carica prima un PDF.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If String.IsNullOrEmpty(txtIndicePath.Text) OrElse Not File.Exists(txtIndicePath.Text) Then
            MessageBox.Show("Seleziona prima il file indice.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim indice = CaricaIndice(txtIndicePath.Text)
        If indice.Count = 0 Then
            MessageBox.Show("Il file indice è vuoto o non valido.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Legge i flag dall'interfaccia
        Dim flagMostraVociSenzaDivisorio As Boolean = chkMostraVociSenzaDivisorio.Checked
        Dim flagPuntaDopoDiv As Boolean = chkPuntaDopoDiv.Checked

        ' Costruisce dizionario divisorio → numero pagina
        Dim divisoriPagina As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataGridViewRow In dgvReport.Rows
            Dim div As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            If div <> "" AndAlso div <> "ERR" Then
                If Not divisoriPagina.ContainsKey(div) Then
                    divisoriPagina.Add(div, row.Index + 1)
                End If
            End If
        Next

        ' Costruisce set delle pagine divisorio (per sapere se la pagina successiva è divisorio)
        Dim pagineDivisorio As New HashSet(Of Integer)
        For Each row As DataGridViewRow In dgvReport.Rows
            Dim div As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            If div <> "" AndAlso div <> "ERR" Then
                pagineDivisorio.Add(row.Index + 1)
            End If
        Next

        If divisoriPagina.Count = 0 AndAlso Not flagMostraVociSenzaDivisorio Then
            MessageBox.Show("Nessun divisorio trovato nella griglia.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim totalePagine As Integer = dgvReport.Rows.Count

        Dim outPath As String = IO.Path.Combine(
            IO.Path.GetDirectoryName(PdfPath),
            IO.Path.GetFileNameWithoutExtension(PdfPath) & "_bookmarks.pdf")

        Try
            Using reader As New iText.Kernel.Pdf.PdfReader(PdfPath)
                Using writer As New iText.Kernel.Pdf.PdfWriter(outPath)
                    Using pdfDoc As New iText.Kernel.Pdf.PdfDocument(reader, writer,
                          New iText.Kernel.Pdf.StampingProperties().UseAppendMode())

                        Dim rootOutline = pdfDoc.GetOutlines(False)
                        Dim outlineMap As New Dictionary(Of String, PdfOutline)(StringComparer.OrdinalIgnoreCase)

                        ' Colore grigio per voci senza destinazione
                        Dim colorGrigio As New iText.Kernel.Colors.DeviceRgb(0.6F, 0.6F, 0.6F)

                        For Each voce In indice
                            Dim numero = voce.Numero
                            Dim desc = voce.Descrizione
                            Dim livello = GetLivello(numero)

                            Dim paginaDiv As Integer = -1
                            divisoriPagina.TryGetValue(numero, paginaDiv)

                            ' Salta le voci senza divisorio se il flag non è attivo
                            If paginaDiv <= 0 AndAlso Not flagMostraVociSenzaDivisorio Then
                                Continue For
                            End If

                            ' Determina la pagina di destinazione
                            Dim paginaDest As Integer = paginaDiv

                            If flagPuntaDopoDiv AndAlso paginaDiv > 0 Then
                                Dim paginaSuccessiva As Integer = paginaDiv + 1
                                If paginaSuccessiva <= totalePagine AndAlso
                                   Not pagineDivisorio.Contains(paginaSuccessiva) Then
                                    ' La pagina successiva esiste e non è un divisorio → punta lì
                                    paginaDest = paginaSuccessiva
                                Else
                                    ' La pagina successiva non esiste o è un divisorio → punta al divisorio
                                    paginaDest = paginaDiv
                                End If
                            End If

                            ' Determina il parent
                            Dim parentOutline As PdfOutline
                            If livello = 1 Then
                                parentOutline = rootOutline
                            Else
                                Dim numeropadre = GetPadre(numero)
                                If outlineMap.ContainsKey(numeropadre) Then
                                    parentOutline = outlineMap(numeropadre)
                                Else
                                    parentOutline = rootOutline
                                End If
                            End If

                            Dim child = parentOutline.AddOutline(desc)

                            If paginaDest > 0 AndAlso paginaDest <= pdfDoc.GetNumberOfPages() Then
                                ' Voce con destinazione
                                child.AddDestination(
                                    iText.Kernel.Pdf.Navigation.PdfExplicitDestination.
                                    CreateFit(pdfDoc.GetPage(paginaDest)))
                            Else
                                ' Voce senza divisorio → grigio, nessuna destinazione
                                child.SetColor(colorGrigio)
                            End If

                            outlineMap(numero) = child
                        Next

                    End Using
                End Using
            End Using

            MessageBox.Show($"File salvato:{Environment.NewLine}{outPath}",
                    "Completato", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Errore durante la creazione dei bookmark:" &
                    Environment.NewLine & ex.ToString(),
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    ' ══════════════════════════════════════════════════════════
    '  INDICE BOOKMARK
    ' ══════════════════════════════════════════════════════════

    Private Sub btnSfogliaIndice_Click(sender As Object, e As EventArgs) Handles btnSfogliaIndice.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Seleziona file indice"
            ofd.Filter = "File di testo (*.txt)|*.txt|Tutti i file (*.*)|*.*"
            If ofd.ShowDialog() = DialogResult.OK Then
                txtIndicePath.Text = ofd.FileName
                _indiceBookmark = CaricaIndice(ofd.FileName)
                My.Settings.IndicePath = ofd.FileName
                My.Settings.Save()
            End If
        End Using
    End Sub

    Private Function CaricaIndice(path As String) As List(Of (Numero As String, Descrizione As String))
        Dim result As New List(Of (Numero As String, Descrizione As String))
        If Not File.Exists(path) Then Return result

        For Each line As String In File.ReadAllLines(path)
            Dim trimmed = line.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For
            Dim parts = trimmed.Split(";"c)
            If parts.Length >= 2 Then
                result.Add((parts(0).Trim(), parts(1).Trim()))
            End If
        Next
        Return result
    End Function

    Private Function GetLivello(numero As String) As Integer
        Dim parts = numero.Split("."c)
        Return parts.Length - 1
    End Function

    Private Function GetPadre(numero As String) As String
        Dim idx = numero.LastIndexOf("."c)
        If idx <= 0 Then Return ""
        Return numero.Substring(0, idx)
    End Function

    Private Sub txtPdfPath_TextChanged(sender As Object, e As EventArgs) Handles txtPdfPath.TextChanged

    End Sub

    ' ══════════════════════════════════════════════════════════
    '  RESET
    ' ══════════════════════════════════════════════════════════

    Private Sub ResetTutto()
        ' Svuota griglia pagine
        dgvReport.Rows.Clear()
        ' Svuota griglia totali
        If dgvTotali IsNot Nothing Then dgvTotali.Rows.Clear()
        ' Reset testo e path
        txtPdfPath.Text = ""
        PdfPath = ""
        ' Chiude il PDF aperto
        If CurrentPdf IsNot Nothing Then
            CurrentPdf.Dispose()
            CurrentPdf = Nothing
        End If
        ' Reset progress bar
        AggiornaTesto(0, 0)
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ResetTutto()
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  ABILITA / DISABILITA CONTROLLI
    ' ══════════════════════════════════════════════════════════

    Private Sub AbilitaControlli(abilita As Boolean)
        btnReset.Enabled = abilita
        btnInviaStampante1.Enabled = abilita
        btnInviaStampante2.Enabled = abilita
        btnAggiungiBookmark.Enabled = abilita
        dgvReport.Enabled = abilita
        nudCopie.Enabled = abilita
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  SALVA / CARICA IMPOSTAZIONI JSON
    ' ══════════════════════════════════════════════════════════

    Private Class AppSettings
        Public Property NomeCartaA4 As String = ""
        Public Property NomeCartaA3 As String = ""
        Public Property NomeCartaDivisorio As String = ""
        Public Property PercorsoStampante1 As String = ""
        Public Property PercorsoStampante2 As String = ""
        Public Property MostraVociSenzaDivisorio As Boolean = False
        Public Property PuntaDopoDiv As Boolean = False
        Public Property Copie As Integer = 1
        Public Property IndicePath As String = ""
    End Class

    Private Sub btnSalvaImpostazioniFile_Click(sender As Object, e As EventArgs) Handles btnSalvaImpostazioniFile.Click
        Using sfd As New SaveFileDialog()
            sfd.Title = "Salva impostazioni"
            sfd.Filter = "File impostazioni (*.json)|*.json"
            sfd.FileName = "pdfrosetti_settings.json"
            If sfd.ShowDialog() <> DialogResult.OK Then Return

            Dim settings As New AppSettings With {
                .NomeCartaA4 = txtCartaA4.Text.Trim(),
                .NomeCartaA3 = txtCartaA3.Text.Trim(),
                .NomeCartaDivisorio = txtCartaDivisorio.Text.Trim(),
                .PercorsoStampante1 = txtStampante1.Text.Trim(),
                .PercorsoStampante2 = txtStampante2.Text.Trim(),
                .MostraVociSenzaDivisorio = chkMostraVociSenzaDivisorio.Checked,
                .PuntaDopoDiv = chkPuntaDopoDiv.Checked,
                .Copie = CInt(nudCopie.Value),
                .IndicePath = txtIndicePath.Text.Trim()
            }

            Dim json As String = SerializzaJson(settings)
            File.WriteAllText(sfd.FileName, json, System.Text.Encoding.UTF8)
            MessageBox.Show("Impostazioni salvate.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
    End Sub

    Private Sub btnCaricaImpostazioniFile_Click(sender As Object, e As EventArgs) Handles btnCaricaImpostazioniFile.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Carica impostazioni"
            ofd.Filter = "File impostazioni (*.json)|*.json"
            If ofd.ShowDialog() <> DialogResult.OK Then Return

            Try
                Dim json As String = File.ReadAllText(ofd.FileName, System.Text.Encoding.UTF8)
                Dim settings As AppSettings = DeserializzaJson(json)

                ' Applica ai controlli
                txtCartaA4.Text = settings.NomeCartaA4
                txtCartaA3.Text = settings.NomeCartaA3
                txtCartaDivisorio.Text = settings.NomeCartaDivisorio
                txtStampante1.Text = settings.PercorsoStampante1
                txtStampante2.Text = settings.PercorsoStampante2
                chkMostraVociSenzaDivisorio.Checked = settings.MostraVociSenzaDivisorio
                chkPuntaDopoDiv.Checked = settings.PuntaDopoDiv
                nudCopie.Value = Math.Max(nudCopie.Minimum, Math.Min(nudCopie.Maximum, settings.Copie))
                If Not String.IsNullOrWhiteSpace(settings.IndicePath) AndAlso File.Exists(settings.IndicePath) Then
                    txtIndicePath.Text = settings.IndicePath
                    _indiceBookmark = CaricaIndice(settings.IndicePath)
                End If

                ' Salva anche nei My.Settings
                My.Settings.NomeCartaA4 = settings.NomeCartaA4
                My.Settings.NomeCartaA3 = settings.NomeCartaA3
                My.Settings.NomeCartaDivisorio = settings.NomeCartaDivisorio
                My.Settings.PercorsoStampante1 = settings.PercorsoStampante1
                My.Settings.PercorsoStampante2 = settings.PercorsoStampante2
                My.Settings.MostraVociSenzaDivisorio = settings.MostraVociSenzaDivisorio
                My.Settings.PuntaDopoDiv = settings.PuntaDopoDiv
                My.Settings.IndicePath = settings.IndicePath
                My.Settings.Save()

                MessageBox.Show("Impostazioni caricate.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("Errore durante il caricamento:" & Environment.NewLine & ex.Message,
                                "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    ' Serializzazione JSON manuale (senza dipendenze esterne)
    Private Function SerializzaJson(s As AppSettings) As String
        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("{")
        sb.AppendLine($"  ""NomeCartaA4"": ""{EscapeJson(s.NomeCartaA4)}"",")
        sb.AppendLine($"  ""NomeCartaA3"": ""{EscapeJson(s.NomeCartaA3)}"",")
        sb.AppendLine($"  ""NomeCartaDivisorio"": ""{EscapeJson(s.NomeCartaDivisorio)}"",")
        sb.AppendLine($"  ""PercorsoStampante1"": ""{EscapeJson(s.PercorsoStampante1)}"",")
        sb.AppendLine($"  ""PercorsoStampante2"": ""{EscapeJson(s.PercorsoStampante2)}"",")
        sb.AppendLine($"  ""MostraVociSenzaDivisorio"": {s.MostraVociSenzaDivisorio.ToString().ToLower()},")
        sb.AppendLine($"  ""PuntaDopoDiv"": {s.PuntaDopoDiv.ToString().ToLower()},")
        sb.AppendLine($"  ""Copie"": {s.Copie},")
        sb.AppendLine($"  ""IndicePath"": ""{EscapeJson(s.IndicePath)}""")
        sb.AppendLine("}")
        Return sb.ToString()
    End Function

    Private Function EscapeJson(s As String) As String
        If String.IsNullOrEmpty(s) Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""")
    End Function

    Private Function DeserializzaJson(json As String) As AppSettings
        Dim s As New AppSettings()
        For Each line In json.Split(New Char() {Chr(10)}, StringSplitOptions.RemoveEmptyEntries)
            line = line.Trim().TrimEnd(",")
            Dim parts = line.Split(New Char() {":"c}, 2)
            If parts.Length < 2 Then Continue For
            Dim key = parts(0).Trim().Trim(""""c)
            Dim val = parts(1).Trim()
            Select Case key
                Case "NomeCartaA4" : s.NomeCartaA4 = val.Trim(""""c).Replace("\\", "\")
                Case "NomeCartaA3" : s.NomeCartaA3 = val.Trim(""""c).Replace("\\", "\")
                Case "NomeCartaDivisorio" : s.NomeCartaDivisorio = val.Trim(""""c).Replace("\\", "\")
                Case "PercorsoStampante1" : s.PercorsoStampante1 = val.Trim(""""c).Replace("\\", "\")
                Case "PercorsoStampante2" : s.PercorsoStampante2 = val.Trim(""""c).Replace("\\", "\")
                Case "MostraVociSenzaDivisorio" : s.MostraVociSenzaDivisorio = (val.ToLower() = "true")
                Case "PuntaDopoDiv" : s.PuntaDopoDiv = (val.ToLower() = "true")
                Case "Copie" : Integer.TryParse(val, s.Copie)
                Case "IndicePath" : s.IndicePath = val.Trim(""""c).Replace("\\", "\")
            End Select
        Next
        Return s
    End Function

    Private Sub chbSettaggi_CheckedChanged(sender As Object, e As EventArgs) Handles chbSettaggi.CheckedChanged
        If chbSettaggi.Checked Then
            GroupBox2.Visible = True
            GroupBox1.Visible = True
        Else
            GroupBox2.Visible = False
            GroupBox1.Visible = False
        End If
    End Sub

    Private Sub chkMostraVociSenzaDivisorio_CheckedChanged(sender As Object, e As EventArgs) Handles chkMostraVociSenzaDivisorio.CheckedChanged
        My.Settings.MostraVociSenzaDivisorio = chkMostraVociSenzaDivisorio.Checked
        My.Settings.Save()
    End Sub

    Private Sub chkPuntaDopoDiv_CheckedChanged(sender As Object, e As EventArgs) Handles chkPuntaDopoDiv.CheckedChanged
        My.Settings.PuntaDopoDiv = chkPuntaDopoDiv.Checked
        My.Settings.Save()
    End Sub

End Class

''' <summary>
''' ProgressBar personalizzata che disegna il testo progressivo al centro.
''' </summary>
Public Class ProgressBarConTesto
    Inherits ProgressBar

    Public Property Testo As String = ""

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        If m.Msg = &HF AndAlso Testo <> "" Then
            Using g As Graphics = Graphics.FromHwnd(Me.Handle)
                Dim font As New System.Drawing.Font("Segoe UI", 9, FontStyle.Bold)
                Dim brush As New SolidBrush(Color.Black)
                Dim sz As SizeF = g.MeasureString(Testo, font)
                Dim x As Single = (Me.Width - sz.Width) / 2
                Dim y As Single = (Me.Height - sz.Height) / 2
                g.DrawString(Testo, font, brush, x, y)
                font.Dispose()
                brush.Dispose()
            End Using
        End If
    End Sub

End Class
