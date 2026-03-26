Imports System.IO
Imports System.Drawing.Imaging
Imports PdfiumViewer
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Outlines
Imports ZXing
Imports ZXing.Common


Imports PdfiumDoc = PdfiumViewer.PdfDocument

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
        ' Ripristina file indice salvato
        If Not String.IsNullOrWhiteSpace(My.Settings.IndicePath) AndAlso File.Exists(My.Settings.IndicePath) Then
            txtIndicePath.Text = My.Settings.IndicePath
            _indiceBookmark = CaricaIndice(My.Settings.IndicePath)
        End If
        SetupGrid()
        AggiornaTesto(0, 0)
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
            If files.Length > 0 AndAlso Path.GetExtension(files(0)).ToLower = ".pdf" Then
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
                    g.DrawImage(bmp, New Rectangle(pad, pad, cropW, cropH),
                            New Rectangle(0, cropY, cropW, cropH), GraphicsUnit.Pixel)
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
            Dim data = bmp24.LockBits(New Rectangle(0, 0, bmp24.Width, bmp24.Height),
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
    '  GENERAZIONE XPIF JOB TICKET
    ' ══════════════════════════════════════════════════════════

    Private Function MmToXpif(mm As Double) As Integer
        Return CInt(mm * 100.0 / 25.4)
    End Function

    Private Function GetMediaSize(dimensione As String, isDivisorio As Boolean) As (X As Integer, Y As Integer)
        If isDivisorio Then Return (22500, 29700)
        Select Case dimensione
            Case "A0", "A1", "A2" : Return (29700, 42000)
            Case "A3" : Return (29700, 42000)
            Case "A4" : Return (21000, 29700)
            Case "A5" : Return (14800, 21000)
            Case Else
                Dim parts = dimensione.Replace(" mm", "").Split("x"c)
                If parts.Length = 2 Then
                    Dim w As Double, h As Double
                    If Double.TryParse(parts(0).Trim(), w) AndAlso Double.TryParse(parts(1).Trim(), h) Then
                        Return (CInt(w * 100), CInt(h * 100))
                    End If
                End If
                Return (21000, 29700)
        End Select
    End Function

    Private Function RichiudePiega(dimensione As String, isDivisorio As Boolean) As Boolean
        If isDivisorio Then Return False
        Return dimensione = "A3" OrElse dimensione = "A2" OrElse
               dimensione = "A1" OrElse dimensione = "A0"
    End Function

    Private Function RichiedeZoom(dimensione As String) As Boolean
        Return dimensione = "A2" OrElse dimensione = "A1" OrElse dimensione = "A0"
    End Function

    Private Sub btnGeneraXpif_Click(sender As Object, e As EventArgs) Handles btnGeneraXpif.Click

        If String.IsNullOrEmpty(PdfPath) Then
            MessageBox.Show("Carica prima un PDF.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        If dgvReport.Rows.Count = 0 Then
            MessageBox.Show("Nessuna pagina in griglia.", "Attenzione",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim outPath As String = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(PdfPath),
            System.IO.Path.GetFileNameWithoutExtension(PdfPath) & ".xpif")

        Try
            Dim sb As New System.Text.StringBuilder()

            sb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            sb.AppendLine("<!DOCTYPE xpif SYSTEM ""xpif-v2000.dtd"">")
            sb.AppendLine("<xpif version=""1.0"" cpss-version=""2.0"" xml:lang=""en"">")
            sb.AppendLine()

            Dim jobName As String = System.IO.Path.GetFileNameWithoutExtension(PdfPath)
            sb.AppendLine("  <xpif-operation-attributes>")
            sb.AppendLine($"    <job-name syntax=""name"" xml:space=""preserve"">{jobName}</job-name>")
            sb.AppendLine("  </xpif-operation-attributes>")
            sb.AppendLine()

            sb.AppendLine("  <job-template-attributes>")
            sb.AppendLine()
            sb.AppendLine("    <page-overrides syntax=""1setOf"">")
            sb.AppendLine()

            For Each row As DataGridViewRow In dgvReport.Rows

                Dim pageNum As Integer = row.Index + 1
                Dim dimensione As String = If(row.Cells("colDimensione").Value?.ToString(), "")
                Dim colore As String = If(row.Cells("colBNColore").Value?.ToString(), "BN")
                Dim divisorio As String = If(row.Cells("colDivisorio").Value?.ToString(), "")

                Dim isDivisorio As Boolean = (divisorio <> "" AndAlso divisorio <> "ERR")
                Dim mediaSize = GetMediaSize(dimensione, isDivisorio)
                Dim piega As Boolean = RichiudePiega(dimensione, isDivisorio)
                Dim zoom As Boolean = RichiedeZoom(dimensione) AndAlso Not isDivisorio
                Dim colorEffect As String = If(isDivisorio OrElse colore = "BN", "monochrome", "full-color")

                sb.AppendLine($"      <value syntax=""collection"">")
                sb.AppendLine($"        <input-documents syntax=""1setOf"">")
                sb.AppendLine($"          <value syntax=""rangeOfInteger"">")
                sb.AppendLine($"            <lower-bound syntax=""integer"">1</lower-bound>")
                sb.AppendLine($"            <upper-bound syntax=""integer"">1</upper-bound>")
                sb.AppendLine($"          </value>")
                sb.AppendLine($"        </input-documents>")
                sb.AppendLine($"        <pages syntax=""1setOf"">")
                sb.AppendLine($"          <value syntax=""rangeOfInteger"">")
                sb.AppendLine($"            <lower-bound syntax=""integer"">{pageNum}</lower-bound>")
                sb.AppendLine($"            <upper-bound syntax=""integer"">{pageNum}</upper-bound>")
                sb.AppendLine($"          </value>")
                sb.AppendLine($"        </pages>")
                sb.AppendLine($"        <media-col syntax=""collection"">")
                sb.AppendLine($"          <media-size syntax=""collection"">")
                sb.AppendLine($"            <x-dimension syntax=""integer"">{mediaSize.X}</x-dimension>")
                sb.AppendLine($"            <y-dimension syntax=""integer"">{mediaSize.Y}</y-dimension>")
                sb.AppendLine($"          </media-size>")
                sb.AppendLine($"        </media-col>")
                sb.AppendLine($"        <color-effects-type syntax=""keyword"">{colorEffect}</color-effects-type>")

                If zoom Then
                    sb.AppendLine($"        <x-auto-scale syntax=""keyword"">fit-to-page</x-auto-scale>")
                End If

                If piega Then
                    sb.AppendLine($"        <finishings-col syntax=""collection"">")
                    sb.AppendLine($"          <finishing-template syntax=""keyword"">fold-z</finishing-template>")
                    sb.AppendLine($"        </finishings-col>")
                End If

                If isDivisorio Then
                    Dim shiftUnits As Integer = MmToXpif(15.0)
                    sb.AppendLine($"        <x-side1-image-shift syntax=""integer"">{shiftUnits}</x-side1-image-shift>")
                End If

                sb.AppendLine($"      </value>")
                sb.AppendLine()

            Next

            sb.AppendLine("    </page-overrides>")
            sb.AppendLine()
            sb.AppendLine("  </job-template-attributes>")
            sb.AppendLine()
            sb.AppendLine("</xpif>")

            File.WriteAllText(outPath, sb.ToString(), System.Text.Encoding.UTF8)

            MessageBox.Show($"File XPIF salvato:{Environment.NewLine}{outPath}",
                            "Completato", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("Errore durante la generazione XPIF:" &
                            Environment.NewLine & ex.ToString(),
                            "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

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

        Dim divisoriPagina As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataGridViewRow In dgvReport.Rows
            Dim div As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            If div <> "" AndAlso div <> "ERR" Then
                If Not divisoriPagina.ContainsKey(div) Then
                    divisoriPagina.Add(div, row.Index + 1)
                End If
            End If
        Next

        If divisoriPagina.Count = 0 Then
            MessageBox.Show("Nessun divisorio trovato nella griglia.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

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

                        For Each voce In indice
                            Dim numero = voce.Numero
                            Dim desc = voce.Descrizione
                            Dim livello = GetLivello(numero)

                            Dim pagina As Integer = -1
                            divisoriPagina.TryGetValue(numero, pagina)

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

                            If pagina > 0 AndAlso pagina <= pdfDoc.GetNumberOfPages() Then
                                child.AddDestination(
                                    iText.Kernel.Pdf.Navigation.PdfExplicitDestination.
                                    CreateFit(pdfDoc.GetPage(pagina)))
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
                Dim font As New Font("Segoe UI", 9, FontStyle.Bold)
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
