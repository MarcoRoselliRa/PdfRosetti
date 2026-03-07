Imports System.IO
Imports System.Drawing.Imaging
Imports PdfiumViewer
Imports iText.Kernel.Pdf

Imports PdfiumDoc = PdfiumViewer.PdfDocument
Imports ITextPdfDoc = iText.Kernel.Pdf.PdfDocument

Public Class Form1

    Private PdfPath As String = ""
    Private CurrentPdf As PdfiumDoc = Nothing
    Private CurrentPreviewBmp As Bitmap = Nothing

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

    ''' <summary>
    ''' Inizializza il form: abilita drag&drop e configura la griglia.
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        SetupGrid()
        AggiornaTesto(0, 0)
    End Sub

    ''' <summary>
    ''' Configura colonne e proprietà della griglia report.
    ''' </summary>
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
            If col.Name <> "colPagina" Then
                col.DefaultCellStyle = centrata
            End If
        Next

        SetupGridTotali()
    End Sub

    ''' <summary>
    ''' Configura colonne e proprietà della griglia totali.
    ''' Due colonne: Descrizione e Quantità.
    ''' </summary>
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

    ''' <summary>
    ''' Aggiorna la progressbar e il testo centrale disegnato sopra di essa.
    ''' Se totale = 0 azzera tutto. Chiamare sempre dal thread UI.
    ''' </summary>
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

    ''' <summary>
    ''' Accetta il drag solo se il file trascinato è un PDF.
    ''' </summary>
    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length > 0 AndAlso Path.GetExtension(files(0)).ToLower = ".pdf" Then
                e.Effect = DragDropEffects.Copy
            End If
        End If
    End Sub

    ''' <summary>
    ''' Carica il PDF quando viene rilasciato sul form e avvia l'analisi completa.
    ''' </summary>
    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
        If files.Length = 0 Then Exit Sub
        CaricaPdf(files(0))
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  CARICAMENTO PDF
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Carica il PDF, popola la griglia con dimensioni e orientamento,
    ''' poi avvia immediatamente l'analisi automatica pagina per pagina.
    ''' </summary>
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

    ''' <summary>
    ''' Popola la griglia con numero pagina, formato ISO (se riconosciuto) e orientamento.
    ''' Le colonne BN/Colore e Divisorio vengono lasciate vuote in attesa dell'analisi.
    ''' </summary>
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

    ''' <summary>
    ''' Converte punti PDF in millimetri (1 pt = 25.4/72 mm).
    ''' </summary>
    Private Function PtToMm(pt As Double) As Double
        Return pt * 25.4 / 72.0
    End Function

    ''' <summary>
    ''' Riconosce il formato ISO della pagina (A0-A5) con tolleranza ±5mm.
    ''' Normalizza sempre in Portrait per il confronto (lato corto x lato lungo).
    ''' Se non corrisponde a nessun formato standard restituisce le dimensioni in mm.
    ''' </summary>
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

    ''' <summary>
    ''' Analizza tutte le pagine in background, una alla volta.
    ''' Per ogni pagina esegue un unico render e rileva: BN/Colore e Divisorio.
    ''' Aggiorna la griglia, evidenzia la riga corrente e aggiorna la progressbar.
    ''' Al termine torna alla prima riga e azzera la progressbar.
    ''' </summary>
    Private Async Sub AnalizzaTuttoAsync()

        Dim totale As Integer = CurrentPdf.PageCount
        AggiornaTesto(0, totale)

        Await Task.Run(Sub()

            ' Apre istanza separata del PDF nel thread background
            ' (PdfiumViewer non è thread-safe, non si può usare CurrentPdf)
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

                        ' Render unico usato sia per colore che per divisori
                        Using bmp As Bitmap = pdf.Render(rowIndex, w, h, 120, 120, PdfRenderFlags.Annotations)

                            ' Analisi colore
                            risultatoColore = If(IsColorPixel(bmp), "Colore", "BN")

                            ' Analisi divisorio
                            Dim blackBox = FindBlackDividerBox(bmp)
                            If blackBox.HasValue Then
                                Dim numero As String = ReadDividerNumber(bmp, blackBox.Value)
                                risultatoDivisorio = If(numero <> "", numero, "DIV")
                            End If

                        End Using

                    Catch ex As Exception
                        risultatoColore = "ERR"
                        Debug.WriteLine($"Errore pagina {rowIndex}: {ex.Message}")
                    End Try

                    ' Aggiorna solo i dati e la progressbar — niente selezione né scroll
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

        ' Torna alla prima riga e azzera progressbar
        If dgvReport.Rows.Count > 0 Then
            dgvReport.ClearSelection()
            dgvReport.Rows(0).Selected = True
            dgvReport.FirstDisplayedScrollingRowIndex = 0
        End If
        AggiornaTesto(0, 0)
        AggiornaTotali()

    End Sub

    ' ══════════════════════════════════════════════════════════
    '  TOTALI
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Legge i dati dalla griglia report e popola dgvTotali con i conteggi
    ''' raggruppati per formato + BN/Colore, più divisori e fuori formato.
    ''' Viene chiamata al termine dell'analisi automatica.
    ''' </summary>
    Private Sub AggiornaTotali()
        dgvTotali.Rows.Clear()

        ' Dizionario per conteggi: chiave = "A4 Colore", "A4 BN", ecc.
        Dim conteggi As New Dictionary(Of String, Integer)
        Dim divisori As Integer = 0
        Dim fuoriFormato As Integer = 0

        Dim formatiNoti As New HashSet(Of String)({"A0", "A1", "A2", "A3", "A4", "A5"})

        For Each row As DataGridViewRow In dgvReport.Rows
            Dim dimensione As String = If(row.Cells("colDimensione").Value?.ToString(), "")
            Dim colore As String = If(row.Cells("colBNColore").Value?.ToString(), "")
            Dim divisorio As String = If(row.Cells("colDivisorio").Value?.ToString(), "")

            ' Conta divisori
            If divisorio <> "" Then
                divisori += 1
                Continue For  ' la pagina divisorio non conta nei formati
            End If

            ' Fuori formato
            If Not formatiNoti.Contains(dimensione) Then
                fuoriFormato += 1
                Continue For
            End If

            ' Formato noto: raggruppa per formato + BN/Colore
            If colore = "BN" OrElse colore = "Colore" Then
                Dim chiave As String = $"{dimensione} {colore}"
                If conteggi.ContainsKey(chiave) Then
                    conteggi(chiave) += 1
                Else
                    conteggi(chiave) = 1
                End If
            End If
        Next

        ' Ordine di visualizzazione: A0→A5, prima Colore poi BN
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

        ' Divisori
        If divisori > 0 Then
            dgvTotali.Rows.Add("Divisori", divisori)
        End If

        ' Fuori formato
        If fuoriFormato > 0 Then
            dgvTotali.Rows.Add("Fuori formato", fuoriFormato)
        End If
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  ANALISI COLORE
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Verifica se la bitmap contiene pixel colorati (non BN).
    ''' Campiona ogni 4 pixel per velocità.
    ''' Soglia: somma delle differenze canali RGB > 10.
    ''' </summary>
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

    ''' <summary>
    ''' Aggiorna la preview quando l'utente clicca su una riga della griglia.
    ''' </summary>
    Private Sub dgvReport_SelectionChanged(sender As Object, e As EventArgs) Handles dgvReport.SelectionChanged
        If CurrentPdf Is Nothing Then Exit Sub
        If dgvReport.SelectedRows.Count = 0 Then Exit Sub
        Dim pageIndex As Integer = dgvReport.SelectedRows(0).Index
        AggiornaAnteprimaPagina(pageIndex)
    End Sub

    ''' <summary>
    ''' Renderizza la pagina selezionata e la mostra nella picPreview.
    ''' </summary>
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

    ''' <summary>
    ''' Libera le risorse grafiche e il documento PDF alla chiusura del form.
    ''' </summary>
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If CurrentPreviewBmp IsNot Nothing Then CurrentPreviewBmp.Dispose()
        If CurrentPdf IsNot Nothing Then CurrentPdf.Dispose()
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  DIVISORI - VERIFICA PAGINA BIANCA
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Verifica che la zona sinistra della pagina (primo 78% della larghezza)
    ''' sia prevalentemente bianca. Se c'è troppo contenuto non è un divisorio.
    ''' Soglia: almeno 98% di pixel con luminosità > 200.
    ''' </summary>
    Private Function PaginaPrevalentementeBianca(bmp As Bitmap) As Boolean
        Dim checkW As Integer = CInt(bmp.Width * 0.78)
        Dim totalPixels As Integer = 0
        Dim whitePixels As Integer = 0

        For y As Integer = 0 To bmp.Height - 1 Step 4
            For x As Integer = 0 To checkW Step 4
                Dim c As Color = bmp.GetPixel(x, y)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3
                If lum > 200 Then whitePixels += 1
                totalPixels += 1
            Next
        Next

        If totalPixels = 0 Then Return False
        Return (whitePixels / totalPixels) >= 0.98
    End Function

    ' ══════════════════════════════════════════════════════════
    '  DIVISORI - RICERCA BOX NERO
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Cerca il box nero del divisorio nella fascia destra della pagina (80-98% larghezza).
    ''' Prima verifica che la pagina sia prevalentemente bianca (caratteristica dei divisori).
    ''' Scansiona finestre orizzontali cercando la zona con più pixel scuri,
    ''' poi stringe il bounding box ed estende verso alto e basso.
    ''' Restituisce Nothing se non trova nessun divisorio.
    ''' </summary>
    Private Function FindBlackDividerBox(bmp As Bitmap) As Rectangle?

        If Not PaginaPrevalentementeBianca(bmp) Then Return Nothing

        Dim roiX As Integer = CInt(bmp.Width * 0.8)
        Dim roiW As Integer = CInt(bmp.Width * 0.18)
        If roiX < 0 Then roiX = 0
        If roiX >= bmp.Width Then Return Nothing
        If roiX + roiW > bmp.Width Then roiW = bmp.Width - roiX
        If roiW <= 0 Then Return Nothing

        Dim winH As Integer = Math.Max(60, CInt(bmp.Height * 0.08))
        Dim bestY As Integer = -1
        Dim bestRatio As Double = 0

        For y As Integer = 0 To Math.Max(0, bmp.Height - winH) Step 6
            Dim darkCount As Integer = 0
            Dim total As Integer = 0
            For yy As Integer = y To Math.Min(y + winH - 1, bmp.Height - 1) Step 2
                For xx As Integer = roiX To Math.Min(roiX + roiW - 1, bmp.Width - 1) Step 2
                    Dim c As Color = bmp.GetPixel(xx, yy)
                    Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3
                    If lum < 110 Then darkCount += 1
                    total += 1
                Next
            Next
            If total > 0 Then
                Dim ratio As Double = darkCount / total
                If ratio > bestRatio Then
                    bestRatio = ratio
                    bestY = y
                End If
            End If
        Next

        If bestY < 0 OrElse bestRatio <= 0.18 Then Return Nothing

        Dim minX As Integer = bmp.Width
        Dim maxX As Integer = 0
        Dim minY As Integer = bmp.Height
        Dim maxY As Integer = 0
        Dim found As Boolean = False

        For y As Integer = bestY To Math.Min(bestY + winH - 1, bmp.Height - 1)
            For x As Integer = roiX To Math.Min(roiX + roiW - 1, bmp.Width - 1)
                Dim c As Color = bmp.GetPixel(x, y)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3
                If lum < 90 Then
                    found = True
                    If x < minX Then minX = x
                    If x > maxX Then maxX = x
                    If y < minY Then minY = y
                    If y > maxY Then maxY = y
                End If
            Next
        Next

        If Not found Then Return Nothing
        If maxX <= minX OrElse maxY <= minY Then Return Nothing

        ' Estendi verso l'alto
        Dim yTop As Integer = minY
        While yTop > 0
            Dim darkFound As Boolean = False
            For x As Integer = minX To maxX
                Dim c As Color = bmp.GetPixel(x, yTop)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3
                If lum < 100 Then darkFound = True : Exit For
            Next
            If Not darkFound Then Exit While
            yTop -= 1
        End While

        ' Estendi verso il basso
        Dim yBottom As Integer = maxY
        While yBottom < bmp.Height - 1
            Dim darkFound As Boolean = False
            For x As Integer = minX To maxX
                Dim c As Color = bmp.GetPixel(x, yBottom)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3
                If lum < 100 Then darkFound = True : Exit For
            Next
            If Not darkFound Then Exit While
            yBottom += 1
        End While

        Dim pad As Integer = 6
        Dim rx As Integer = Math.Max(0, minX - pad)
        Dim ry As Integer = Math.Max(0, yTop - pad)
        Dim rw As Integer = Math.Min(bmp.Width - rx, (maxX - minX + 1) + pad * 2)
        Dim rh As Integer = Math.Min(bmp.Height - ry, (yBottom - yTop + 1) + pad * 2)

        Return New Rectangle(rx, ry, rw, rh)
    End Function

    ' ══════════════════════════════════════════════════════════
    '  DIVISORI - LETTURA NUMERO
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Ritaglia una porzione rettangolare dalla bitmap sorgente.
    ''' </summary>
    Private Function CropBitmap(src As Bitmap, r As Rectangle) As Bitmap
        Dim bmp As New Bitmap(r.Width, r.Height)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.DrawImage(src, New Rectangle(0, 0, r.Width, r.Height), r, GraphicsUnit.Pixel)
        End Using
        Return bmp
    End Function

    ''' <summary>
    ''' Legge il numero del divisorio dal box trovato.
    ''' Applica Rotate90 (rotazione corretta per questi PDF), prepara l'immagine
    ''' e la passa a Tesseract. Restituisce il numero (1-99) o stringa vuota.
    ''' </summary>
    Private Function ReadDividerNumber(pageBmp As Bitmap, rect As Rectangle) As String
        Using crop As Bitmap = CropBitmap(pageBmp, rect)
            Dim rotated As Bitmap = CType(crop.Clone(), Bitmap)
            rotated.RotateFlip(RotateFlipType.Rotate90FlipNone)

            Using prep As Bitmap = PrepareDividerForOcr(rotated)
                rotated.Dispose()

                Dim res As String = ""
                Dim conf As Single = 0.0F
                RunTesseractMultiMode(prep, res, conf)

                If res <> "" Then
                    Dim n As Integer
                    If Integer.TryParse(res, n) Then
                        If n >= 1 AndAlso n <= 99 Then Return n.ToString()
                    End If
                End If
            End Using
        End Using

        Return ""
    End Function

    ''' <summary>
    ''' Prepara il crop del divisorio per OCR:
    ''' 1) Scala 4x con interpolazione bicubica
    ''' 2) Binarizzazione con threshold Otsu adattivo
    ''' 3) Inversione colori (sfondo nero → bianco, testo bianco → nero)
    ''' 4) Taglio 15% per lato per eliminare la cornice nera del divisorio
    ''' 5) Aggiunta bordo bianco per migliorare il riconoscimento Tesseract
    ''' </summary>
    Private Function PrepareDividerForOcr(src As Bitmap) As Bitmap

        ' 1) Scala 4x
        Dim enlarged As New Bitmap(src.Width * 4, src.Height * 4, Imaging.PixelFormat.Format24bppRgb)
        Using g As Graphics = Graphics.FromImage(enlarged)
            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
            g.DrawImage(src, New Rectangle(0, 0, enlarged.Width, enlarged.Height))
        End Using

        ' 2) Calcola threshold Otsu
        Dim histogram(255) As Integer
        For y As Integer = 0 To enlarged.Height - 1
            For x As Integer = 0 To enlarged.Width - 1
                Dim c As Color = enlarged.GetPixel(x, y)
                Dim lum As Integer = CInt(c.R * 0.299 + c.G * 0.587 + c.B * 0.114)
                histogram(lum) += 1
            Next
        Next
        Dim thresh As Integer = OtsuThreshold(histogram, enlarged.Width * enlarged.Height)

        ' 3) Binarizza e inverti
        Dim outBmp As New Bitmap(enlarged.Width, enlarged.Height, Imaging.PixelFormat.Format24bppRgb)
        For y As Integer = 0 To enlarged.Height - 1
            For x As Integer = 0 To enlarged.Width - 1
                Dim c As Color = enlarged.GetPixel(x, y)
                Dim lum As Integer = CInt(c.R * 0.299 + c.G * 0.587 + c.B * 0.114)
                Dim v As Integer = If(lum < thresh, 255, 0)
                outBmp.SetPixel(x, y, Color.FromArgb(v, v, v))
            Next
        Next
        enlarged.Dispose()

        ' 4) Taglia il 15% per lato per eliminare la cornice nera
        Dim cx As Integer = CInt(outBmp.Width * 0.15)
        Dim cy As Integer = CInt(outBmp.Height * 0.15)
        Dim innerW As Integer = outBmp.Width - cx * 2
        Dim innerH As Integer = outBmp.Height - cy * 2

        Dim trimmed As New Bitmap(innerW, innerH, Imaging.PixelFormat.Format24bppRgb)
        Using g As Graphics = Graphics.FromImage(trimmed)
            g.Clear(Color.White)
            g.DrawImage(outBmp, New Rectangle(0, 0, innerW, innerH),
                    New Rectangle(cx, cy, innerW, innerH), GraphicsUnit.Pixel)
        End Using
        outBmp.Dispose()
        outBmp = trimmed

        ' 5) Bordo bianco
        Dim pad As Integer = 30
        Dim padded As New Bitmap(outBmp.Width + pad * 2, outBmp.Height + pad * 2,
                             Imaging.PixelFormat.Format24bppRgb)
        Using g As Graphics = Graphics.FromImage(padded)
            g.Clear(Color.White)
            g.DrawImage(outBmp, pad, pad)
        End Using
        outBmp.Dispose()

        Return padded
    End Function

    ''' <summary>
    ''' Calcola il threshold ottimale di binarizzazione con il metodo Otsu.
    ''' Massimizza la varianza inter-classe tra sfondo e primo piano.
    ''' </summary>
    Private Function OtsuThreshold(histogram() As Integer, totalPixels As Integer) As Integer
        Dim sum As Double = 0
        For i As Integer = 0 To 255
            sum += i * histogram(i)
        Next
        Dim sumB As Double = 0
        Dim wB As Integer = 0
        Dim maxVariance As Double = 0
        Dim threshold As Integer = 128
        For t As Integer = 0 To 255
            wB += histogram(t)
            If wB = 0 Then Continue For
            Dim wF As Integer = totalPixels - wB
            If wF = 0 Then Exit For
            sumB += t * histogram(t)
            Dim mB As Double = sumB / wB
            Dim mF As Double = (sum - sumB) / wF
            Dim variance As Double = CDbl(wB) * CDbl(wF) * (mB - mF) ^ 2
            If variance > maxVariance Then
                maxVariance = variance
                threshold = t
            End If
        Next
        Return threshold
    End Function

    ''' <summary>
    ''' Esegue Tesseract OCR con 3 modalità (SingleChar, SingleWord, SingleBlock)
    ''' e restituisce il risultato con confidence più alta.
    ''' Whitelist limitata alle cifre 0-9.
    ''' </summary>
    Private Sub RunTesseractMultiMode(bmp As Bitmap, ByRef result As String, ByRef confidence As Single)
        result = ""
        confidence = 0.0F

        Dim tempFile As String = Path.Combine(Path.GetTempPath(),
                                              "div_" & Guid.NewGuid().ToString("N") & ".png")
        Try
            bmp.Save(tempFile, Imaging.ImageFormat.Png)

            Using engine As New Tesseract.TesseractEngine("./tessdata", "eng",
                                                           Tesseract.EngineMode.Default)
                engine.SetVariable("tessedit_char_whitelist", "0123456789")
                engine.SetVariable("classify_bln_numeric_mode", "1")

                Using pix = Tesseract.Pix.LoadFromFile(tempFile)

                    Dim modes() As Tesseract.PageSegMode = {
                        Tesseract.PageSegMode.SingleChar,
                        Tesseract.PageSegMode.SingleWord,
                        Tesseract.PageSegMode.SingleBlock
                    }

                    For Each mode As Tesseract.PageSegMode In modes
                        Using page = engine.Process(pix, mode)
                            Dim conf As Single = page.GetMeanConfidence()
                            Dim txt As String = page.GetText()
                            If txt IsNot Nothing Then
                                Dim digits As String =
                                    New String(txt.Trim().Where(Function(ch) Char.IsDigit(ch)).ToArray())
                                If digits.Length > 0 AndAlso conf > confidence Then
                                    Dim candidate As String = digits.Substring(0, Math.Min(2, digits.Length))
                                    Dim num As Integer
                                    If Integer.TryParse(candidate, num) AndAlso num >= 1 AndAlso num <= 99 Then
                                        confidence = conf
                                        result = num.ToString()
                                    End If
                                End If
                            End If
                        End Using
                    Next

                End Using
            End Using

        Catch ex As Exception
            Debug.WriteLine("Tesseract error: " & ex.Message)
        Finally
            Try
                If File.Exists(tempFile) Then File.Delete(tempFile)
            Catch
            End Try
        End Try
    End Sub

    ' ══════════════════════════════════════════════════════════
    '  BOOKMARK PDF
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Crea una copia del PDF con un bookmark per ogni divisorio trovato.
    ''' Il file viene salvato nella stessa cartella dell'originale
    ''' con suffisso "_bookmarks.pdf".
    ''' </summary>
    Private Sub btnAggiungiBookmark_Click(sender As Object, e As EventArgs) Handles btnAggiungiBookmark.Click

        If String.IsNullOrEmpty(PdfPath) Then
            MessageBox.Show("Carica prima un PDF.", "Attenzione",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Raccoglie i divisori dalla griglia: pagina (1-based) e numero
        Dim divisori As New List(Of (Pagina As Integer, Numero As String))

        For Each row As DataGridViewRow In dgvReport.Rows
            Dim div As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            If div <> "" AndAlso div <> "DIV" AndAlso div <> "ERR" Then
                divisori.Add((row.Index + 1, div))
            End If
        Next

        If divisori.Count = 0 Then
            MessageBox.Show("Nessun divisorio trovato nella griglia.", "Attenzione",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Percorso file output
        Dim outPath As String = System.IO.Path.Combine(
        System.IO.Path.GetDirectoryName(PdfPath),
        System.IO.Path.GetFileNameWithoutExtension(PdfPath) & "_bookmarks.pdf")

        Try
            Using reader As New iText.Kernel.Pdf.PdfReader(PdfPath)
                Using writer As New iText.Kernel.Pdf.PdfWriter(outPath)
                    Using pdfDoc As New iText.Kernel.Pdf.PdfDocument(reader, writer,
                          New iText.Kernel.Pdf.StampingProperties().UseAppendMode())

                        Dim outline = pdfDoc.GetOutlines(False)

                        ' Ordina per numero divisorio
                        divisori = divisori.OrderBy(Function(d) CInt(d.Numero)).ToList()

                        For Each div In divisori
                            Dim titolo As String = $"Divisorio {div.Numero}"
                            Dim child = outline.AddOutline(titolo)
                            child.AddDestination(
                            iText.Kernel.Pdf.Navigation.PdfExplicitDestination.
                            CreateFit(pdfDoc.GetPage(div.Pagina)))
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

End Class

''' <summary>
''' ProgressBar personalizzata che disegna il testo progressivo al centro.
''' Sostituisce la Label trasparente che non funziona in WinForms.
''' </summary>
Public Class ProgressBarConTesto
    Inherits ProgressBar

    Public Property Testo As String = ""

    Protected Overrides Sub WndProc(ByRef m As Message)
        MyBase.WndProc(m)

        ' WM_PAINT = &HF — ridisegniamo il testo sopra la barra
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
