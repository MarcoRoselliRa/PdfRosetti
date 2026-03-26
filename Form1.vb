Imports System.IO
Imports System.Drawing.Imaging
Imports PdfiumViewer
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Outlines

Imports PdfiumDoc = PdfiumViewer.PdfDocument
Imports ITextPdfDoc = iText.Kernel.Pdf.PdfDocument

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
        Dim indiceLocale = _indiceBookmark

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
                                               Dim numero As String = ReadDividerNumber(bmp, blackBox.Value, indiceLocale)
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
        Dim checkW As Integer = CInt(bmp.Width * 0.72)
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

        ' Il tab è negli ultimi 5mm su 210mm = ultimo ~2.4% della larghezza
        ' Usiamo l'ultimo 8% per sicurezza
        Dim roiX As Integer = CInt(bmp.Width * 0.92)
        Dim roiW As Integer = bmp.Width - roiX
        If roiX < 0 Then roiX = 0
        If roiX >= bmp.Width Then Return Nothing
        If roiW <= 0 Then Return Nothing

        ' Scansiona tutta l'altezza cercando la finestra con più pixel scuri
        ' Il tab è 30mm su 297mm = ~10% dell'altezza
        Dim winH As Integer = Math.Max(40, CInt(bmp.Height * 0.1))
        Dim bestY As Integer = -1
        Dim bestRatio As Double = 0

        For y As Integer = 0 To Math.Max(0, bmp.Height - winH) Step 4
            Dim darkCount As Integer = 0
            Dim total As Integer = 0
            For yy As Integer = y To Math.Min(y + winH - 1, bmp.Height - 1) Step 2
                For xx As Integer = roiX To bmp.Width - 1 Step 1
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

        ' Soglia bassa perché il tab è piccolo
        If bestY < 0 OrElse bestRatio <= 0.05 Then Return Nothing

        ' Bounding box preciso
        Dim minX As Integer = bmp.Width
        Dim maxX As Integer = 0
        Dim minY As Integer = bmp.Height
        Dim maxY As Integer = 0
        Dim found As Boolean = False

        For y As Integer = bestY To Math.Min(bestY + winH - 1, bmp.Height - 1)
            For x As Integer = roiX To bmp.Width - 1
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

        Dim pad As Integer = 4
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
    Private Function ReadDividerNumber(pageBmp As Bitmap, rect As Rectangle,
                                   Optional indice As List(Of (Numero As String, Descrizione As String)) = Nothing) As String
        Using crop As Bitmap = CropBitmap(pageBmp, rect)
            Dim rotated As Bitmap = CType(crop.Clone(), Bitmap)
            rotated.RotateFlip(RotateFlipType.Rotate90FlipNone)

            Using prep As Bitmap = PrepareDividerForOcr(rotated)
                rotated.Dispose()

                Dim res As String = ""
                Dim conf As Single = 0.0F
                RunTesseractMultiMode(prep, res, conf)

                If res <> "" AndAlso res.Any(Function(c) Char.IsDigit(c)) Then
                    If indice IsNot Nothing AndAlso indice.Count > 0 Then
                        Return TrovaCorrispondenzaMigliore(res, indice)
                    End If
                    Return res
                End If
            End Using
        End Using

        Return ""
    End Function

    Private Function TrovaCorrispondenzaMigliore(ocrResult As String,
                                              indice As List(Of (Numero As String, Descrizione As String))) As String
        ' Prima prova: corrispondenza esatta
        For Each voce In indice
            If voce.Numero = ocrResult Then Return voce.Numero
        Next

        ' Estrai cifre dall'OCR
        Dim ocrDigits = New String(ocrResult.Where(Function(c) Char.IsDigit(c)).ToArray())
        Dim ocrDigitsAlt = ocrDigits.Replace("4"c, "1"c)  ' correzione 4→1

        Debug.WriteLine($"OCR letto: '{ocrResult}'  cifre: '{ocrDigits}'")

        Dim bestMatch As String = ""
        Dim bestScore As Integer = 999

        For Each voce In indice
            Dim indiceDigits = New String(voce.Numero.Where(Function(c) Char.IsDigit(c)).ToArray())

            Debug.WriteLine($"  confronto con: '{voce.Numero}'  cifre: '{indiceDigits}'")

            If indiceDigits = ocrDigits Then
                ' Corrispondenza esatta sulle cifre
                Dim lenDiff = Math.Abs(voce.Numero.Length - ocrResult.Length)
                Dim score = lenDiff
                If score < bestScore Then
                    bestScore = score
                    bestMatch = voce.Numero
                End If
            Else
                Dim dist = LevenshteinDistance(ocrDigits, indiceDigits)
                ' Usa la versione alt (4→1) solo se la distanza normale è > 1
                ' così non disturba i casi dove il 4 è corretto
                Dim distAlt = If(dist > 1, LevenshteinDistance(ocrDigitsAlt, indiceDigits), dist)
                Dim bestDist = Math.Min(dist, distAlt)
                Debug.WriteLine($"    Levenshtein '{ocrDigits}' vs '{indiceDigits}' = {dist}  alt={distAlt}")

                ' Soglia 2 per numeri corti, 3 per numeri lunghi (4+ cifre)
                Dim soglia = If(indiceDigits.Length >= 4, 3, 2)

                If bestDist <= soglia Then
                    Dim lenDiff = Math.Abs(ocrDigits.Length - indiceDigits.Length)
                    Dim score = bestDist * 10 + lenDiff
                    If score < bestScore Then
                        bestScore = score
                        bestMatch = voce.Numero
                    End If
                End If
            End If
        Next

        Return If(bestMatch <> "", bestMatch, ocrResult)
    End Function

    Private Function LevenshteinDistance(s As String, t As String) As Integer
        Dim n = s.Length
        Dim m = t.Length
        Dim d(n, m) As Integer

        If n = 0 Then Return m
        If m = 0 Then Return n

        For i = 0 To n : d(i, 0) = i : Next
        For j = 0 To m : d(0, j) = j : Next

        For j = 1 To m
            For i = 1 To n
                Dim cost = If(s(i - 1) = t(j - 1), 0, 1)
                d(i, j) = Math.Min(Math.Min(
                d(i - 1, j) + 1,
                d(i, j - 1) + 1),
                d(i - 1, j - 1) + cost)
            Next
        Next

        Return d(n, m)
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
                ' Aggiunto il punto alla whitelist
                engine.SetVariable("tessedit_char_whitelist", "0123456789.")
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
                                ' Estrai cifre e punti, rimuovi spazi/newline
                                Dim cleaned As String =
                                New String(txt.Trim().Where(
                                    Function(ch) Char.IsDigit(ch) OrElse ch = "."c
                                ).ToArray())

                                ' Valida formato: deve essere N o N.N o N.NN
                                If cleaned.Length > 0 AndAlso conf > confidence Then
                                    ' Verifica che sia un numero valido tipo "1.0", "5.11", "10"
                                    Dim parts = cleaned.Split("."c)
                                    Dim valido As Boolean = False
                                    If parts.Length = 1 AndAlso parts(0).Length >= 1 Then
                                        valido = True  ' es. "10"
                                    ElseIf parts.Length = 2 AndAlso parts(0).Length >= 1 AndAlso parts(1).Length >= 1 Then
                                        valido = True  ' es. "1.0", "5.11"
                                    End If

                                    If valido Then
                                        confidence = conf
                                        result = cleaned
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
    '  GENERAZIONE XPIF JOB TICKET
    ' ══════════════════════════════════════════════════════════

    ''' <summary>
    ''' Converte millimetri in unità XPIF (1/100 di pollice).
    ''' </summary>
    Private Function MmToXpif(mm As Double) As Integer
        Return CInt(mm * 100.0 / 25.4)
    End Function

    ''' <summary>
    ''' Restituisce le dimensioni media-size XPIF in 1/100 di mm per un formato.
    ''' A2/A1/A0 vengono mappati su A3 (stampa con zoom).
    ''' I divisori usano 225x297mm.
    ''' </summary>
    Private Function GetMediaSize(dimensione As String, isDivisorio As Boolean) As (X As Integer, Y As Integer)
        If isDivisorio Then Return (22500, 29700)
        Select Case dimensione
            Case "A0", "A1", "A2" : Return (29700, 42000)
            Case "A3"             : Return (29700, 42000)
            Case "A4"             : Return (21000, 29700)
            Case "A5"             : Return (14800, 21000)
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

    ''' <summary>
    ''' Restituisce true se il formato richiede piega a Z (A3 e formati grandi).
    ''' </summary>
    Private Function RichiudePiega(dimensione As String, isDivisorio As Boolean) As Boolean
        If isDivisorio Then Return False
        Return dimensione = "A3" OrElse dimensione = "A2" OrElse
               dimensione = "A1" OrElse dimensione = "A0"
    End Function

    ''' <summary>
    ''' Restituisce true se il formato richiede riduzione su A3 (A2, A1, A0).
    ''' </summary>
    Private Function RichiedeZoom(dimensione As String) As Boolean
        Return dimensione = "A2" OrElse dimensione = "A1" OrElse dimensione = "A0"
    End Function

    ''' <summary>
    ''' Genera il file XPIF job ticket leggendo tutti i dati dalla griglia.
    ''' Un unico file con page-overrides per ogni pagina.
    ''' Regole:
    '''   A4          → A4, BN/Colore da griglia, no piega
    '''   A3          → A3, BN/Colore da griglia, piega Z
    '''   A2/A1/A0    → A3 con zoom, BN/Colore da griglia, piega Z
    '''   Divisorio   → 225x297mm, monochrome, shift +15mm destra
    '''   Fuori formato → dimensioni reali, BN/Colore da griglia, no piega
    ''' </summary>
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

            ' Intestazione XML
            sb.AppendLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            sb.AppendLine("<!DOCTYPE xpif SYSTEM ""xpif-v2000.dtd"">")
            sb.AppendLine("<xpif version=""1.0"" cpss-version=""2.0"" xml:lang=""en"">")
            sb.AppendLine()

            ' Attributi operazione
            Dim jobName As String = System.IO.Path.GetFileNameWithoutExtension(PdfPath)
            sb.AppendLine("  <xpif-operation-attributes>")
            sb.AppendLine($"    <job-name syntax=""name"" xml:space=""preserve"">{jobName}</job-name>")
            sb.AppendLine("  </xpif-operation-attributes>")
            sb.AppendLine()

            ' Job template
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

        If String.IsNullOrEmpty(txtIndicePath.Text) OrElse Not File.Exists(txtIndicePath.Text) Then
            MessageBox.Show("Seleziona prima il file indice.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Carica indice dal file
        Dim indice = CaricaIndice(txtIndicePath.Text)
        If indice.Count = 0 Then
            MessageBox.Show("Il file indice è vuoto o non valido.", "Attenzione",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Raccoglie i divisori dalla griglia: numero → pagina (1-based)
        Dim divisoriPagina As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataGridViewRow In dgvReport.Rows
            Dim div As String = If(row.Cells("colDivisorio").Value?.ToString(), "")
            If div <> "" AndAlso div <> "DIV" AndAlso div <> "ERR" Then
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

        ' Percorso file output
        Dim outPath As String = IO.Path.Combine(
        IO.Path.GetDirectoryName(PdfPath),
        IO.Path.GetFileNameWithoutExtension(PdfPath) & "_bookmarks.pdf")

        Try
            Using reader As New iText.Kernel.Pdf.PdfReader(PdfPath)
                Using writer As New iText.Kernel.Pdf.PdfWriter(outPath)
                    Using pdfDoc As New iText.Kernel.Pdf.PdfDocument(reader, writer,
                      New iText.Kernel.Pdf.StampingProperties().UseAppendMode())

                        Dim rootOutline = pdfDoc.GetOutlines(False)

                        ' Dizionario per tenere traccia dei bookmark padre già creati
                        ' chiave = numero (es. "2.0"), valore = oggetto outline
                        Dim outlineMap As New Dictionary(Of String, PdfOutline)(
                        StringComparer.OrdinalIgnoreCase)

                        For Each voce In indice
                            Dim numero = voce.Numero
                            Dim desc = voce.Descrizione
                            Dim livello = GetLivello(numero)

                            ' Cerca la pagina corrispondente nel PDF
                            Dim pagina As Integer = -1
                            divisoriPagina.TryGetValue(numero, pagina)

                            ' Determina il parent outline
                            Dim parentOutline As PdfOutline
                            If livello = 1 Then
                                ' Padre → attacca alla root
                                parentOutline = rootOutline
                            Else
                                ' Figlio → cerca il padre nel dizionario
                                Dim numeropadre = GetPadre(numero)
                                If outlineMap.ContainsKey(numeropadre) Then
                                    parentOutline = outlineMap(numeropadre)
                                Else
                                    ' Padre non trovato, attacca alla root
                                    parentOutline = rootOutline
                                End If
                            End If

                            ' Crea il bookmark
                            Dim child = parentOutline.AddOutline(desc)

                            ' Collega alla pagina se trovata, altrimenti bookmark senza destinazione
                            If pagina > 0 AndAlso pagina <= pdfDoc.GetNumberOfPages() Then
                                child.AddDestination(
                                iText.Kernel.Pdf.Navigation.PdfExplicitDestination.
                                CreateFit(pdfDoc.GetPage(pagina)))
                            End If

                            ' Registra nel dizionario per eventuali figli
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

    'Bottone carico indice
    Private Sub btnSfogliaIndice_Click(sender As Object, e As EventArgs) Handles btnSfogliaIndice.Click
        Using ofd As New OpenFileDialog()
            ofd.Title = "Seleziona file indice"
            ofd.Filter = "File di testo (*.txt)|*.txt|Tutti i file (*.*)|*.*"
            If ofd.ShowDialog() = DialogResult.OK Then
                txtIndicePath.Text = ofd.FileName
                _indiceBookmark = CaricaIndice(ofd.FileName)
            End If
        End Using
    End Sub

    'Funzione che legge
    Private Function CaricaIndice(path As String) As List(Of (Numero As String, Descrizione As String))
        Dim result As New List(Of (Numero As String, Descrizione As String))
        If Not File.Exists(path) Then Return result

        For Each line As String In File.ReadAllLines(path)
            Dim trimmed = line.Trim()
            If String.IsNullOrEmpty(trimmed) Then Continue For
            Dim parts = trimmed.Split(";"c)
            If parts.Length >= 2 Then
                Dim numero = parts(0).Trim()
                Dim desc = parts(1).Trim()
                result.Add((numero, desc))
            End If
        Next
        Return result
    End Function

    Private Function GetLivello(numero As String) As Integer
        ' "1.0" → 1 parte prima del punto → padre
        ' "1.0.0" → 2 parti dopo il primo punto → figlio
        Dim parts = numero.Split("."c)
        Return parts.Length - 1  ' 1.0=1, 1.0.0=2, 1.0.0.0=3
    End Function

    Private Function GetPadre(numero As String) As String
        ' "2.0.1" → padre è "2.0"
        Dim idx = numero.LastIndexOf("."c)
        If idx <= 0 Then Return ""
        Return numero.Substring(0, idx)
    End Function



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
