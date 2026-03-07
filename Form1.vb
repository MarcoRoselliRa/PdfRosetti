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

    ''' <summary>
    ''' Inizializza il form.
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.AllowDrop = True
        SetupGrid()

    End Sub

    ''' <summary>
    ''' Configura la griglia del report.
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
        dgvReport.Columns.Add("colDimensione", "Dimensione (mm)")
        dgvReport.Columns.Add("colBNColore", "BN/Colore")
        dgvReport.Columns.Add("colOrientamento", "Orientamento")
        dgvReport.Columns.Add("colDivisorio", "Divisorio")

    End Sub

    ''' <summary>
    ''' Gestisce drag enter.
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
    ''' Gestisce il rilascio del PDF.
    ''' </summary>
    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop

        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        If files.Length = 0 Then Exit Sub

        PdfPath = files(0)

        If Not File.Exists(PdfPath) Then Exit Sub

        txtPdfPath.Text = Path.GetFileName(PdfPath)

        If CurrentPdf IsNot Nothing Then
            CurrentPdf.Dispose()
            CurrentPdf = Nothing
        End If

        CurrentPdf = PdfiumDoc.Load(PdfPath)

        CaricaPagineInGriglia()

        AnalizzaColoriAsync()

        If dgvReport.Rows.Count > 0 Then
            dgvReport.ClearSelection()
            dgvReport.Rows(0).Selected = True
        End If

    End Sub

    ''' <summary>
    ''' Carica dimensioni e orientamento.
    ''' </summary>
    Private Sub CaricaPagineInGriglia()

        dgvReport.Rows.Clear()

        For i As Integer = 0 To CurrentPdf.PageCount - 1

            Dim sizePt = CurrentPdf.PageSizes(i)

            Dim wMm As Double = Math.Round(PtToMm(sizePt.Width), 1)
            Dim hMm As Double = Math.Round(PtToMm(sizePt.Height), 1)

            Dim orientamento As String = If(hMm >= wMm, "Portrait", "Landscape")

            Dim dimensione As String = $"{wMm} x {hMm}"

            dgvReport.Rows.Add((i + 1).ToString(), dimensione, "", orientamento, "")

        Next

    End Sub

    ''' <summary>
    ''' Converte punti PDF in millimetri.
    ''' </summary>
    Private Function PtToMm(pt As Double) As Double
        Return pt * 25.4 / 72.0
    End Function

    ''' <summary>
    ''' Aggiorna preview pagina selezionata.
    ''' </summary>
    Private Sub dgvReport_SelectionChanged(sender As Object, e As EventArgs) Handles dgvReport.SelectionChanged

        If CurrentPdf Is Nothing Then Exit Sub
        If dgvReport.SelectedRows.Count = 0 Then Exit Sub

        Dim pageIndex As Integer = dgvReport.SelectedRows(0).Index

        AggiornaAnteprimaPagina(pageIndex)

    End Sub

    ''' <summary>
    ''' Mostra l’anteprima della pagina e disegna le aree di ricerca del divisorio.
    ''' </summary>
    Private Sub AggiornaAnteprimaPagina(pageIndex As Integer)

        'smaltisce preview precedente
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

        'render pagina
        CurrentPreviewBmp = CurrentPdf.Render(pageIndex, outW, outH, 120, 120, PdfRenderFlags.Annotations)

        'cerchiamo il box nero prima di disegnare
        Dim blackBox = FindBlackDividerBox(CurrentPreviewBmp)

        'disegniamo overlay
        Using g As Graphics = Graphics.FromImage(CurrentPreviewBmp)

            'fascia rossa di ricerca
            Dim roiX As Integer = CInt(CurrentPreviewBmp.Width * 0.8)
            Dim roiY As Integer = 0
            Dim roiW As Integer = CInt(CurrentPreviewBmp.Width * 0.2)
            Dim roiH As Integer = CurrentPreviewBmp.Height

            g.DrawRectangle(Pens.Red, roiX, roiY, roiW, roiH)

            'se troviamo il box nero lo evidenziamo
            If blackBox.HasValue Then
                g.DrawRectangle(New Pen(Color.Lime, 3), blackBox.Value)
            End If
            If blackBox.HasValue Then
                g.DrawRectangle(New Pen(Color.Lime, 3), blackBox.Value)

                Dim numero As String = ReadDividerNumber(CurrentPreviewBmp, blackBox.Value)

                If numero <> "" Then
                    dgvReport.Rows(pageIndex).Cells("colDivisorio").Value = numero
                Else
                    dgvReport.Rows(pageIndex).Cells("colDivisorio").Value = ""
                End If
            End If
        End Using

        picPreview.Image = CurrentPreviewBmp

    End Sub


    ''' <summary>
    ''' Analizza BN / Colore in background.
    ''' </summary>
    Private Async Sub AnalizzaColoriAsync()

        Await Task.Run(
        Sub()

            Using reader As New PdfReader(PdfPath)
                Using pdf As New ITextPdfDoc(reader)

                    For i As Integer = 1 To pdf.GetNumberOfPages()

                        Dim risultato As String = "BN"

                        Try

                            Using bmp As Bitmap = RenderPageForColorAnalysis(i - 1)

                                Dim isColor As Boolean = IsColorPixel(bmp)

                                risultato = If(isColor, "Colore", "BN")

                            End Using

                        Catch
                            risultato = "ERR"
                        End Try

                        Dim rowIndex As Integer = i - 1

                        Me.BeginInvoke(
                            Sub()
                                If rowIndex < dgvReport.Rows.Count Then
                                    dgvReport.Rows(rowIndex).Cells("colBNColore").Value = risultato
                                End If
                            End Sub)

                    Next

                End Using
            End Using

        End Sub)

    End Sub

    ''' <summary>
    ''' Render pagina per analisi colore.
    ''' </summary>
    Private Function RenderPageForColorAnalysis(index As Integer) As Bitmap

        Dim sz = CurrentPdf.PageSizes(index)

        Dim maxSidePx As Integer = 1000
        Dim scale As Double = maxSidePx / CDbl(Math.Max(sz.Width, sz.Height))

        Dim w As Integer = Math.Max(1, CInt(sz.Width * scale))
        Dim h As Integer = Math.Max(1, CInt(sz.Height * scale))

        Return CurrentPdf.Render(index, w, h, 120, 120, PdfRenderFlags.Annotations)

    End Function

    ''' <summary>
    ''' Verifica presenza colore reale.
    ''' </summary>
    Private Function IsColorPixel(bmp As Bitmap) As Boolean

        For y As Integer = 0 To bmp.Height - 1 Step 4
            For x As Integer = 0 To bmp.Width - 1 Step 4

                Dim c As Color = bmp.GetPixel(x, y)

                Dim diff As Integer =
                    Math.Abs(c.R - c.G) +
                    Math.Abs(c.R - c.B) +
                    Math.Abs(c.G - c.B)

                If diff > 25 Then Return True

            Next
        Next

        Return False

    End Function

    ''' <summary>
    ''' Libera risorse.
    ''' </summary>
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

        If CurrentPreviewBmp IsNot Nothing Then
            CurrentPreviewBmp.Dispose()
        End If

        If CurrentPdf IsNot Nothing Then
            CurrentPdf.Dispose()
        End If

    End Sub




    ''' <summary>
    ''' Cerca il box nero del divisorio nella fascia destra della pagina.
    ''' </summary>
    Private Function FindBlackDividerBox(bmp As Bitmap) As Rectangle?

        Dim roiX As Integer = CInt(bmp.Width * 0.8)
        Dim roiW As Integer = CInt(bmp.Width * 0.18)

        If roiX < 0 Then roiX = 0
        If roiX >= bmp.Width Then Return Nothing

        If roiX + roiW > bmp.Width Then
            roiW = bmp.Width - roiX
        End If

        If roiW <= 0 Then Return Nothing

        Dim winH As Integer = Math.Max(60, CInt(bmp.Height * 0.08))
        Dim bestY As Integer = -1
        Dim bestRatio As Double = 0

        '1) troviamo la fascia più scura
        For y As Integer = 0 To Math.Max(0, bmp.Height - winH) Step 6

            Dim darkCount As Integer = 0
            Dim total As Integer = 0

            For yy As Integer = y To Math.Min(y + winH - 1, bmp.Height - 1) Step 2
                For xx As Integer = roiX To Math.Min(roiX + roiW - 1, bmp.Width - 1) Step 2

                    Dim c As Color = bmp.GetPixel(xx, yy)
                    Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3

                    If lum < 110 Then
                        darkCount += 1
                    End If

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

        '2) stringiamo il rettangolo sui pixel scuri reali
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

        '3) estendiamo verso l’alto
        Dim yTop As Integer = minY
        While yTop > 0

            Dim darkFound As Boolean = False

            For x As Integer = minX To maxX
                Dim c As Color = bmp.GetPixel(x, yTop)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3

                If lum < 100 Then
                    darkFound = True
                    Exit For
                End If
            Next

            If Not darkFound Then Exit While

            yTop -= 1
        End While

        '4) estendiamo verso il basso
        Dim yBottom As Integer = maxY
        While yBottom < bmp.Height - 1

            Dim darkFound As Boolean = False

            For x As Integer = minX To maxX
                Dim c As Color = bmp.GetPixel(x, yBottom)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3

                If lum < 100 Then
                    darkFound = True
                    Exit For
                End If
            Next

            If Not darkFound Then Exit While

            yBottom += 1
        End While

        '5) margine finale
        Dim pad As Integer = 6

        Dim rx As Integer = Math.Max(0, minX - pad)
        Dim ry As Integer = Math.Max(0, yTop - pad)

        Dim rw As Integer = Math.Min(bmp.Width - rx, (maxX - minX + 1) + pad * 2)
        Dim rh As Integer = Math.Min(bmp.Height - ry, (yBottom - yTop + 1) + pad * 2)

        Return New Rectangle(rx, ry, rw, rh)

    End Function


    ''' <summary>
    ''' Ritaglia una porzione di bitmap.
    ''' </summary>
    Private Function CropBitmap(src As Bitmap, r As Rectangle) As Bitmap

        Dim bmp As New Bitmap(r.Width, r.Height)

        Using g As Graphics = Graphics.FromImage(bmp)
            g.DrawImage(src, New Rectangle(0, 0, r.Width, r.Height), r, GraphicsUnit.Pixel)
        End Using

        Return bmp

    End Function



    ''' <summary>
    ''' Prepara il box del divisorio per l'OCR.
    ''' </summary>
    Private Function PrepareDividerForOcr(src As Bitmap) As Bitmap

        Dim rotated As Bitmap = CType(src.Clone(), Bitmap)

        'Questa rotazione va tenuta uguale a quella che ti ha mostrato il numero dritto
        rotated.RotateFlip(RotateFlipType.Rotate270FlipNone)
        rotated.RotateFlip(RotateFlipType.Rotate180FlipNone)

        Dim enlarged As New Bitmap(rotated.Width * 5, rotated.Height * 5, Imaging.PixelFormat.Format24bppRgb)

        Using g As Graphics = Graphics.FromImage(enlarged)
            g.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
            g.DrawImage(rotated, New Rectangle(0, 0, enlarged.Width, enlarged.Height))
        End Using

        rotated.Dispose()

        Dim outBmp As New Bitmap(enlarged.Width, enlarged.Height, Imaging.PixelFormat.Format24bppRgb)

        For y As Integer = 0 To enlarged.Height - 1
            For x As Integer = 0 To enlarged.Width - 1

                Dim c As Color = enlarged.GetPixel(x, y)
                Dim lum As Integer = (CInt(c.R) + CInt(c.G) + CInt(c.B)) \ 3

                'Threshold
                Dim v As Integer = If(lum < 170, 0, 255)

                'Inverti: numero nero su fondo bianco
                v = 255 - v

                outBmp.SetPixel(x, y, Color.FromArgb(v, v, v))
            Next
        Next

        enlarged.Dispose()

        Return outBmp

    End Function


    ''' <summary>
    ''' Legge il numero dentro il rettangolo divisorio separando le cifre.
    ''' </summary>
    Private Function ReadDividerNumber(pageBmp As Bitmap, rect As Rectangle) As String

        Using crop As Bitmap = CropBitmap(pageBmp, rect)
            Using prep As Bitmap = PrepareDividerForOcr(crop)

                Dim digitRects As List(Of Rectangle) = FindDigitComponents(prep)

                If digitRects.Count = 0 Then Return ""

                'Ordina da sinistra a destra
                digitRects = digitRects.OrderBy(Function(r) r.Left).ToList()

                Dim result As String = ""

                For Each r As Rectangle In digitRects
                    Using digitBmp As Bitmap = CropBitmap(prep, r)

                        Dim ch As String = ReadSingleDigit(digitBmp)

                        If ch <> "" Then
                            result &= ch
                        End If

                    End Using
                Next

                'Valida risultato finale
                If result.Length > 0 Then
                    Dim n As Integer
                    If Integer.TryParse(result, n) Then
                        If n >= 1 AndAlso n <= 30 Then
                            Return n.ToString()
                        End If
                    End If
                End If

            End Using
        End Using

        Return ""

    End Function
    ''' <summary>
    ''' Trova i componenti connessi neri che rappresentano le cifre.
    ''' </summary>
    Private Function FindDigitComponents(bmp As Bitmap) As List(Of Rectangle)

        Dim result As New List(Of Rectangle)

        Dim visited(bmp.Width - 1, bmp.Height - 1) As Boolean
        Dim q As New Queue(Of Point)

        For y As Integer = 0 To bmp.Height - 1
            For x As Integer = 0 To bmp.Width - 1

                If visited(x, y) Then Continue For
                visited(x, y) = True

                Dim c As Color = bmp.GetPixel(x, y)

                'nero
                If c.R > 50 Then Continue For

                Dim minX As Integer = x
                Dim maxX As Integer = x
                Dim minY As Integer = y
                Dim maxY As Integer = y
                Dim area As Integer = 0

                q.Clear()
                q.Enqueue(New Point(x, y))

                While q.Count > 0
                    Dim p As Point = q.Dequeue()
                    area += 1

                    If p.X < minX Then minX = p.X
                    If p.X > maxX Then maxX = p.X
                    If p.Y < minY Then minY = p.Y
                    If p.Y > maxY Then maxY = p.Y

                    Dim nx() As Integer = {p.X - 1, p.X + 1, p.X, p.X}
                    Dim ny() As Integer = {p.Y, p.Y, p.Y - 1, p.Y + 1}

                    For k As Integer = 0 To 3
                        Dim xx As Integer = nx(k)
                        Dim yy As Integer = ny(k)

                        If xx < 0 OrElse xx >= bmp.Width OrElse yy < 0 OrElse yy >= bmp.Height Then Continue For
                        If visited(xx, yy) Then Continue For

                        visited(xx, yy) = True

                        Dim cc As Color = bmp.GetPixel(xx, yy)
                        If cc.R < 50 Then
                            q.Enqueue(New Point(xx, yy))
                        End If
                    Next
                End While

                Dim w As Integer = maxX - minX + 1
                Dim h As Integer = maxY - minY + 1

                'Filtra rumore
                If area < 80 Then Continue For
                If w < 8 OrElse h < 20 Then Continue For
                If h < w Then Continue For

                result.Add(New Rectangle(minX, minY, w, h))

            Next
        Next

        'Tieni al massimo le 2 cifre più grandi
        result = result.OrderByDescending(Function(r) r.Height * r.Width).Take(2).ToList()

        Return result

    End Function


    ''' <summary>
    ''' Legge una singola cifra usando Tesseract.
    ''' </summary>
    Private Function ReadSingleDigit(bmp As Bitmap) As String

        Dim tempFile As String = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() & ".png")

        Try
            bmp.Save(tempFile, Imaging.ImageFormat.Png)

            Using engine As New Tesseract.TesseractEngine("./tessdata", "eng", Tesseract.EngineMode.Default)

                engine.SetVariable("tessedit_char_whitelist", "0123456789")
                engine.SetVariable("classify_bln_numeric_mode", "1")

                Using pixImg = Tesseract.Pix.LoadFromFile(tempFile)
                    Using page = engine.Process(pixImg, Tesseract.PageSegMode.SingleChar)

                        Dim txt As String = page.GetText()

                        If txt IsNot Nothing Then
                            txt = txt.Trim()

                            Dim digits As String =
                                New String(txt.Where(Function(c) Char.IsDigit(c)).ToArray())

                            If digits.Length > 0 Then
                                Return digits.Substring(0, 1)
                            End If
                        End If

                    End Using
                End Using
            End Using

        Catch
        Finally
            Try
                If File.Exists(tempFile) Then File.Delete(tempFile)
            Catch
            End Try
        End Try

        Return ""

    End Function
End Class