Imports iTextSharp.text.pdf
Imports System.Drawing
Public Class cBarCode

#Region "Code39"
    ''' <summary>
    ''' Devuelve la imagen de código de barras del número especificado
    ''' </summary>
    ''' <param name="Codigo">Número al que se le generará el código de barras</param>
    ''' <param name="Numero">True si se quiere que debajo del codigo de barras aparesca el numero y false de lo contrario</param>
    ''' <param name="Tamaño">Tamaño delcódigo de barra (Altura)</param>
    ''' <returns>Bitmap</returns>
    ''' <remarks></remarks>
    Public Function CodigoDeBarra(ByVal Codigo As String, Optional ByVal Numero As Boolean = False, Optional ByVal Tamaño As Integer = 50) As Bitmap
        Return GenerarBarCode(Codigo, Numero, Tamaño)
    End Function
    Private Shared Function GenerarBarCode(ByVal _code As String, Optional ByVal PrintTextInCode As Boolean = False, Optional ByVal Height As Single = 50, Optional ByVal GenerateChecksum As Boolean = False, Optional ByVal ChecksumText As Boolean = False) As Bitmap
        If _code.Trim = "" Then
            Return Nothing
        Else
            Dim barcode As New Barcode39
            barcode.StartStopText = True
            barcode.GenerateChecksum = GenerateChecksum
            barcode.ChecksumText = ChecksumText

            If Height <> 0 Then barcode.BarHeight = Height
            barcode.Code = _code
            Try
                Dim bm As New System.Drawing.Bitmap(barcode.CreateDrawingImage(System.Drawing.Color.Black, System.Drawing.Color.White))
                If PrintTextInCode = False Then
                    Return bm
                Else
                    Dim bmT As Image
                    bmT = New Bitmap(bm.Width, bm.Height + 14)
                    Dim g As Graphics = Graphics.FromImage(bmT)
                    g.FillRectangle(New SolidBrush(Color.White), 0, 0, bm.Width, bm.Height + 14)

                    Dim drawFont As New Font("Arial", 10)
                    Dim drawBrush As New SolidBrush(Color.Black)

                    Dim stringSize As New SizeF
                    stringSize = g.MeasureString(_code, drawFont)
                    Dim xCenter As Single = (bm.Width - stringSize.Width) / 2
                    Dim x As Single = xCenter
                    Dim y As Single = bm.Height

                    Dim drawFormat As New StringFormat
                    drawFormat.FormatFlags = StringFormatFlags.NoWrap

                    g.DrawImage(bm, 0, 0)
                    g.DrawString(_code, drawFont, drawBrush, x, y, drawFormat)

                    Return bmT
                End If

            Catch ex As Exception
                Throw New Exception("Error generating code39 barcode. Desc:" & ex.Message)
            End Try
        End If
    End Function
#End Region
End Class
