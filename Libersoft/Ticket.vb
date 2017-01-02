#Region "Importaciones"
Imports System.Drawing                  'Usado para el objeto Image
Imports System.Windows.Forms            'Usado para el DataGridView
Imports System.Drawing.Printing         'Usado para imprimir con PrintDocument

#End Region
Public Class Ticket
#Region "Declaraciones de Datos del ticket"
    '***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET *****************

    Private _Logotipo As Image = Nothing                            'Logotipo de la empresa         ID ->1
    Private _Empresa As String = "Aceros Inoxidables Refacciones y Equipos" 'Nombre de la empresa
    Private _Calle As String = "Calle Lino Merino #226"             'Nombre de la calle donde esta ubicada
    Private _Colonia As String = "Colonia Centro"                   'Nombre de la colonia
    Private _Ciudad As String = "Villahermosa Tab. Mex."            'Nombre del ciudad
    Private _Telefono As String = "314-99-06"                       'Telefono
    Private _CP As String = "86000"                                 'Código Postal
    Private _BarCode_Text As String = ""                            'Code 39
    Private _Barcode_Ima As Image = Nothing                         'Imagen del código de barra     ID ->0
    Private _Tabla As DataGridView = Nothing                        'Número del codigo de barra
    Private _Mensaje As String = "¡Gracias por su preferencia!"     'Mensaje de fin de ticket : "Gracias por su preferencia"
    Private _Total As String = "445.00"                             'Total dela venta
    Private _Correo As String = "plasticos_y_derivados@hotmail.com" 'Correo de la empresa
    Private _Cambio As String = "5.50"                              'Cambio de la venta
    Private _Efectivo As String = "500.00"                          'Efectivo con el que se pagó
    Private _Transaccion As String = "Operación"                    'Tipo de transacción que se está realizando
    Private _Tipo_pago As String = "Efectivo"                       'Forma de pago: Efectivo, Cheque
    Private _Fecha As String                                        'Fecha en que se registra la transacción
    Private _Hora As String                                         'Hora en que se registra la transacción

#End Region
#Region "Declaraciones de Funcionamiento de Impresión"
    '***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO *****
    Private WithEvents PD As New PrintDocument                      'Documento a imprimir
    Private PDBody As PrintPageEventArgs = Nothing                  'Cuerpo del documento
    Private _Art As String = "nombre_corto"                         'Indice de la columna articulo en el DataGridView
    Private _Cant As String = "cantidad"                            'Indice de la columna cantidad en el DataGridView
    Private _Sub As String = "subtotal"                             'Indice de la columna subtotal en el DataGridView
    Private _Precio As String = "precio"
    Private _Impresora As String = "POS-58"                         'Nombre de la impresora
    Private _Descuento As String = ""
    Private _ImagenPrint As Boolean = True                          'True imprime logotipo; false imprime código de barra
    Private _AnchoHoja As Decimal = 195                             'Ancho de la hoja de impresión
    Private _Espacio As Decimal = 5                                 'Espacio entre lineas
    Private _X As Integer = 0                                       'Posición X en la impresión
    Private _Y As Integer = 0                                       'Posición Y en la impresión
    Private AreaImpresion As Rectangle                              'Area de impresión
    Private Titulo_F As New Font("Arial", 12, FontStyle.Bold)       'Fuente de Titulo
    Private Encabezado_F As New Font("Arial", 9, FontStyle.Regular) 'Fuente de encabezado
    Private Cuerpo_F As New Font("Arial", 8, FontStyle.Regular)     'Fuente de cuerpo
    Private Columna_F As New Font("Arial", 8, FontStyle.Bold)       'Fuente de columna
    Private eCenter As New StringFormat()                           'Centra el texto
    Private eLeft As New StringFormat()                             'Alineación a la izquierda
    Private eRight As New StringFormat()                            'Alineación a la derecha
    Private _Aling As StringFormat                                  'Auxiliar para alineación
    Private _PrintView As New PrintPreviewDialog
#End Region

    'Contructor de clase
    Public Sub New()
        eCenter.Alignment = StringAlignment.Center
        eCenter.LineAlignment = StringAlignment.Center
        eLeft.Alignment = StringAlignment.Near
        eLeft.LineAlignment = StringAlignment.Center
        eRight.Alignment = StringAlignment.Far
        eRight.LineAlignment = StringAlignment.Center
    End Sub


#Region "Propiedades"
    ''' <summary>
    ''' Tipo de pago
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property TipoPago As String
        Get
            Return _Tipo_pago
        End Get
        Set(value As String)
            _Tipo_pago = value
        End Set
    End Property
    ''' <summary>
    ''' Fecha en que se registra la transacción
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Fecha As String
        Get
            Return _Fecha
        End Get
        Set(value As String)
            If Not value = "" Then
                _Fecha = value
            Else
                _Fecha = Format(CDate(Date.Today), "dd/MM/yyyy").ToString
            End If

        End Set
    End Property
    ''' <summary>
    ''' Hora en que se registra la transacción
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Hora As String
        Get
            Return _Hora
        End Get
        Set(value As String)
            If Not value = "" Then
                _Hora = value
            Else
                _Hora = Format(CDate(Date.Today), "hh:mm tt")
            End If

        End Set
    End Property
    ''' <summary>
    ''' Descuento total en la venta
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Descuento As String
        Get
            Return _Descuento
        End Get
        Set(value As String)
            _Descuento = value
        End Set
    End Property
    ''' <summary>
    ''' Transacción que se está realizando: Venta, Cotización
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Transaccion As String
        Get
            Return _Transaccion
        End Get
        Set(value As String)
            _Transaccion = value
        End Set
    End Property
    ''' <summary>
    ''' Efectivo con el cual está pagando el cliente
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Efectivo As String
        Get
            Return _Efectivo
        End Get
        Set(value As String)
            _Efectivo = value
        End Set
    End Property
    ''' <summary>
    ''' Espacio entre lineas
    ''' </summary>
    ''' <value>Decimal</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Espacio As String
        Get
            Return _Espacio
        End Get
        Set(value As String)
            _Espacio = value
        End Set
    End Property
    ''' <summary>
    ''' Ancho de la hoja de impresión en puntos
    ''' </summary>
    ''' <value>Decimal</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Ancho_Hoja As String
        Get
            Return _AnchoHoja
        End Get
        Set(value As String)
            _AnchoHoja = value
        End Set
    End Property
    ''' <summary>
    ''' Cambio que se le devlverá al cliente
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Cambio As String
        Get
            Return _Cambio
        End Get
        Set(value As String)
            _Cambio = value
        End Set
    End Property
    ''' <summary>
    ''' Correo electrónico de la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Property Correo As String
        Get
            Return _Correo
        End Get
        Set(value As String)
            _Correo = value
        End Set
    End Property
    ''' <summary>
    ''' Nombre de la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Empresa As String
        Get
            Return _Empresa
        End Get
        Set(value As String)
            _Empresa = value
        End Set
    End Property
    ''' <summary>
    ''' Indice de la columna articulo en el DataGridView
    ''' </summary>
    ''' <value>Integer</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IndexArticulo As String
        Get
            Return _Art
        End Get
        Set(value As String)
            _Art = value
        End Set
    End Property
    ''' <summary>
    ''' Indice de la columna cantidad en el DataGridView
    ''' </summary>
    ''' <value>Integer</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IndexCantidad As String
        Get
            Return _Cant
        End Get
        Set(value As String)
            _Cant = value
        End Set
    End Property
    ''' <summary>
    ''' Indice de la columna articulo en el DataGridView
    ''' </summary>
    ''' <value>Integer</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property IndexSubTotal As String
        Get
            Return _Sub
        End Get
        Set(value As String)
            _Sub = value
        End Set
    End Property
    ''' <summary>
    ''' Nombre de la calle donde está ubicada la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Calle As String
        Get
            Return _Calle
        End Get
        Set(value As String)
            _Calle = value
        End Set
    End Property
    ''' <summary>
    ''' Nombre de la colonia donde está ubicada la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Colonia As String
        Get
            Return _Colonia
        End Get
        Set(value As String)
            _Colonia = value
        End Set
    End Property
    ''' <summary>
    ''' Nombre de la ciudad donde está ubicada la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Municipio As String
        Get
            Return _Ciudad
        End Get
        Set(value As String)
            _Ciudad = value
        End Set
    End Property
    ''' <summary>
    ''' Logotipo de la empresa
    ''' </summary>
    ''' <value>Image</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Logotipo As Image
        Get
            Return _Logotipo
        End Get
        Set(value As Image)
            _Logotipo = value
        End Set
    End Property
    ''' <summary>
    ''' Telefono de la empresa
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Telefono As String
        Get
            Return _Telefono
        End Get
        Set(value As String)
            _Telefono = value
        End Set
    End Property
    ''' <summary>
    ''' Texto/Número de código de barra
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BarCode_Text As String
        Get
            Return _BarCode_Text
        End Get
        Set(value As String)
            _BarCode_Text = value
        End Set
    End Property
    ''' <summary>
    ''' Imágen de código de barras
    ''' </summary>
    ''' <value>Image</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BarCode_Ima As Image
        Get
            Return _Barcode_Ima
        End Get
        Set(value As Image)
            _Barcode_Ima = value
        End Set
    End Property
    ''' <summary>
    ''' DataGridView donde esta el cuerpo del contenido del Ticket
    ''' </summary>
    ''' <value>DataGridView</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Tabla As DataGridView
        Get
            Return _Tabla
        End Get
        Set(value As DataGridView)
            _Tabla = value
        End Set
    End Property
    ''' <summary>
    ''' Código postal
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CodigoPostal As String
        Get
            Return _CP
        End Get
        Set(value As String)
            _CP = value
        End Set
    End Property
    ''' <summary>
    ''' Nombre de la impresora
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Impresora As String
        Get
            Return _Impresora
        End Get
        Set(value As String)
            _Impresora = value
        End Set
    End Property
    ''' <summary>
    ''' Mensaje de fin de ticket. Ej. : ¡Gracias por su preferencia!
    ''' </summary>
    ''' <value>String</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Mensaje As String
        Get
            Return _Mensaje
        End Get
        Set(value As String)
            _Mensaje = value
        End Set
    End Property

#End Region
#Region "Funciones públicas"
    ''' <summary>
    ''' Imprime ticket
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ImprimirTicket()
        Imprimir()
    End Sub
    ''' <summary>
    ''' Convierte un número a texto especificando los pesos y centavos
    ''' </summary>
    ''' <param name="Num">Número a convertir a texto</param>
    ''' <returns>String</returns>
    ''' <remarks></remarks>
    Public Function NumeroToTexto(ByVal Num As Decimal) As String
        Dim Num2 As Integer
        Dim Cadena As String
        Num2 = Fix(Num)
        Cadena = NumToTex(Num2) + "pesos"
        Num2 = (Num - Num2) * 100
        If Num2 Then
            Cadena = Cadena + " y " + NumToTex(Num2) + "centavos"
        End If
        Return StrConv(Cadena, VbStrConv.ProperCase)
    End Function
#End Region
#Region "Funciones privadas"
#Region "Funciones generales"
    ''' <summary>
    ''' Recibe como parametro un numero y devuelve el numero expresado en cadena
    ''' </summary>
    ''' <param name="Num">Número a convertir a texto</param>
    ''' <returns>Cadena</returns>
    ''' <remarks></remarks>
    Private Function NumToTex(ByVal Num As String) As String
        Dim Num2 As Integer
        Dim Num3 As Integer
        Dim Cd As String = ""
        Dim ant As Boolean = False
        Num2 = Num
        While Not Num2 = 0
            If Num2 > 999 Then
                If Num2 < 2000 Then
                    Cd = "mil "
                ElseIf Num2 < 20000 Then
                    Cd = Numero(CInt(Num2 \ 1000)) + " " + Numero(1000) + " "
                Else
                    Num3 = CInt(Num2 \ 1000)
                    Cd = Numero(CInt(Num3 \ 10) * 10) + " "
                    Num3 = Num3 Mod 10
                    If Num3 Then
                        Cd = Cd + " y " + Numero(Num3) + " mil "
                    Else
                        Cd = Cd + " mil "
                    End If
                End If
                Num2 = Num2 Mod 1000
            ElseIf Num2 > 99 Then
                If Num2 = 100 Then
                    Cd = Cd + "cien "
                Else
                    Cd = Cd + Numero(CInt(Num2 \ 100) * 100) + " "
                End If
                Num2 = Num2 Mod 100
            ElseIf Num2 > 19 Then
                Cd = Cd + Numero(CInt(Num2 \ 10) * 10) + " "
                Num2 = Num2 Mod 10
                ant = True
            Else
                If ant Then
                    Cd = Cd + "y " + Numero(Num2) + " "
                Else
                    Cd = Cd + Numero(Num2) + " "
                End If

                Num2 = 0
            End If
        End While

        Return Cd
    End Function
    Private Function Numero(ByVal Num As Integer) As String
        Select Case Num
            Case 1
                Return "uno"
            Case 2
                Return "dos"
            Case 3
                Return "tres"
            Case 4
                Return "cuatro"
            Case 5
                Return "cinco"
            Case 6
                Return "seis"
            Case 7
                Return "siete"
            Case 8
                Return "ocho"
            Case 9
                Return "nueve"
            Case 10
                Return "diez"
            Case 11
                Return "once"
            Case 12
                Return "doce"
            Case 13
                Return "trece"
            Case 14
                Return "catorce"
            Case 15
                Return "quince"
            Case 16
                Return "dieciséis"
            Case 17
                Return "diecisiete"
            Case 18
                Return "dieciocho"
            Case 19
                Return "diecinueve"
            Case 20
                Return "veinte"
            Case 30
                Return "treinta"
            Case 40
                Return "cuarenta"
            Case 50
                Return "cincuenta"
            Case 60
                Return "sesenta"
            Case 70
                Return "setenta"
            Case 80
                Return "ochenta"
            Case 90
                Return "noventa"
            Case 100
                Return "ciento"
            Case 200
                Return "doscientos"
            Case 300
                Return "trescientos"
            Case 400
                Return "cuatrocientos"
            Case 500
                Return "quinientos"
            Case 600
                Return "seiscientos"
            Case 700
                Return "setecientos"
            Case 800
                Return "ochocientos"
            Case 900
                Return "novecientos"
            Case 1000
                Return "mil"
        End Select
        Return ""
    End Function
    Private Function DistribuirCadenas(ByVal Cadena1 As String, ByVal Cadena2 As String, ByVal Cadena3 As String) As String
        Dim C1L As Integer
        Dim C2L As Integer
        Dim i As Integer = 1
        Dim Aux1 As String = ""

        C2L = Cadena3.Length
        C1L = Cadena1.Length
        If C1L > 20 Then
            Cadena1 = Cadena1.Substring(0, 20) + " "
        Else
            Cadena1 = Cadena1 + "                    "
            Cadena1 = Cadena1.Substring(0, 21)
        End If

        C1L = Cadena2.Length
        If C1L = 1 Then
            Cadena2 = " " + Cadena2
            C1L = Cadena2.Length
        End If

        If (C1L + C2L) < 12 Then
            For i = 1 To (11 - (C1L + C2L))
                Aux1 = Aux1 + " "
            Next
            Cadena3 = Aux1 + Cadena3
        End If
        Cadena1 = Cadena1 + Cadena2 + Cadena3
        Return Cadena1
    End Function
#End Region
#Region "Operaciones basicas con la impresora"
    Private Function Imprimir() As Boolean
        Try
            If PD.PrinterSettings.IsValid Then
                PD.DocumentName = "Ticket"
                PD.PrinterSettings.PrinterName = _Impresora
                PD.PrintController = New StandardPrintController

                AddHandler PD.PrintPage, AddressOf PrintDocu_PrintPage

                '_PrintView.Document = PD
                '_PrintView.Show()
                PD.Print()
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox("¡Error al intentar imprimir!: " + ex.ToString, vbCritical)
            Return False
        End Try
    End Function
    Private Sub PrintDocu_PrintPage(sender As Object, e As PrintPageEventArgs)
        StartPrint(e)
        If Not IsNothing(_Logotipo) Then
            PrintImage(_Logotipo)
        End If
        If Not _Calle = "" Then
            PrintText(_Calle, Encabezado_F)
        End If
        If Not _Colonia = "" Then
            PrintText(_Colonia + " C.P. " + _CP, Encabezado_F)
        End If
        If Not _Ciudad = "" Then
            PrintText(_Ciudad, Encabezado_F)
        End If
        If Not _Telefono = "" Then
            PrintText("Tel. " + _Telefono, Encabezado_F)
        End If
        If Not _Correo = "" Then
            PrintText("E-mail: " + _Correo, Cuerpo_F)
        End If
        If Not _Correo = "" Then
            PrintText("Fecha: " + Format(CDate(_Fecha), "dd/MM/yyyy").ToString, Cuerpo_F)
        End If
        If Not _Correo = "" Then
            PrintText("Hora: " + Format(CDate(_Hora), "hh:mm tt").ToString, Cuerpo_F)
        End If
        eSpace(2)
        If Not IsNothing(_Barcode_Ima) And Not _Transaccion = "Cotización" Then
            PrintText("Folio: " & _BarCode_Text, Columna_F)
        End If
        eSpace(2)
        PrintText("Transacción: " + _Transaccion, Columna_F)
        PrintBody()
        If Not _Total = "" Then
            PrintText(NumeroToTexto(_Total), Cuerpo_F)
        End If
        eSpace(2)
        If Not _Transaccion = "Cotización" Then
            PrintText("Pago en " + _Tipo_pago + " una sola exibición", Columna_F)
            eSpace(2)
            If TipoPago = "Cheque" Then
                PrintText("Cheque: $" + Format((_Efectivo * 1), "##,##0.00"), Columna_F)
            Else
                PrintText("Efectivo: $" + Format((_Efectivo * 1), "##,##0.00"), Columna_F)
                PrintText("Cambio: $" + Format((_Efectivo - _Total), "##,##0.00"), Columna_F)
            End If

            eSpace(2)
        End If

        'If Not _Descuento = "" And Not -Descuento = "0" Then
        '  PrintText("Con el descuento aplicado has ahorrado $" + _Descuento, Columna_F)
        '   eSpace(2)
        'End If

        If Not IsNothing(_Barcode_Ima) And Not _Transaccion = "Cotización" Then
            PrintImage(_Barcode_Ima)
            eSpace(2)
        End If
        If Not _Mensaje = "" Then
            PrintText(_Mensaje, Columna_F)
        End If
        eSpace(6)
        PrintText(".", Cuerpo_F)
        e = EndPrint()
        'PD.Dispose()
    End Sub
    Private Sub PrintBody()
        Dim X1 As Integer
        Dim X2 As Integer
        Dim X3 As Integer
        Dim X4 As Integer
        Dim W1 As Integer
        Dim W2 As Integer
        Dim W3 As Integer
        Dim W4 As Integer
        Dim TF As Decimal
        Dim Lineas As Integer
        Dim Total_ As Decimal = 0
        If Not IsNothing(_Tabla) Then
            Lineas = _Tabla.RowCount
            TF = Cuerpo_F.GetHeight(PDBody.Graphics)
            W4 = PDBody.Graphics.MeasureString("88000", Cuerpo_F).Width
            W3 = W4
            W2 = PDBody.Graphics.MeasureString("500", Cuerpo_F).Width
            W1 = _AnchoHoja - (W4 + W2 + W3 + 10)
            X1 = 0
            X2 = W1 + 5
            X3 = X2 + W2
            X4 = X3 + W3

            eLine()
            AreaImpresion = New Rectangle(X1, _Y, W1, TF)
            PDBody.Graphics.DrawString("Articulo:", Columna_F, Brushes.Black, AreaImpresion, eLeft)
            AreaImpresion = New Rectangle(X2, _Y, W2, TF)
            PDBody.Graphics.DrawString("Cant.", Columna_F, Brushes.Black, AreaImpresion, eCenter)
            AreaImpresion = New Rectangle(X3, _Y, W3, TF)
            PDBody.Graphics.DrawString("P.U.", Columna_F, Brushes.Black, AreaImpresion, eCenter)
            AreaImpresion = New Rectangle(X4, _Y, W3, TF)
            PDBody.Graphics.DrawString("SubT", Columna_F, Brushes.Black, AreaImpresion, eLeft)
            _Y += 10
            eLine()
            Total_ = 0
            For i As Integer = 0 To Lineas - 1
                _Y += TF
                AreaImpresion = New Rectangle(X1, _Y, W1, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Art, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eLeft)
                AreaImpresion = New Rectangle(X2, _Y, W2, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Cant, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eCenter)
                AreaImpresion = New Rectangle(X3, _Y, W3, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Precio, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eCenter)
                AreaImpresion = New Rectangle(X4, _Y, W4, TF)
                PDBody.Graphics.DrawString(_Tabla.Item(_Sub, i).Value, Cuerpo_F, Brushes.Black, AreaImpresion, eLeft)
                Total_ += _Tabla.Item(_Sub, i).Value
            Next
            _Total = Format((Total_ * 1), "##,##0.00")
            _Y += TF
            AreaImpresion = New Rectangle(X2 - 20, _Y, _AnchoHoja - (X2 - 20), Columna_F.GetHeight(PDBody.Graphics))
            PDBody.Graphics.DrawString("Total: $" + _Total, Columna_F, Brushes.Black, AreaImpresion, eLeft)

            _Y += Columna_F.GetHeight(PDBody.Graphics) + 10
            eLine()

        End If


    End Sub
    ''' <summary>
    ''' Agrega una linea al documento. Alineación: 0-> Izquierda; 1->Centro; 2-> Derecha
    ''' </summary>
    ''' <param name="Texto">Texto a imprimir</param>
    ''' <param name="Fuente_F">Titulo; Encabezado; Cuerpo; Columna</param>
    ''' <param name="Alineacion">Alineación: 0-> Izquierda; 1->Centro; 2-> Derecha</param>
    ''' <remarks></remarks>
    Private Sub PrintText(ByVal Texto As String, ByVal Fuente_F As Font, Optional ByVal Alineacion As Integer = 1)
        Dim TFuente As Decimal = 12
        If Not IsNothing(PDBody) Then
            Select Case Alineacion
                Case 0
                    _Aling = eLeft
                Case 1
                    _Aling = eCenter
                Case 2
                    _Aling = eRight
                Case Else
                    _Aling = eCenter
            End Select
            TFuente = PDBody.Graphics.MeasureString(Texto, Fuente_F).Height
            If PDBody.Graphics.MeasureString(Texto, Fuente_F).Width > Ancho_Hoja Then
                If (PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja) Mod 1 Then
                    TFuente = TFuente * (Fix((PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja) + 1))
                Else
                    TFuente = TFuente * (PDBody.Graphics.MeasureString(Texto, Fuente_F).Width / _AnchoHoja)
                End If
            Else
                Fuente_F.GetHeight(PDBody.Graphics)
            End If
            AreaImpresion = New Rectangle(_X, _Y, Ancho_Hoja, TFuente)
            PDBody.Graphics.DrawString(Texto, Fuente_F, Brushes.Black, AreaImpresion, _Aling)
            _Y = _Y + TFuente
        Else
            MsgBox("¡No se ha indicado el inicio de documento al crear el Ticket!", vbOKOnly + vbExclamation, "Ticket")
        End If
    End Sub

    Private Sub PrintImage(ByVal Imagen As Image)
        Dim ImagenW As Integer
        Dim ImagenH As Integer
        Dim XPos As Integer

        ImagenH = Imagen.Height
        ImagenW = Imagen.Width
        If _AnchoHoja > ImagenW Then
            XPos = (_AnchoHoja - ImagenW) \ 2
        Else
            XPos = 0
        End If
        AreaImpresion = New Rectangle(XPos, _Y, ImagenW, ImagenH)
        PDBody.Graphics.DrawImage(Imagen, AreaImpresion)
        _Y = _Y + ImagenH

    End Sub
    ''' <summary>
    ''' Indica el termino de un impresión
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EndPrint() As PrintPageEventArgs
        PDBody.HasMorePages = False
        Return PDBody
    End Function
    Private Sub eSpace(Optional ByVal num As Integer = 1)
        Dim Fuente_F As New Font("Arial", 10, FontStyle.Regular)
        Dim TFuente As Decimal
        'Dim Texto As String = "- - - - - - - - - - - - - - - - - - - - - - - - - - - -"

        TFuente = Fuente_F.GetHeight(PDBody.Graphics)
        AreaImpresion = New Rectangle(_X, _Y, Ancho_Hoja, _Espacio * num)
        PDBody.Graphics.DrawRectangle(Pens.White, AreaImpresion)
        _Y = _Y + (num * _Espacio)
    End Sub
    Private Sub eLine()
        Dim Fuente_F As New Font("Arial", 10, FontStyle.Regular)
        Dim TFuente As Decimal
        Dim Texto As String = "- - - - - - - - - - - - - - - - - - - - - - - - - - - -"
        TFuente = Fuente_F.GetHeight(PDBody.Graphics)
        AreaImpresion = New Rectangle(_X, _Y, Ancho_Hoja, TFuente)
        PDBody.Graphics.DrawString(Texto, Fuente_F, Brushes.Black, AreaImpresion, _Aling)
        _Y = _Y + TFuente
    End Sub
    ''' <summary>
    ''' Indica el inicio de la creación de un documento
    ''' </summary>
    ''' <param name="e">PrintPageEventArgs</param>
    ''' <remarks></remarks>
    Private Sub StartPrint(ByVal e As PrintPageEventArgs)
        'PDBody = Nothing
        PDBody = e
    End Sub
#End Region
#End Region
End Class
