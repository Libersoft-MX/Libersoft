#Region "Importaciones"
Imports System.Drawing                  'Usado para el objeto Image
Imports System.Windows.Forms            'Usado para el DataGridView
Imports System.Drawing.Printing         'Usado para imprimir con PrintDocument

#End Region
Public Class cTicket
#Region "Declaraciones de Datos del ticket"
    '***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET *****************

    Private _Logotipo As Image = Nothing                            'Logotipo de la empresa         ID ->1
    Private _Empresa As String = "Aceros Inoxidablestyzxc Refacciones y Equipos" 'Nombre de la empresa
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

#End Region
#Region "Declaraciones de Funcionamiento de Impresión"
    '***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO *****
    Private WithEvents PD As New PrintDocument                      'Documento a imprimir
    Private PDBody As PrintPageEventArgs = Nothing                  'Cuerpo del documento
    Private _Art As Integer = 0                                     'Indice de la columna articulo en el DataGridView
    Private _Cant As Integer = 1                                    'Indice de la columna cantidad en el DataGridView
    Private _Sub As Integer = 2                                     'Indice de la columna subtotal en el DataGridView
    Private _Impresora As String = "POS58"                          'Nombre de la impresora
    Private _ImagenPrint As Boolean = True                          'True imprime logotipo; false imprime código de barra
    Private _AnchoHoja As Decimal = 194                             'Ancho de la hoja de impresión
    Private _Espacio As Decimal = 15                                'Espacio entre lineas
    Private _X As Integer = 0                                       'Posición X en la impresión
    Private _Y As Integer = 0                                       'Posición Y en la impresión
    Dim AreaImpresion As Rectangle                                  'Area de impresión
    Dim F_Titulo As New Font("Arial", 14, FontStyle.Bold)           'Fuente de Titulo
    Dim F_Encabezado As New Font("Arial", 9, FontStyle.Regular)     'Fuente de encabezado
    Dim F_Cuerpo As New Font("Arial", 8, FontStyle.Regular)         'Fuente de cuerpo
    Dim F_Columna As New Font("Arial", 8, FontStyle.Bold)           'Fuente de columna
    Dim aCenter As New StringFormat()                               'Centra el texto
    Dim aLeft As New StringFormat()                                 'Alineación a la izquierda
    Dim aRight As New StringFormat()                                'Alineación a la derecha
#End Region

    'Contructor de clase
    Public Sub New()
        aCenter.Alignment = StringAlignment.Center
        aCenter.LineAlignment = StringAlignment.Center
        aLeft.Alignment = StringAlignment.Near
        aLeft.LineAlignment = StringAlignment.Center
        aRight.Alignment = StringAlignment.Far
        aRight.LineAlignment = StringAlignment.Center
    End Sub


#Region "Propiedades"
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
    Public Property Logo As Image
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
            PD.PrinterSettings.PrinterName = _Impresora
            PD.PrintController = New StandardPrintController
            If PD.PrinterSettings.IsValid Then
                PD.DocumentName = "Ticket"
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
    Private Sub PrintDocu_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PD.PrintPage
        StartPrint(e)


        Print_(_Empresa)


        e = EndPrint()
    End Sub
    ''' <summary>
    ''' Agrega una linea al documento
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Print_(ByVal Texto As String)
        If Not IsNothing(PDBody) Then
            AreaImpresion = New Rectangle(_X, _Y, Ancho_Hoja, 22)
            PDBody.Graphics.DrawString(Texto, F_Cuerpo, Brushes.Black, AreaImpresion, aCenter)
        Else
            MsgBox("¡No se ha indicado el inicio de documento al crear el Ticket!", vbOKOnly + vbExclamation, "Ticket")
        End If
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
    ''' <summary>
    ''' Indica el inicio de la creación de un documento
    ''' </summary>
    ''' <param name="e">PrintPageEventArgs</param>
    ''' <remarks></remarks>
    Private Sub StartPrint(ByVal e As PrintPageEventArgs)
        PDBody = Nothing
        PDBody = e
    End Sub
#End Region
#End Region
End Class
