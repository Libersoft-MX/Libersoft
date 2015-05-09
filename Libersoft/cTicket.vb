#Region "Importaciones"
Imports System.Drawing                  'Usado para el objeto Image
Imports System.Windows.Forms            'Usado para el DataGridView
Imports System.Drawing.Printing         'Usado para imprimir con PrintDocument
Imports System.Runtime.InteropServices  'Usado para imprimir ESC/POS

#End Region
Public Class cTicket

#Region "Declaraciones de Datos del ticket"
    '***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET ***** DATOS DEL TICKET *****
    Private _Logotipo As Image          'Logotipo de la empresa         ID ->1
    Private _Empresa As String          'Nombre de la empresa
    Private _Calle As String            'Nombre de la calle donde esta ubicada
    Private _Colonia As String          'Nombre de la colonia
    Private _Ciudad As String           'Nombre del ciudad
    Private _Telefono As String         'Telefono
    Private _CP As String = ""          'Código Postal
    Private _BarCode_Text As String     'Code 39
    Private _Barcode_Ima As Image       'Imagen del código de barra     ID ->0
    Private _Tabla As DataGridView      'Número del codigo de barra
    Private _Impresora As String        'Nombre de la impresora
    Private _Mensaje As String          'Mensaje de fin de ticket : "Gracias por su preferencia"

#End Region
#Region "Declaraciones de Funcionamiento"
    '***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO ***** FUNCIONAMIENTO *****
    Private WithEvents PrintDocument1 As PrintDocument
    Private ImagenPrint As Boolean

    '*************************************
    Private Const eClear As String = Chr(27) + "@"
    Private Const eCentre As String = Chr(27) + Chr(97) + "1"
    Private Const eLeft As String = Chr(27) + Chr(97) + "0"
    Private Const eRight As String = Chr(27) + Chr(97) + "2"
    Private Const eDrawer As String = eClear + Chr(27) + "p" + Chr(0) + ".}"
    Private Const eCut As String = Chr(27) + "i" + vbCrLf
    Private Const eSmlText As String = Chr(27) + "!" + Chr(1)
    Private Const eNmlText As String = Chr(27) + "!" + Chr(0)
    Private Const eInit As String = eNmlText + Chr(13) + Chr(27) + _
    "c6" + Chr(1) + Chr(27) + "R3" + vbCrLf
    Private Const eBigCharOn As String = Chr(27) + "!" + Chr(56)
    Private Const eBigCharOff As String = Chr(27) + "!" + Chr(0)

    Private hPrinter As New IntPtr(0)
    Private di As New DOCINFOA()
    Private PrinterOpen As Boolean = False
    '*************************************
#End Region
#Region "Propiedades"
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

#End Region
#Region "Funciones privadas"
#Region "Funciones generales"
    Private Function Imprimir_Cuerpo() As Boolean
        Return False
    End Function
    Private Function Imprimir_Imagen(ByVal Ima As Boolean) As Boolean
        Dim Imagen_item As Image
        If Ima Then
            Imagen_item = _Logotipo
        Else
            Imagen_item = _Barcode_Ima
        End If

        Try
            PrintDocument1 = New PrintDocument
            PrintDocument1.PrinterSettings.PrinterName = _Impresora
            PrintDocument1.PrintController = New StandardPrintController
            If PrintDocument1.PrinterSettings.IsValid Then
                PrintDocument1.DocumentName = "Ticket"
                PrintDocument1.Print()
            End If
            Return True
        Catch ex As Exception
            MsgBox("Error al intentar imprimir imágen : " + ex.ToString, vbCritical)
            Return False
        End Try
    End Function
    Private Sub PrintDocu_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        If ImagenPrint Then
            e.Graphics.DrawImage(_Logotipo, 0, 0)
        Else
            e.Graphics.DrawImage(_Barcode_Ima, 60, 0)
        End If
    End Sub
#End Region
#Region "Operaciones basicas con la impresora"
    Private Sub StartPrint()
        OpenPrint(_Impresora)
    End Sub

    Private Sub PrintHeader()
        'eInit Espacio en blanco
        If Not Telefono = "" Then
            Print(eCentre + _Telefono)
        End If
        Print(_Calle + " C.P. " + _CP)
        Print(_Colonia)
        Print(_Ciudad + eLeft)

        PrintLinea()
    End Sub

    Private Sub PrintFooter()
        PrintLinea()
        Print(eCentre + _Mensaje + eLeft)
        Print(vbLf + vbCrLf + eCut + eDrawer)
    End Sub

    Private Sub Print(Line As String)
        SendStringToPrinter(_Impresora, Line + vbLf)
    End Sub

    Private Sub PrintLinea()
        Print(eInit + "- - - - - - - - - - - - - - - -" + eInit)
    End Sub

    Private Sub EndPrint()
        ClosePrint()
    End Sub
#End Region

#End Region
#Region "RawPrinterHelper"
    ' Structure and API declarations:
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> _
    Private Class DOCINFOA
        <MarshalAs(UnmanagedType.LPStr)> _
        Public pDocName As String
        <MarshalAs(UnmanagedType.LPStr)> _
        Public pOutputFile As String
        <MarshalAs(UnmanagedType.LPStr)> _
        Public pDataType As String
    End Class

    <DllImport("winspool.Drv", EntryPoint:="OpenPrinterA", _
    SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True, _
    CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function OpenPrinter(<MarshalAs(UnmanagedType.LPStr)> _
    szPrinter As String, ByRef hPrinter As IntPtr, pd As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="ClosePrinter", _
    SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function ClosePrinter(hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="StartDocPrinterA", _
    SetLastError:=True, CharSet:=CharSet.Ansi, ExactSpelling:=True, _
    CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartDocPrinter(hPrinter As IntPtr, level As Int32, _
    <[In](), MarshalAs(UnmanagedType.LPStruct)> di As DOCINFOA) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="EndDocPrinter", _
    SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndDocPrinter(hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="StartPagePrinter", _
    SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function StartPagePrinter(hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="EndPagePrinter", _
    SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function EndPagePrinter(hPrinter As IntPtr) As Boolean
    End Function

    <DllImport("winspool.Drv", EntryPoint:="WritePrinter", _
    SetLastError:=True, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)> _
    Private Shared Function WritePrinter(hPrinter As IntPtr, pBytes As IntPtr, _
    dwCount As Int32, ByRef dwWritten As Int32) As Boolean
    End Function

    Private ReadOnly Property PrinterIsOpen As Boolean
        Get
            PrinterIsOpen = PrinterOpen
        End Get
    End Property

    Private Function OpenPrint(szPrinterName As String) As Boolean
        If PrinterOpen = False Then
            di.pDocName = ".NET RAW Document"
            di.pDataType = "RAW"

            If OpenPrinter(szPrinterName.Normalize(), hPrinter, IntPtr.Zero) Then
                ' Start a document.
                If StartDocPrinter(hPrinter, 1, di) Then
                    If StartPagePrinter(hPrinter) Then
                        PrinterOpen = True
                    End If
                End If
            End If
        End If

        OpenPrint = PrinterOpen
    End Function

    Private Sub ClosePrint()
        If PrinterOpen Then
            EndPagePrinter(hPrinter)
            EndDocPrinter(hPrinter)
            ClosePrinter(hPrinter)
            PrinterOpen = False
        End If
    End Sub

    Private Function SendStringToPrinter(szPrinterName As String, szString As String) As Boolean
        If PrinterOpen Then
            Dim pBytes As IntPtr
            Dim dwCount As Int32
            Dim dwWritten As Int32 = 0
            dwCount = szString.Length
            pBytes = Marshal.StringToCoTaskMemAnsi(szString)
            SendStringToPrinter = WritePrinter(hPrinter, pBytes, dwCount, dwWritten)
            Marshal.FreeCoTaskMem(pBytes)
        Else
            SendStringToPrinter = False
        End If
    End Function
#End Region
End Class
