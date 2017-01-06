
Public Class ValidacionCaracteres
    Private _Permitidos As String   'Caracteres permitidos en la validación
    Private _Prohibidos As String   'Caracteres prohibidos en la validación
    Private _TipoValidacion As Byte 'Tipo de validación seleccionada


#Region "Propiedades"

    ''' <summary>
    ''' Establece o consulta la lista de caracteres permitidos
    ''' </summary>
    ''' <param name="Caracteres">Cadena de caracteres permitidos</param>
    ''' <value>Cadena sin separadores</value>
    ''' <returns>Cadena tipo Strings</returns>
    ''' <remarks></remarks>
    Public Property Permitidos(ByVal Caracteres As String)
        Get
            Return _Permitidos
        End Get
        Set(value)
            _Permitidos = Caracteres
        End Set
    End Property

    ''' <summary>
    ''' Establece o consulta la lista de caracteres prohibidos
    ''' </summary>
    ''' <param name="Caracteres">Cadena de caracteres prohibidos</param>
    ''' <value>Cadena sin separadores</value>
    ''' <returns>Cadena tipo Strings</returns>
    ''' <remarks></remarks>
    Public Property Prohibidos(ByVal Caracteres As String)
        Get
            Return _Prohibidos
        End Get
        Set(value)
            _Prohibidos = Caracteres
        End Set
    End Property

    ''' <summary>
    ''' Especifica que tipo de validación se aplicará
    ''' </summary>
    ''' <param name="Valor">Tipo de validación</param>
    ''' <value>Entero entre 1 y 4</value>
    ''' <returns>Tipo de validación Integer</returns>
    ''' <remarks></remarks>
    Public Property TipoValidacion(ByVal Valor As Byte)
        Get
            Return _TipoValidacion
        End Get
        Set(value)
            _TipoValidacion = Valor
        End Set
    End Property

#End Region

    ''' <summary>
    ''' Inicialización de la clase ValidacionCaracteres
    ''' </summary>
    ''' <param name="TipoValidacion">
    ''' Opción:
    ''' 1- Letras y números
    ''' 2- Solo letras
    ''' 3- Solo números
    ''' 4- Caracteres Permitidos
    ''' 5- Caracteres Prohibidos</param>
    ''' <param name="Permitidos">Cadena de caracteres permitidos (Opcional)</param>
    ''' <remarks><list type="numeric">
    ''' <item><description>Lestras y números</description></item>
    ''' <item><description>Solo Lestras</description></item>
    ''' <item><description>Solo números</description></item>
    ''' </list></remarks>
    Public Sub New(Optional ByVal TipoValidacion As Byte = 1, Optional ByVal Permitidos As String = "")
        _TipoValidacion = TipoValidacion
        _Permitidos = Permitidos
    End Sub

    ''' <summary>
    ''' Verificar
    ''' </summary>
    ''' <param name="Cadena">Caracter ó cadena a verificar</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Verificar(ByVal Cadena As String) As String

        Return ""
    End Function
End Class
