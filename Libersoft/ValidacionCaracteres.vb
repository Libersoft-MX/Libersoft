Public Class ValidacionCaracteres
    Private Especificos As String
    Private _TipoValidación As Integer

    ''' <summary>
    ''' Asignación del tipo de validación
    ''' </summary>
    ''' <remarks>Fundamental inicializar la clase con esta función antes de intentar validar</remarks>
    ''' <param name="Tipo"><para>0- Solo letras y números</para><para>1- Solo letras</para><para>2- Solo Números</para><para>3- Caracteres especificos</para></param>
    ''' <param name="Cadena">Especifique Todos los caracteres aceptados</param>
    Public Sub TipoValidacion(ByVal Tipo As Integer, Optional ByVal Cadena As String = "")
        _TipoValidación = Tipo
        If Tipo = 4 Then
            Especificos = Cadena
        End If
    End Sub

    Public Function Verificar(ByVal Cadena As String) As String

        Return ""
    End Function
End Class
