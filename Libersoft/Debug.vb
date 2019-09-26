Module Debug
#Region "variables"
    Public G_MostrarErrores As Boolean = False
    Public G_NombreEmpresa As String = "SIESM"
#End Region
#Region "FUNCIONES GENERALES"
    Public Sub X(ByVal ex As Exception)
        Const fic As String = "log.txt"

        If Not ex Is Nothing Then
            If G_MostrarErrores Then
                Msg(ex.StackTrace.ToString, 2)
            Else
                Try
                    Dim sw As New IO.StreamWriter(fic, True)
                    sw.WriteLine(vbCrLf + "------------------------------------------------------------------------------------------------------")
                    sw.WriteLine("/// ----> " + Now.ToString)
                    sw.WriteLine(ex.Source.ToString)
                    sw.WriteLine(ex.Message.ToString)
                    sw.WriteLine(ex.StackTrace.ToString)
                    sw.Close()
                    Console.WriteLine(ex.StackTrace.ToString)
                Catch exp As Exception
                End Try
            End If
        End If
    End Sub

    Public Sub Msg()
        Msg("<..>", 3)
    End Sub

    ''' <summary>  
    ''' Muestra mensaje en pantalla
    ''' </summary>
    ''' <param name="Mensaje">Mensaje a mostrar</param>
    ''' <param name="Icono">0-Info. 1-Warning. 2-Error. 3-Question</param>
    Public Sub Msg(ByVal Mensaje As String, Optional ByVal Icono As Integer = 0)
        Select Case Icono
            Case 0
                MsgBox(Mensaje, CType(vbOK + vbInformation, MsgBoxStyle), G_NombreEmpresa)
            Case 1
                MsgBox(Mensaje, CType(vbOK + vbExclamation, MsgBoxStyle), G_NombreEmpresa)
            Case 2
                MsgBox(Mensaje, CType(vbOK + vbCritical, MsgBoxStyle), G_NombreEmpresa)
            Case 3
                MsgBox(Mensaje, CType(vbOK + vbQuestion, MsgBoxStyle), G_NombreEmpresa)
        End Select
    End Sub
#End Region
End Module
