
Imports MySql.Data.MySqlClient
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Data
Imports System.Net.Mail
Imports System.Net

Module basico
    Public errorBD As String
    Public horaDesde As DateTime
    Public ultimaFalla
    Public autenticado As Boolean
    Public cadenaConexion As String
    Public be_log_activar As Boolean = False
    Public rutaBD As String = "robot"
    Public traduccion As String()
    Public be_idioma


    Sub Main(argumentos As String())

        'cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
        'notificarOP("cambio", 2864, 2865)

        If Process.GetProcessesByName _
          (Process.GetCurrentProcess.ProcessName).Length > 1 Then
        ElseIf argumentos.Length = 0 Then
            MsgBox("String connection missing", MsgBoxStyle.Critical, "SIGMA")
        Else
            cadenaConexion = argumentos(0)
            If argumentos.Length = 2 Then
                If argumentos(1) = "test" Then
                    pruebaMail()
                Else
                    Dim arreParametros = argumentos(1).Split(New Char() {";"c})
                    If arreParametros(0) = "P_EMERGENCIA" Or arreParametros(0) = "R_EMERGENCIA" Then
                        notificarParo(arreParametros(0))
                    ElseIf arreParametros(0) = "completada" Or arreParametros(0) = "cambio" Then
                        Dim uParametro = 0
                        If arreParametros.Length = 3 Then
                            uParametro = arreParametros(2)
                        End If
                        notificarOP(arreParametros(0), arreParametros(1), uParametro)
                    End If

                End If
            End If
        End If
        Application.Exit()
    End Sub
    Public Function consultaACT(cadena As String) As Integer
        Dim miConexion = New MySqlConnection

        miConexion.ConnectionString = cadenaConexion

        miConexion.Open()
        consultaACT = 0
        errorBD = ""
        If miConexion.State = ConnectionState.Open Then
            Try
                Dim comandoSQL As MySqlCommand = New MySqlCommand(cadena)
                comandoSQL.Connection = miConexion
                consultaACT = comandoSQL.ExecuteNonQuery()

            Catch ex As Exception
                errorBD = ex.Message
            End Try
        End If

        miConexion.Dispose()
        miConexion.Close()
        miConexion = Nothing
    End Function

    Public Function consultaSEL(cadena As String) As Data.DataSet

        Try
            errorBD = ""
            Dim miConexion = New MySqlConnection

            miConexion.ConnectionString = cadenaConexion

            miConexion.Open()

            If miConexion.State = ConnectionState.Open Then
                Try
                    Dim comandoSQL As MySqlCommand = New MySqlCommand(cadena)
                    comandoSQL.Connection = miConexion
                    Dim adapter As New MySqlDataAdapter(comandoSQL)
                    Dim LaData As New DataSet
                    adapter.Fill(LaData, "miData")

                    Return LaData
                Catch ex As Exception
                    errorBD = ex.Message
                End Try
            End If
            miConexion.Dispose()
            miConexion.Close()
            miConexion = Nothing
        Catch ex As Exception
            errorBD = ex.Message
        End Try

    End Function

    Function ValNull(ByVal ArVar As Object, ByVal arTipo As String) As Object
        Try
            'para columnas vacias sin datos
            If ArVar.Equals(System.DBNull.Value) Then
                Select Case arTipo
                    Case "A"
                        ValNull = ""
                    Case "N"
                        ValNull = 0
                    Case "D"
                        ValNull = 0
                    Case "F"
                        ValNull = CDate("00/00/0000")
                    Case "DT"
                        ValNull = New DateTime(1, 1, 1)
                    Case Else
                        ValNull = ""
                End Select
                Exit Function
            End If

            If Len(ArVar) > 0 Then
                Select Case arTipo
                    Case "A"
                        ValNull = ArVar
                    Case "N"
                        ValNull = Val(ArVar)
                    Case "D"
                        ValNull = CDec(ArVar)
                    Case "F"
                        If ArVar = "0" Then
                            ValNull = ""
                        Else
                            If InStr(ArVar, "/") > 0 Then
                                ValNull = ArVar
                            Else
                                ValNull = Format(ArVar, "dd/MM/yyyy")
                            End If
                        End If
                    Case Else
                        ValNull = ArVar
                End Select
            Else
                Select Case arTipo
                    Case "A"
                        ValNull = ""
                    Case "N"
                        ValNull = 0
                    Case "D"
                        ValNull = 0
                    Case "F"
                        ValNull = CDate("dd/MM/yyyy")
                    Case Else
                        ValNull = ""
                End Select
            End If
        Catch ex As Exception
            Select Case arTipo
                Case "A"
                    ValNull = ""
                Case "N"
                    ValNull = 0
                Case "D"
                    ValNull = 0
                Case "F"
                    ValNull = CDate("00000000")
                Case Else
                    ValNull = " "
            End Select
        End Try
    End Function

    'Function cadenaConexion() As String
    '   cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"
    'cadenaConexion = "server=10.241.241.30;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"

    'End Function

    Function calcularTiempo(Seg) As String
        calcularTiempo = ""
        If Seg < 60 Then
            calcularTiempo = Seg & traduccion(24)
        ElseIf Seg < 3600 Then
            calcularTiempo = Math.Round(Seg / 60, 1) & traduccion(25)
        Else
            calcularTiempo = Math.Round(Seg / 3600, 1) & traduccion(26)
        End If
    End Function

    Function calcularTiempoCad(Seg) As String
        calcularTiempoCad = "-"
        Dim horas = Math.Floor(Seg / 3600)
        Dim minutos = Math.Floor((Seg Mod 3600) / 60)
        Dim segundos = (Seg Mod 3600) Mod 60
        calcularTiempoCad = horas & ":" & Format(minutos, "00") & ":" & Format(segundos, "00")
    End Function

    Private Sub agregarLOG(cadena As String, Optional reporte As Integer = 0, Optional tipo As Integer = 0, Optional aplicacion As Integer = 40)
        If Not be_log_activar Then Exit Sub
        'tipo 0: Info
        'tipo 2: Advertencia
        'tipo 9: Error
        Dim regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".log (aplicacion, tipo, proceso, texto) VALUES (" & aplicacion & ", " & tipo & ", " & reporte & ", '" & Microsoft.VisualBasic.Strings.Left(cadena, 250) & "')")
    End Sub

    Sub etiquetas()
        Dim general = consultaSEL("SELECT cadena FROM " & rutaBD & ".det_idiomas_back WHERE idioma = " & IIf(be_idioma = 0, 1, be_idioma) & " AND modulo = 3 ORDER BY linea")
        Dim cadenaTrad = ""
        If general.Tables(0).Rows.Count > 0 Then
            For Each cadena In general.Tables(0).Rows
                cadenaTrad = cadenaTrad & cadena!cadena
            Next
        End If
        traduccion = cadenaTrad.Split(New Char() {";"c})
    End Sub
    Sub pruebaMail()
        Dim cadSQL As String = "SELECT idioma_defecto, correo_cuenta, correo_clave, correo_puerto, correo_ssl, correo_host FROM " & rutaBD & ".configuracion"
        Dim readerDS As DataSet = consultaSEL(cadSQL)
        Dim escape_mensaje = ""
        If readerDS.Tables(0).Rows.Count > 0 Then
            Dim reader As DataRow = readerDS.Tables(0).Rows(0)
            be_idioma = ValNull(reader!idioma_defecto, "N")
            etiquetas()
            agregarLOG(traduccion(28), 9, 0)
            Dim correo_cuenta As String = ValNull(reader!correo_cuenta, "A")
            Dim correo_clave As String = ValNull(reader!correo_clave, "A")
            Dim correo_puerto = ValNull(reader!correo_puerto, "A")
            Dim correo_ssl = ValNull(reader!correo_ssl, "A") = "S"
            Dim correo_host = ValNull(reader!correo_host, "A")
            Dim smtpServer As New SmtpClient()
            Try
                smtpServer.Credentials = New Net.NetworkCredential(correo_cuenta, correo_clave)
                smtpServer.Port = correo_puerto
                smtpServer.Host = correo_host
                smtpServer.EnableSsl = correo_ssl
                '
                Dim mail As New MailMessage
                mail.From = New MailAddress(correo_cuenta)
                mail.To.Add(correo_cuenta)
                traduccion(31) = Strings.Replace(traduccion(31), vbCrLf, "")
                traduccion(31) = Strings.Replace(traduccion(31), vbCr, "")
                traduccion(31) = Strings.Replace(traduccion(31), vbLf, "")

                mail.Subject = traduccion(31)
                mail.Body = traduccion(31)
                smtpServer.Send(mail)
                mail.Dispose()
                Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET correo_prueba = 'N', correo_respuesta = '" & Format(DateAndTime.Now, "yyyy-MMM-dd HH:mm:ss") & ": " & traduccion(28) & "'")
                agregarLOG(traduccion(28), 9, 0)
            Catch ex As Exception
                Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET correo_prueba = 'N', correo_respuesta = '" & Format(DateAndTime.Now, "yyyy-MMM-dd HH:mm:ss") & ": " & traduccion(29) & ex.Message & "'")
                agregarLOG(traduccion(29), 9, 0)
            End Try
        End If
    End Sub

    Sub notificarParo(mensaje As String)

        Dim cadSQL As String = "SELECT idioma_defecto, correo_cuenta, correo_clave, correo_puerto, correo_ssl, correo_host, oee_por_turno_cuentas_hxh, oee_por_turno_cuentas_hxh_turno, oee_por_turno_cuentas_hxh_dia, ruta_archivos_enviar FROM " & rutaBD & ".configuracion"
        Dim readerDS As DataSet = consultaSEL(cadSQL)
        Dim escape_mensaje = ""


        If readerDS.Tables(0).Rows.Count > 0 Then
            Dim reader As DataRow = readerDS.Tables(0).Rows(0)
            be_idioma = ValNull(reader!idioma_defecto, "N")
            etiquetas()
            Dim correo_cuenta As String = ValNull(reader!correo_cuenta, "A")
            Dim correo_clave As String = ValNull(reader!correo_clave, "A")
            Dim correo_puerto = ValNull(reader!correo_puerto, "A")
            Dim correo_ssl = ValNull(reader!correo_ssl, "A") = "S"
            Dim correo_host = ValNull(reader!correo_host, "A")
            Dim smtpServer As New SmtpClient()


            Try
                smtpServer.Credentials = New Net.NetworkCredential(correo_cuenta, correo_clave)
                smtpServer.Port = correo_puerto
                smtpServer.Host = correo_host
                smtpServer.EnableSsl = correo_ssl
                '
            Catch ex As Exception
                Application.Exit()
                Exit Sub
            End Try
            Dim correos As String()
            Dim correos_copia As String()
            Dim correos_oculta As String()
            Dim tempArray As String()
            Dim totalItems = 0
            cadSQL = "SELECT * FROM " & rutaBD & ".cat_correos WHERE estatus = 'A' AND tipo = " & If(mensaje = "P_EMERGENCIA", "2", "3")
            Dim mensajesDS As DataSet = consultaSEL(cadSQL)
            If mensajesDS.Tables(0).Rows.Count > 0 Then
                    For Each eMensaje In mensajesDS.Tables(0).Rows

                        Dim arreCanales = eMensaje!para.Split(New Char() {";"c})
                        For i = LBound(arreCanales) To UBound(arreCanales)
                            'Redimensionamos el Array temporal y preservamos el valor  
                            ReDim Preserve correos(totalItems + i)
                            correos(totalItems + i) = arreCanales(i)
                        Next
                        tempArray = correos
                        totalItems = correos.Length

                        Dim x As Integer, y As Integer
                        Dim z As Integer

                    For x = 0 To UBound(correos)
                        z = 0
                        For y = 0 To UBound(correos) - 1
                            'Si el elemento del array es igual al array temporal  
                            If correos(x) = tempArray(z) And y <> x Then
                                'Entonces Eliminamos el valor duplicado  
                                correos(y) = ""
                            End If
                            z = z + 1
                        Next y
                    Next x

                    totalItems = 0

                    arreCanales = eMensaje!copia.Split(New Char() {";"c})
                        For i = LBound(arreCanales) To UBound(arreCanales)
                            'Redimensionamos el Array temporal y preservamos el valor  
                            ReDim Preserve correos_copia(totalItems + i)
                            correos_copia(totalItems + i) = arreCanales(i)
                        Next
                        tempArray = correos_copia
                        totalItems = correos_copia.Length

                        x = 0
                        y = 0
                        z = 0

                    For x = 0 To UBound(correos_copia)
                        z = 0
                        For y = 0 To UBound(correos_copia) - 1
                            'Si el elemento del array es igual al array temporal  
                            If correos_copia(x) = tempArray(z) And y <> x Then
                                'Entonces Eliminamos el valor duplicado  
                                correos_copia(y) = ""
                            End If
                            z = z + 1
                        Next y
                    Next x

                    totalItems = 0

                    arreCanales = eMensaje!oculta.Split(New Char() {";"c})
                    For i = LBound(arreCanales) To UBound(arreCanales)
                            'Redimensionamos el Array temporal y preservamos el valor  
                            ReDim Preserve correos_oculta(totalItems + i)
                            correos_oculta(totalItems + i) = arreCanales(i)
                        Next
                        tempArray = correos_oculta
                        totalItems = correos_oculta.Length

                        x = 0
                        y = 0
                        z = 0

                        For x = 0 To UBound(correos_oculta)
                            z = 0
                            For y = 0 To UBound(correos_oculta) - 1
                                'Si el elemento del array es igual al array temporal  
                                If correos_oculta(x) = tempArray(z) And y <> x Then
                                    'Entonces Eliminamos el valor duplicado  
                                    correos_oculta(y) = ""
                                End If
                                z = z + 1
                            Next y
                        Next x

                        Dim mail As New MailMessage
                        Try
                            mail.From = New MailAddress(correo_cuenta)
                            For i = 0 To UBound(correos)
                                If correos(i).Length > 0 Then
                                    mail.To.Add(correos(i))
                                End If
                            Next i
                            For i = 0 To UBound(correos_copia)
                                If correos_copia(i).Length > 0 Then
                                    mail.CC.Add(correos_copia(i))
                                End If
                            Next i
                            For i = 0 To UBound(correos_oculta)
                                If correos_oculta(i).Length > 0 Then
                                    mail.Bcc.Add(correos_oculta(i))
                                End If
                            Next i
                        Dim cuerpo As String = ValNull(eMensaje!cuerpo, "A")
                        cuerpo = IIf(cuerpo.Length = 0, IIf(mensaje = "P_EMERGENCIA", "PARO DE EMERGENCIA", "RESOLUCIÓN DE PARO DE EMERGENCIA"), cuerpo)
                        mail.Body = cuerpo
                        cuerpo = ValNull(eMensaje!titulo, "A")
                        cuerpo = IIf(cuerpo.Length = 0, IIf(mensaje = "P_EMERGENCIA", "PARO DE EMERGENCIA", "RESOLUCIÓN DE PARO DE EMERGENCIA"), cuerpo)
                        mail.Subject = cuerpo
                        smtpServer.Send(mail)
                        agregarLOG(traduccion(32), 0, 0)

                    Catch ex As Exception
                            agregarLOG(ex.Message, 9, 0)
                        Finally
                            mail.Dispose()
                        End Try

                    Next

                End If

        End If
    End Sub


    Sub notificarOP(mensaje As String, entrante As Long, saliente As Long)

        Dim cadSQL As String = "SELECT idioma_defecto, correo_cuenta, correo_clave, correo_puerto, correo_ssl, correo_host, oee_por_turno_cuentas_hxh, oee_por_turno_cuentas_hxh_turno, oee_por_turno_cuentas_hxh_dia, ruta_archivos_enviar FROM " & rutaBD & ".configuracion"
        Dim readerDS As DataSet = consultaSEL(cadSQL)
        Dim escape_mensaje = ""


        If readerDS.Tables(0).Rows.Count > 0 Then
            Dim reader As DataRow = readerDS.Tables(0).Rows(0)
            be_idioma = ValNull(reader!idioma_defecto, "N")
            etiquetas()
            Dim correo_cuenta As String = ValNull(reader!correo_cuenta, "A")
            Dim correo_clave As String = ValNull(reader!correo_clave, "A")
            Dim correo_puerto = ValNull(reader!correo_puerto, "A")
            Dim correo_ssl = ValNull(reader!correo_ssl, "A") = "S"
            Dim correo_host = ValNull(reader!correo_host, "A")
            Dim smtpServer As New SmtpClient()


            Try
                smtpServer.Credentials = New Net.NetworkCredential(correo_cuenta, correo_clave)
                smtpServer.Port = correo_puerto
                smtpServer.Host = correo_host
                smtpServer.EnableSsl = correo_ssl
                '
            Catch ex As Exception
                Application.Exit()
                Exit Sub
            End Try
            Dim correos As String()
            Dim correos_copia As String()
            Dim correos_oculta As String()
            Dim tempArray As String()
            Dim totalItems = 0
            Dim datosOrden As String = ""
            cadSQL = "SELECT * FROM " & rutaBD & ".cat_correos WHERE estatus = 'A' AND tipo = " & If(mensaje = "completada", "5", "4")
            Dim mensajesDS As DataSet = consultaSEL(cadSQL)
            If mensajesDS.Tables(0).Rows.Count > 0 Then
                cadSQL = "SELECT a.*, b.numero, c.referencia, c.nombre AS nparte FROM " & rutaBD & ".equipos_objetivo a INNER JOIN " & rutaBD & ".lotes b ON a.lote = b.id INNER JOIN " & rutaBD & ".cat_partes c ON a.parte = c.id WHERE a.lote = " & entrante
                readerDS = consultaSEL(cadSQL)
                Dim sacos = 0
                Dim tarimas = 0
                If readerDS.Tables(0).Rows.Count > 0 Then
                    Dim tmp As Double
                    If readerDS.Tables(0).Rows(0)!kg_saco > 0 Then

                        tmp = readerDS.Tables(0).Rows(0)!van / readerDS.Tables(0).Rows(0)!kg_saco
                        If tmp - Math.Floor(tmp) > 0 Then
                            tmp = Math.Floor(tmp) + 1
                        End If
                        sacos = tmp
                    End If
                    If readerDS.Tables(0).Rows(0)!sacos_tarima > 0 Then
                        tmp = tmp / readerDS.Tables(0).Rows(0)!sacos_tarima
                        If tmp - Math.Floor(tmp) > 0 Then

                            tmp = Math.Floor(tmp) + 1
                        End If
                        tarimas = tmp
                    End If
                    If mensaje = "completada" Then
                        datosOrden = traduccion(35) & ": " & readerDS.Tables(0).Rows(0)!numero & Environment.NewLine
                        datosOrden = datosOrden & traduccion(36) & ": " & ValNull(readerDS.Tables(0).Rows(0)!nparte, "A") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(46) & ": " & ValNull(readerDS.Tables(0).Rows(0)!referencia, "A") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(37) & ": " & Format(readerDS.Tables(0).Rows(0)!objetivo, "###,###,###,##0.00") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(38) & ": " & Format(readerDS.Tables(0).Rows(0)!van, "###,###,###,##0.00") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(39) & ": " & Format(sacos, "###,###,###,##0") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(40) & ": " & Format(tarimas, "###,###,###,##0") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(45) & ": " & ValNull(readerDS.Tables(0).Rows(0)!notas, "A") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(41) & ": " & Format(Now(), "ddd, dd-MMM-yyyy HH:mm:ss") & Environment.NewLine
                    ElseIf mensaje = "cambio" Then
                        datosOrden = traduccion(42) & ": " & readerDS.Tables(0).Rows(0)!numero & Environment.NewLine
                        datosOrden = datosOrden & traduccion(36) & ": " & ValNull(readerDS.Tables(0).Rows(0)!nparte, "A") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(46) & ": " & ValNull(readerDS.Tables(0).Rows(0)!referencia, "A") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(37) & ": " & Format(readerDS.Tables(0).Rows(0)!objetivo, "###,###,###,##0.00") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(38) & ": " & Format(readerDS.Tables(0).Rows(0)!van, "###,###,###,##0.00") & Environment.NewLine
                        'datosOrden = datosOrden & traduccion(39) & ": " & Format(readerDS.Tables(0).Rows(0)!sacos, "###,###,###,##0") & Environment.NewLine
                        'datosOrden = datosOrden & traduccion(40) & ": " & Format(readerDS.Tables(0).Rows(0)!tarimas, "###,###,###,##0") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(39) & ": " & Format(sacos, "###,###,###,##0") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(40) & ": " & Format(tarimas, "###,###,###,##0") & Environment.NewLine
                        datosOrden = datosOrden & traduccion(45) & ": " & ValNull(readerDS.Tables(0).Rows(0)!notas, "A") & Environment.NewLine
                        cadSQL = "SELECT a.*, b.numero, c.nombre AS nparte FROM " & rutaBD & ".equipos_objetivo a INNER JOIN " & rutaBD & ".lotes b ON a.lote = b.id INNER JOIN " & rutaBD & ".cat_partes c ON a.parte = c.id WHERE a.lote = " & saliente
                        Dim readerDS2 As DataSet = consultaSEL(cadSQL)
                        If readerDS2.Tables(0).Rows.Count > 0 Then
                            datosOrden = datosOrden & Environment.NewLine & traduccion(43) & ": " & readerDS2.Tables(0).Rows(0)!numero & Environment.NewLine
                            datosOrden = datosOrden & traduccion(36) & ": " & ValNull(readerDS2.Tables(0).Rows(0)!nparte, "A") & Environment.NewLine
                            datosOrden = datosOrden & traduccion(46) & ": " & ValNull(readerDS2.Tables(0).Rows(0)!referencia, "A") & Environment.NewLine
                            datosOrden = datosOrden & traduccion(37) & ": " & Format(readerDS2.Tables(0).Rows(0)!objetivo, "###,###,###,##0.00") &
vbCrLf
                            datosOrden = datosOrden & traduccion(45) & ": " & ValNull(readerDS2.Tables(0).Rows(0)!notas, "A") & Environment.NewLine
                        End If
                        datosOrden = datosOrden & Environment.NewLine & traduccion(44) & ": " & Format(Now(), "ddd, dd-MMM-yyyy HH:mm:ss") & Environment.NewLine

                    End If
                End If

                For Each eMensaje In mensajesDS.Tables(0).Rows

                    Dim arreCanales = eMensaje!para.Split(New Char() {";"c})
                    For i = LBound(arreCanales) To UBound(arreCanales)
                        'Redimensionamos el Array temporal y preservamos el valor  
                        ReDim Preserve correos(totalItems + i)
                        correos(totalItems + i) = arreCanales(i)
                    Next
                    tempArray = correos
                    totalItems = correos.Length

                    Dim x As Integer, y As Integer
                    Dim z As Integer

                    For x = 0 To UBound(correos)
                        z = 0
                        For y = 0 To UBound(correos) - 1
                            'Si el elemento del array es igual al array temporal  
                            If correos(x) = tempArray(z) And y <> x Then
                                'Entonces Eliminamos el valor duplicado  
                                correos(y) = ""
                            End If
                            z = z + 1
                        Next y
                    Next x

                    totalItems = 0

                    arreCanales = eMensaje!copia.Split(New Char() {";"c})
                    For i = LBound(arreCanales) To UBound(arreCanales)
                        'Redimensionamos el Array temporal y preservamos el valor  
                        ReDim Preserve correos_copia(totalItems + i)
                        correos_copia(totalItems + i) = arreCanales(i)
                    Next
                    tempArray = correos_copia
                    totalItems = correos_copia.Length

                    x = 0
                    y = 0
                    z = 0

                    For x = 0 To UBound(correos_copia)
                        z = 0
                        For y = 0 To UBound(correos_copia) - 1
                            'Si el elemento del array es igual al array temporal  
                            If correos_copia(x) = tempArray(z) And y <> x Then
                                'Entonces Eliminamos el valor duplicado  
                                correos_copia(y) = ""
                            End If
                            z = z + 1
                        Next y
                    Next x

                    totalItems = 0

                    arreCanales = eMensaje!oculta.Split(New Char() {";"c})
                    For i = LBound(arreCanales) To UBound(arreCanales)
                        'Redimensionamos el Array temporal y preservamos el valor  
                        ReDim Preserve correos_oculta(totalItems + i)
                        correos_oculta(totalItems + i) = arreCanales(i)
                    Next
                    tempArray = correos_oculta
                    totalItems = correos_oculta.Length

                    x = 0
                    y = 0
                    z = 0

                    For x = 0 To UBound(correos_oculta)
                        z = 0
                        For y = 0 To UBound(correos_oculta) - 1
                            'Si el elemento del array es igual al array temporal  
                            If correos_oculta(x) = tempArray(z) And y <> x Then
                                'Entonces Eliminamos el valor duplicado  
                                correos_oculta(y) = ""
                            End If
                            z = z + 1
                        Next y
                    Next x

                    Dim mail As New MailMessage
                    Try
                        mail.From = New MailAddress(correo_cuenta)
                        For i = 0 To UBound(correos)
                            If correos(i).Length > 0 Then
                                mail.To.Add(correos(i))
                            End If
                        Next i
                        For i = 0 To UBound(correos_copia)
                            If correos_copia(i).Length > 0 Then
                                mail.CC.Add(correos_copia(i))
                            End If
                        Next i  
                        For i = 0 To UBound(correos_oculta)
                            If correos_oculta(i).Length > 0 Then
                                mail.Bcc.Add(correos_oculta(i))
                            End If
                        Next i
                        Dim cuerpo As String = ValNull(eMensaje!cuerpo, "A")
                        cuerpo = IIf(cuerpo.Length = 0, IIf(mensaje = "completada", "ORDEN DE PROCESO COMPLETADA", "CAMBIO DE ORDEN DE PROCESO"), cuerpo)
                        cuerpo = cuerpo & Environment.NewLine & Environment.NewLine & datosOrden
                        mail.Body = cuerpo
                        cuerpo = ValNull(eMensaje!titulo, "A")
                        cuerpo = IIf(cuerpo.Length = 0, IIf(mensaje = "completada", "ORDEN DE PROCESO COMPLETADA", "CAMBIO DE ORDEN DE PROCESO"), cuerpo)

                        mail.Subject = cuerpo
                        smtpServer.Send(mail)
                        agregarLOG(traduccion(32), 0, 0)

                    Catch ex As Exception
                        agregarLOG(ex.Message, 9, 0)
                    Finally
                        mail.Dispose()
                    End Try

                Next

            End If

        End If
    End Sub


End Module
