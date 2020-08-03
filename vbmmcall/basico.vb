﻿Imports MySql.Data.MySqlClient
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
    Public usuarioCerrar As String
    Public be_log_activar As Boolean = False
    Public cadenaConexion As String
    Public serialMmcall As String
    Public rutaBD As String = "sigma"


    Sub Main(argumentos As String())
        If Process.GetProcessesByName _
          (Process.GetCurrentProcess.ProcessName).Length > 1 Then
        ElseIf argumentos.Length = 0 Then
            MsgBox("No se puede iniciar el envío de mensajes a MMCall: Se requiere la cadena de conexión", MsgBoxStyle.Critical, "SIGMA Monitor")
        Else

            cadenaConexion = argumentos(0)
            'cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"
            Dim idProceso = Process.GetCurrentProcess.Id
            Dim mensajesDS As DataSet
            Dim registroDS As DataSet
            Dim eMensaje = ""
            Dim audiosGen = 0
            Dim audiosNGen = 0
            Dim respuestaWS As String = ""
            Dim mensajeGenerado As Boolean
            Dim canales As String = ""
            Dim laLinea As String = ""
            Dim laMaquina As String = ""
            Dim laArea As String = ""
            Dim laFalla As String = ""
            Dim fecha
            Dim tiempo As String = ""
            Dim nroReporte As Integer = 0
            Dim be_alarmas_mmcall As Boolean = False

            'Escalada 4
            Dim miError As String = ""
            Dim optimizar As Boolean = False
            Dim mantenerPrioridad As Boolean = False
            Dim validar_reloj As Boolean = False
            Dim regsAfectados = 0

            Dim maximo_largo_mmcall As Integer = 40

            Dim cadSQL2 As String = "SELECT CONCAT(key_number, serial) AS mmcall FROM mmcall.locations"
            Dim reader2 As DataSet = consultaSEL(cadSQL2)
            If reader2.Tables(0).Rows.Count > 0 Then
                serialMmcall = ValNull(reader2.Tables(0).Rows(0)!mmcall, "A")
            End If

            Dim cadSQL As String = "SELECT pagers_val, optimizar_mmcall, be_alarmas_mmcall, mantener_prioridad, maximo_largo_mmcall, be_log_activar FROM " & rutaBD & ".configuracion"
            Dim readerDS As DataSet = consultaSEL(cadSQL)
            If readerDS.Tables(0).Rows.Count > 0 Then
                Dim reader As DataRow = readerDS.Tables(0).Rows(0)
                optimizar = ValNull(reader!optimizar_mmcall, "A") = "S"
                mantenerPrioridad = ValNull(reader!mantener_prioridad, "A") = "S"
                maximo_largo_mmcall = ValNull(reader!maximo_largo_mmcall, "N")
                be_alarmas_mmcall = ValNull(reader!be_alarmas_mmcall, "A") = "S"
                be_log_activar = ValNull(reader!be_log_activar, "A") = "S"
                validar_reloj = ValNull(reader!pagers_val, "A") = "S"
            End If

            If be_alarmas_mmcall Then

                regsAfectados = consultaACT("UPDATE " & rutaBD & ".mensajes SET estatus = '" & idProceso & "' WHERE canal = 3 AND estatus = 'E'")
                Dim agrupado As Boolean = True
                If Not optimizar Then
                    cadSQL = "SELECT a.id, d.mmcall, a.prioridad, z.texto, z.titulo, 1 AS cuenta FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".mensajes_procesados z ON a.id = z.mensaje INNER JOIN " & rutaBD & ".cat_alertas b on a.alerta = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_distribucion d ON a.lista = d.id AND d.estatus = 'A' WHERE a.estatus = '" & idProceso & "' ORDER BY a.prioridad DESC, a.id"
                    agrupado = False
                ElseIf mantenerPrioridad Then
                    cadSQL = "SELECT a.lista, a.prioridad, b.mmcall, COUNT(*) AS cuenta, MAX(a.id) AS id FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".mensajes_procesados z ON a.id = z.mensaje INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c on a.alerta = c.id WHERE a.estatus = '" & idProceso & "' GROUP BY a.prioridad, a.lista, b.mmcall ORDER BY prioridad DESC"
                Else
                    cadSQL = "SELECT a.lista, b.mmcall, 0 AS prioridad, COUNT(*) AS cuenta, MAX(a.id) AS id FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".mensajes_procesados z ON a.id = z.mensaje INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c on a.alerta = c.id WHERE a.estatus = '" & idProceso & "' GROUP BY a.lista, b.mmcall"

                End If
                'Se preselecciona la voz
                mensajesDS = consultaSEL(cadSQL)
                Dim cadMMCALL = ""
                If mensajesDS.Tables(0).Rows.Count > 0 Then
                    For Each elmensaje In mensajesDS.Tables(0).Rows
                        canales = ValNull(elmensaje!mmcall, "A")
                        eMensaje = ""
                        If elmensaje!cuenta > 1 Then
                            eMensaje = "HAY " & elmensaje!cuenta & " MENSAJES POR ATENDER"
                            If mantenerPrioridad And elmensaje!prioridad > 0 Then
                                eMensaje = "HAY " & elmensaje!cuenta & " MENSAJES PRIORITARIOS"
                            End If
                        ElseIf agrupado Then
                            cadSQL = "SELECT * FROM " & rutaBD & ".mensajes_procesados WHERE mensaje = " & elmensaje!id
                            Dim mensajeAgrupado As DataSet = consultaSEL(cadSQL)
                            If mensajeAgrupado.Tables(0).Rows.Count > 0 Then
                                eMensaje = ValNull(mensajeAgrupado.Tables(0).Rows(0)!texto, "A")
                            End If
                        Else
                            eMensaje = ValNull(elmensaje!texto, "A")
                        End If
                        If eMensaje.Length > 0 And canales.Length > 0 Then
                            Dim mmcalls As String()
                            Dim tempArray As String()
                            Dim totalItems = 0
                            If canales.Length > 0 Then
                                Dim arreCanales = canales.Split(New Char() {";"c})
                                For i = LBound(arreCanales) To UBound(arreCanales)
                                    'Redimensionamos el Array temporal y preservamos el valor  
                                    ReDim Preserve mmcalls(totalItems + i)
                                    mmcalls(totalItems + i) = arreCanales(i)
                                Next
                                tempArray = mmcalls
                                totalItems = mmcalls.Length

                                Dim x As Integer, y As Integer
                                Dim z As Integer

                                For x = 0 To UBound(mmcalls)
                                    z = 0
                                    For y = 0 To UBound(mmcalls) - 1
                                        'Si el elemento del array es igual al array temporal  
                                        If mmcalls(x) = tempArray(z) And y <> x Then
                                            'Entonces Eliminamos el valor duplicado  
                                            mmcalls(y) = ""
                                        End If
                                        z = z + 1
                                    Next y
                                Next x
                                canales = ""
                            End If
                            If maximo_largo_mmcall = 0 Then maximo_largo_mmcall = 40
                            eMensaje = Microsoft.VisualBasic.Strings.Left(eMensaje, maximo_largo_mmcall).Trim
                            mensajeGenerado = False

                            For i = 0 To UBound(mmcalls)
                                If mmcalls(i).Length > 0 Then
                                    'Se valida el reloj
                                    Dim posReloj = Strings.InStr(mmcalls(i), "number=")
                                    Dim iReloj = Strings.Mid(mmcalls(i), posReloj + 7)
                                    If iReloj.Length = 0 Then
                                        If IsNumeric(mmcalls(i)) Then
                                            iReloj = mmcalls(i)
                                        Else
                                            iReloj = "-1"
                                        End If
                                    End If
                                    If iReloj <> "-1" Then
                                        If Not valReloj(iReloj) And validar_reloj Then
                                            agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: ID sin licencia", nroReporte, 9)
                                            Continue For
                                        Else
                                            If cadMMCALL = "" Then
                                                cadMMCALL = "INSERT INTO mmcall.tasks (location_id, task, message, recipients, status, created) VALUES "
                                            Else
                                                cadMMCALL = cadMMCALL + ","
                                            End If
                                            cadMMCALL = cadMMCALL & "(1, 'page', '" & eMensaje & "', '" & (iReloj + 100) & "', 0, NOW())"

                                        End If
                                    End If
                                    'respuestaWS = ""

                                    'Try
                                    'If validarURI(mmcalls(i) & "&message=" & eMensaje) Then
                                    'Dim fr As System.Net.HttpWebRequest
                                    'Dim targetURI As New Uri(mmcalls(i) & "&message=" & eMensaje)
                                    '
                                    'fr = DirectCast(HttpWebRequest.Create(targetURI), System.Net.HttpWebRequest)
                                    'If (fr.GetResponse().ContentLength > 0) Then
                                    'Dim str As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
                                    'respuestaWS = str.ReadToEnd
                                    'Str.Close()
                                    'End If
                                    'mensajeGenerado = respuestaWS = "success" And Not mensajeGenerado
                                    'If mensajeGenerado Then
                                    'audiosGen = audiosGen + 1
                                    'Else

                                    'audiosNGen = audiosNGen + 1
                                    'agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: " & respuestaWS, nroReporte, 9)
                                    'End If
                                    'Else
                                    'audiosNGen = audiosNGen + 1
                                    'agregarLOG("" & mmcalls(i) & "&message=" & eMensaje & " error: la dirección no es válida", nroReporte, 9)
                                    'End If
                                    'Catch ex As System.Net.WebException
                                    'agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: " & ex.Message,  , 9)
                                    'audiosNGen = audiosNGen + 1
                                    'miError = ex.Message
                                    'End Try
                                End If
                            Next
                        End If
                        If cadMMCALL.Length > 0 Then cadMMCALL = cadMMCALL & ";"
                        'If mensajeGenerado Then
                        cadSQL = cadMMCALL & "UPDATE " & rutaBD & ".mensajes SET estatus = 'Z', enviada = NOW()  WHERE id = " & elmensaje!id
                        If optimizar Then
                            If elmensaje!cuenta > 1 Then
                                cadSQL = "UPDATE " & rutaBD & ".mensajes SET estatus = 'Z', enviada = NOW() WHERE canal = 3 AND lista = " & elmensaje!lista & " AND estatus = '" & idProceso & "'"
                                If mantenerPrioridad Then
                                    cadSQL = cadSQL & " AND prioridad = " & elmensaje!prioridad
                                End If
                            End If
                        End If
                        regsAfectados = consultaACT(cadSQL)
                        cadMMCALL = ""
                        'End If
                    Next
                    If audiosGen > 0 Or audiosNGen > 0 Then
                        agregarLOG("Se generaron " & audiosGen & " mensaje(s) a MMCall y no se generaron " & audiosNGen & " mensaje(s) a MMCall ")
                    End If
                End If
            End If
        End If
        Application.Exit()
    End Sub
    Private Sub agregarLOG(cadena As String, Optional reporte As Integer = 0, Optional tipo As Integer = 0, Optional aplicacion As Integer = 20)
        If Not be_log_activar Then Exit Sub
        'tipo 0: Info
        'tipo 2: Advertencia
        'tipo 9: Error
        Dim regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".log (aplicacion, tipo, proceso, texto) VALUES (" & aplicacion & ", " & tipo & ", " & reporte & ", '" & Microsoft.VisualBasic.Strings.Left(cadena, 250) & "')")
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
    'cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"
    'cadenaConexion = "server=10.241.241.30;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"

    'End Function

    Function calcularTiempo(Seg) As String
        calcularTiempo = ""
        If Seg < 60 Then
            calcularTiempo = Seg & " seg"
        ElseIf Seg < 3600 Then
            calcularTiempo = Math.Round(Seg / 60, 1) & " min"
        Else
            calcularTiempo = Math.Round(Seg / 3600, 1) & " hr"
        End If
    End Function

    Function calcularTiempoCad(Seg) As String
        calcularTiempoCad = "-"
        Dim horas = Math.Floor(Seg / 3600)
        Dim minutos = Math.Floor((Seg Mod 3600) / 60)
        Dim segundos = (Seg Mod 3600) Mod 60
        calcularTiempoCad = horas & ":" & Format(minutos, "00") & ":" & Format(segundos, "00")
    End Function

    Function validarURI(ByVal cadena As String) As Boolean
        Dim validatedUri As System.Uri
        Return Uri.TryCreate(cadena, UriKind.RelativeOrAbsolute, validatedUri)
    End Function

    Function valReloj(id As String) As Boolean
        valReloj = False
        Dim tipo As String = "R"
        Dim claveInterna = "CronoEIntelraciVn201i"

        If id = "" Then
            Exit Function
        End If
        Dim cadSQL As String = "SELECT tipo, mmcall, cronos FROM " & rutaBD & ".licencias WHERE tipo = '" & tipo & "' AND mmcall = '" & id & "'"
        Dim reader As DataSet = consultaSEL(cadSQL)
        If reader.Tables(0).Rows.Count > 0 Then

            Dim temporal = ""
            Dim clavePublica = ""
            Dim buscarEn
            Dim numeroActual = 0
            Dim temporal2 = claveInterna
            Dim numero1
            Dim numero2

            Dim num1 = 0
            Dim num2 = 0
            Dim num3 = 0
            Dim posicion

            If claveInterna.Length > id.Length Then
                temporal = ""
                Dim recorrido = 0
                Do While temporal.Length < claveInterna.Length
                    If recorrido >= serialMmcall.Length Then
                        recorrido = 0
                    End If
                    numero1 = Asc(id(recorrido Mod id.Length))
                    numero2 = serialMmcall(recorrido)
                    If numero1.ToString.Length = 1 Then
                        buscarEn = numero1
                        numeroActual = 0
                        num1 = 0
                        num2 = 0
                        num3 = 0
                    ElseIf numero1.ToString.Length = 2 Then

                        num1 = Val(Strings.Mid(numero1.ToString, 1, 1))
                        num2 = Val(Strings.Mid(numero1.ToString, 2))
                        If numeroActual = 0 Then
                            buscarEn = num1
                            numeroActual = 1
                        Else

                            buscarEn = num2
                            numeroActual = 0
                        End If
                        posicion = num1 + num2 + recorrido
                    ElseIf numero1.ToString.Length = 3 Then

                        num1 = Val(Strings.Mid(numero1.ToString, 1, 1))
                        num2 = Val(Strings.Mid(numero1.ToString, 2, 1))
                        num3 = Val(Strings.Mid(numero1.ToString, 3))
                        If numeroActual = 0 Then
                            buscarEn = num1
                            numeroActual = 1
                        ElseIf numeroActual = 1 Then
                            buscarEn = num2
                            numeroActual = 2
                        Else
                            buscarEn = num3
                            numeroActual = 0
                        End If
                        posicion = num1 + num2 + num3 + recorrido
                    End If
                    posicion = posicion + Val(numero2)
                    If posicion > serialMmcall.Length - 1 Then
                        posicion = posicion Mod serialMmcall.Length
                    End If
                    temporal = IIf(temporal.Length = 0, reader.Tables(0).Rows(0)!tipo, "") + temporal + serialMmcall(posicion)
                    recorrido = recorrido + 1
                Loop
                temporal = Strings.Mid(temporal, 1, claveInterna.Length)
            ElseIf claveInterna.Length = id.Length Then
                temporal = id
            ElseIf id.Length > claveInterna.Length Then
                temporal = id
                temporal2 = claveInterna
                Do While temporal2.Length < id.Length
                    temporal2 = temporal2 & claveInterna
                Loop
                temporal2 = Strings.Mid(temporal2, 1, id.Length)
            End If
            Dim cadComparar = ""
            For i = 0 To temporal.Length - 1
                Dim numero As String = Asc(temporal(i)) Xor Asc(temporal2(i))
                If (numero.Length = 1) Then
                    cadComparar = numero
                ElseIf (numero.Length = 2) Then
                    cadComparar = Strings.Mid(numero, 2, 1)
                ElseIf (numero.Length = 3) Then
                    cadComparar = Strings.Mid(numero, 2, 1)
                End If
                clavePublica = clavePublica + cadComparar
            Next
            valReloj = True
            For i = 0 To temporal.Length - 1
                Dim numero As String = Asc(clavePublica(i)) Xor Asc(claveInterna(i))
                If (numero.Length = 1) Then
                    cadComparar = numero
                ElseIf (numero.Length = 2) Then
                    cadComparar = Strings.Mid(numero, 2, 1)
                ElseIf (numero.Length = 3) Then
                    cadComparar = Strings.Mid(numero, 2, 1)
                End If
                If cadComparar <> Strings.Mid(reader.Tables(0).Rows(0)!cronos, i + 1, 1) Then
                    valReloj = False
                    Exit Function
                End If
            Next
        End If
    End Function



    Sub MainAntes(argumentos As String())
        If Process.GetProcessesByName _
          (Process.GetCurrentProcess.ProcessName).Length > 1 Then
        ElseIf argumentos.Length = 0 Then
            MsgBox("No se puede iniciar el envío de mensajes a MMCall: Se requiere la cadena de conexión", MsgBoxStyle.Critical, "SIGMA Monitor")
        Else

            cadenaConexion = argumentos(0)
            'cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"
            Dim idProceso = Process.GetCurrentProcess.Id
            Dim mensajesDS As DataSet
            Dim registroDS As DataSet
            Dim eMensaje = ""
            Dim audiosGen = 0
            Dim audiosNGen = 0
            Dim respuestaWS As String = ""
            Dim mensajeGenerado As Boolean
            Dim canales As String = ""
            Dim laLinea As String = ""
            Dim laMaquina As String = ""
            Dim laArea As String = ""
            Dim laFalla As String = ""
            Dim fecha
            Dim tiempo As String = ""
            Dim nroReporte As Integer = 0
            Dim be_alarmas_mmcall As Boolean = False

            'Escalada 4
            Dim miError As String = ""
            Dim optimizar As Boolean = False
            Dim mantenerPrioridad As Boolean = False
            Dim regsAfectados = 0

            Dim maximo_largo_mmcall As Integer = 40

            Dim cadSQL2 As String = "SELECT CONCAT(key_number, serial) AS mmcall FROM mmcall.locations"
            Dim reader2 As DataSet = consultaSEL(cadSQL2)
            If reader2.Tables(0).Rows.Count > 0 Then
                serialMmcall = ValNull(reader2.Tables(0).Rows(0)!mmcall, "A")
            End If

            Dim cadSQL As String = "SELECT optimizar_mmcall, be_alarmas_mmcall, mantener_prioridad, maximo_largo_mmcall, be_log_activar FROM " & rutaBD & ".configuracion"
            Dim readerDS As DataSet = consultaSEL(cadSQL)
            If readerDS.Tables(0).Rows.Count > 0 Then
                Dim reader As DataRow = readerDS.Tables(0).Rows(0)
                optimizar = ValNull(reader!optimizar_mmcall, "A") = "S"
                mantenerPrioridad = ValNull(reader!mantener_prioridad, "A") = "S"
                maximo_largo_mmcall = ValNull(reader!maximo_largo_mmcall, "N")
                be_alarmas_mmcall = ValNull(reader!be_alarmas_mmcall, "A") = "S"
                be_log_activar = ValNull(reader!be_log_activar, "A") = "S"
            End If

            If be_alarmas_mmcall Then

                regsAfectados = consultaACT("UPDATE " & rutaBD & ".mensajes SET estatus = '" & idProceso & "' WHERE canal = 3 AND estatus = 'A'")

                If Not optimizar Then
                    cadSQL = "SELECT a.id, 1 AS cuenta, b.evento FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_alertas b on a.alerta = b.id WHERE a.canal = 3 AND a.estatus = '" & idProceso & "' ORDER BY a.prioridad DESC, a.id"
                ElseIf mantenerPrioridad Then
                    cadSQL = "SELECT a.prioridad, a.lista, c.evento, b.mmcall, COUNT(*) AS cuenta, MAX(a.id) AS id FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c on a.alerta = c.id WHERE a.canal = 3 AND a.estatus = '" & idProceso & "' GROUP BY a.prioridad, a.lista, c.evento, b.mmcall ORDER BY prioridad DESC"
                Else
                    cadSQL = "SELECT a.lista, c.evento, b.mmcall, 0 AS prioridad, COUNT(*) AS cuenta, MAX(a.id) AS id FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c on a.alerta = c.id WHERE a.canal = 3 AND a.estatus = '" & idProceso & "' GROUP BY a.lista, c.evento, b.mmcall, 4"
                End If
                'Se preselecciona la voz
                mensajesDS = consultaSEL(cadSQL)
                Dim generarMensaje As Boolean
                If mensajesDS.Tables(0).Rows.Count > 0 Then
                    For Each elmensaje In mensajesDS.Tables(0).Rows
                        generarMensaje = False
                        eMensaje = ""
                        If elmensaje!cuenta > 1 Then
                            eMensaje = "HAY " & elmensaje!cuenta & " MENSAJES POR ATENDER"
                            If mantenerPrioridad And elmensaje!prioridad > 0 Then
                                eMensaje = "HAY " & elmensaje!cuenta & " MENSAJES PRIORITARIOS"
                            End If
                            canales = ValNull(elmensaje!mmcall, "A")
                            laLinea = ""
                            laMaquina = ""
                            laArea = ""
                            laFalla = ""
                            tiempo = ""
                            generarMensaje = True
                            nroReporte = 0
                        Else

                            If elmensaje!evento < 200 Then
                                cadSQL = "SELECT a.*, 0 AS rate, 0 AS oee, b.mmcall, e.nombre as nlinea, f.nombre as nmaquina, g.nombre as narea, h.nombre as nfalla, c.id AS idalerta, c.evento AS tipoalerta, c.acumular, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.fecha, d.inicio_atencion, d.inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".reportes d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".cat_lineas e ON d.linea = e.id LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.maquina = f.id LEFT JOIN " & rutaBD & ".cat_areas g ON d.area = g.id LEFT JOIN " & rutaBD & ".cat_fallas h ON d.falla = h.id WHERE a.id = " & elmensaje!id
                            ElseIf elmensaje!evento < 300 Then
                                cadSQL = "SELECT a.*, IF(d.rate_teorico > 0, d.rate / d.rate_teorico * 100, 0) AS rate, d.oee, b.mmcall, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.evento AS tipoalerta, c.acumular, c.mensaje_mmcal, d.rate_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.parada_desde AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                            ElseIf elmensaje!evento = 301 Then
                                cadSQL = "SELECT a.*, e1.referencia, e1.nombre AS producto, b1.numero AS nlote, IFNULL((SELECT MIN(orden) FROM " & rutaBD & ".prioridades WHERE parte = b1.parte AND fecha >= NOW() AND estatus = 'A'), 100) AS prioridad, d.ruta_secuencia, d.ruta_secuencia_antes, IFNULL(c1.nombre, 'N/A') AS ruta_antes, IFNULL(d1.nombre, 'N/A') AS ruta_despues, b.mmcall, c.id AS idalerta, c.evento AS tipoalerta, c.acumular, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS fecha, NOW() AS inicio_atencion, NOW() AS inicio_reporte, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".lotes_historia d ON a.proceso = d.id INNER JOIN " & rutaBD & ".lotes b1 ON d.lote = b1.id LEFT JOIN " & rutaBD & ".det_rutas c1 ON d.ruta_detalle_anterior = c1.id LEFT JOIN " & rutaBD & ".det_rutas d1 ON d.ruta_detalle = d1.id LEFT JOIN " & rutaBD & ".cat_partes e1 ON b1.parte = e1.id WHERE a.id = " & elmensaje!id
                            ElseIf elmensaje!evento = 302 Or elmensaje!evento = 303 Or elmensaje!evento = 305 Or elmensaje!evento = 306 Then
                                cadSQL = "SELECT a.*, d.hasta, d.numero AS nlote, d.fecha, TIME_TO_SEC(TIMEDIFF(d.hasta, NOW())) AS previo, d.ruta_secuencia, c1.referencia, c1.nombre AS producto, IFNULL(b1.nombre, 'N/A') AS ruta_actual, IFNULL(e1.nombre, 'N/A') as equipo, IFNULL(d1.nombre, 'N/A') as nproceso, b.mmcall, c.id AS idalerta, c.evento AS tipoalerta, c.acumular, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS inicio_atencion, NOW() AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".lotes d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".det_rutas b1 ON d.ruta_detalle = b1.id LEFT JOIN " & rutaBD & ".cat_partes c1 ON d.parte = c1.id LEFT JOIN " & rutaBD & ".cat_procesos d1 ON d.proceso = d1.id LEFT JOIN " & rutaBD & ".cat_maquinas e1 ON d.equipo= e1.id WHERE a.id = " & elmensaje!id
                            ElseIf elmensaje!evento = 304 Or elmensaje!evento = 307 Then
                                cadSQL = "SELECT a.*, 0 AS previo, d.carga, d.alarma, d.alarma_rep, d.fecha, d.permitir_reprogramacion, d.equipo, d.fecha, IFNULL(b1.nombre, 'N/A') as nequipo, IFNULL(c1.nombre, 'N/A') as nproceso, IFNULL((SELECT SUM(cantidad) FROM " & rutaBD & ".programacion WHERE carga = d.id AND estatus = 'A'), 0) AS piezas, (SELECT COUNT(*) FROM " & rutaBD & ".lotes WHERE carga = d.id) AS avance, b.mmcall, c.id AS idalerta, c.evento AS tipoalerta, c.acumular, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS inicio_atencion, NOW() AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".cargas d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".cat_maquinas b1 ON d.equipo = b1.id AND b1.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_procesos c1 ON b1.proceso = c1.id AND c1.estatus = 'A' WHERE a.id = " & elmensaje!id
                            End If

                            registroDS = consultaSEL(cadSQL)
                            If registroDS.Tables(0).Rows.Count > 0 Then
                                nroReporte = registroDS.Tables(0).Rows(0)!proceso
                                canales = ValNull(registroDS.Tables(0).Rows(0)!mmcall, "A")
                                If canales.Length > 0 Then
                                    generarMensaje = True


                                    If registroDS.Tables(0).Rows(0)!tipoalerta < 300 Then
                                        laLinea = ValNull(registroDS.Tables(0).Rows(0)!nlinea, "A")
                                        laMaquina = ValNull(registroDS.Tables(0).Rows(0)!nmaquina, "A")
                                        laArea = ValNull(registroDS.Tables(0).Rows(0)!narea, "A")
                                        laFalla = ValNull(registroDS.Tables(0).Rows(0)!nfalla, "A")
                                    Else
                                        laLinea = ""
                                        laMaquina = ""
                                        laArea = ""
                                        laFalla = ""
                                    End If


                                    If registroDS.Tables(0).Rows(0)!tipoalerta = 101 Or registroDS.Tables(0).Rows(0)!tipoalerta = 201 Or registroDS.Tables(0).Rows(0)!tipoalerta = 301 Then
                                        fecha = registroDS.Tables(0).Rows(0)!fecha
                                    ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 102 Or registroDS.Tables(0).Rows(0)!tipoalerta = 202 Or registroDS.Tables(0).Rows(0)!tipoalerta > 300 Then
                                        fecha = registroDS.Tables(0).Rows(0)!inicio_atencion
                                    ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 103 Or registroDS.Tables(0).Rows(0)!tipoalerta = 203 Then
                                        fecha = registroDS.Tables(0).Rows(0)!inicio_reporte
                                    End If
                                    tiempo = calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, fecha, Now))
                                    If registroDS.Tables(0).Rows(0)!tipo = 0 Then
                                        eMensaje = ValNull(registroDS.Tables(0).Rows(0)!mensaje_mmcall, "A")
                                    ElseIf registroDS.Tables(0).Rows(0)!tipo = 8 Then
                                        eMensaje = ValNull(registroDS.Tables(0).Rows(0)!resolucion_mensaje, "A")
                                    ElseIf registroDS.Tables(0).Rows(0)!tipo = 7 Then
                                        eMensaje = ValNull(registroDS.Tables(0).Rows(0)!cancelacion_mensaje, "A")
                                    End If

                                    If eMensaje.Length > 0 Then
                                        eMensaje = Replace(eMensaje, "[0]", nroReporte)
                                        eMensaje = Replace(eMensaje, "[1]", laLinea)
                                        eMensaje = Replace(eMensaje, "[2]", laMaquina)
                                        eMensaje = Replace(eMensaje, "[3]", laArea)
                                        eMensaje = Replace(eMensaje, "[4]", laFalla)
                                        eMensaje = Replace(eMensaje, "[5]", Format(fecha, "dd/MM HH:mm"))
                                        eMensaje = Replace(eMensaje, "[11]", tiempo)
                                        If registroDS.Tables(0).Rows(0)!tipoalerta < 300 Then
                                            eMensaje = Replace(eMensaje, "[12]", Format(registroDS.Tables(0).Rows(0)!rate, "0.0"))
                                            eMensaje = Replace(eMensaje, "[13]", Format(registroDS.Tables(0).Rows(0)!oee, "0.0"))

                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[20]", "R" & ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[20]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!fase - 10, "N") > 0 Then
                                            Dim escala = ValNull(registroDS.Tables(0).Rows(0)!fase, "N") - 10
                                            eMensaje = Replace(eMensaje, "[30]", "Escalado al Nivel " & If(escala > 0, escala, 0))
                                        Else
                                            eMensaje = Replace(eMensaje, "[30]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!escalamientos1, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[31]", "R" & ValNull(registroDS.Tables(0).Rows(0)!escalamientos1, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[31]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!escalamientos2, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[32]", "R" & ValNull(registroDS.Tables(0).Rows(0)!escalamientos2, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[32]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!escalamientos3, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[33]", "R" & ValNull(registroDS.Tables(0).Rows(0)!escalamientos3, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[33]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!escalamientos4, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[34]", "R" & ValNull(registroDS.Tables(0).Rows(0)!escalamientos4, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[34]", "")
                                        End If
                                        If ValNull(registroDS.Tables(0).Rows(0)!escalamientos5, "N") > 0 Then
                                            eMensaje = Replace(eMensaje, "[35]", "R" & ValNull(registroDS.Tables(0).Rows(0)!escalamientos5, "N"))
                                        Else
                                            eMensaje = Replace(eMensaje, "[35]", "")
                                        End If


                                        If elmensaje!evento = 301 Then
                                            eMensaje = Replace(eMensaje, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                            eMensaje = Replace(eMensaje, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                            eMensaje = Replace(eMensaje, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                            eMensaje = Replace(eMensaje, "[70]", ValNull(registroDS.Tables(0).Rows(0)!ruta_antes, "A"))
                                            eMensaje = Replace(eMensaje, "[71]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia_antes, "A"))
                                            eMensaje = Replace(eMensaje, "[72]", ValNull(registroDS.Tables(0).Rows(0)!ruta_despues, "A"))
                                            eMensaje = Replace(eMensaje, "[73]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))

                                        ElseIf elmensaje!evento = 302 Or elmensaje!evento = 305 Then
                                            eMensaje = Replace(eMensaje, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                            eMensaje = Replace(eMensaje, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                            eMensaje = Replace(eMensaje, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                            eMensaje = Replace(eMensaje, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                            eMensaje = Replace(eMensaje, "[44]", ValNull(registroDS.Tables(0).Rows(0)!ruta_actual, "A"))
                                            eMensaje = Replace(eMensaje, "[45]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))
                                            eMensaje = Replace(eMensaje, "[50]", Format(registroDS.Tables(0).Rows(0)!fecha, "dd/MMM/yyyy HH:mm:ss"))
                                            eMensaje = Replace(eMensaje, "[51]", Format(registroDS.Tables(0).Rows(0)!hasta, "dd/MMM/yyyy HH:mm:ss"))
                                            eMensaje = Replace(eMensaje, "[52]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!hasta, DateAndTime.Now)))
                                            eMensaje = Replace(eMensaje, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                                        ElseIf elmensaje!evento = 303 Or elmensaje!evento = 306 Then
                                            eMensaje = Replace(eMensaje, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                            eMensaje = Replace(eMensaje, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                            eMensaje = Replace(eMensaje, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                            eMensaje = Replace(eMensaje, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                            eMensaje = Replace(eMensaje, "[44]", ValNull(registroDS.Tables(0).Rows(0)!ruta_actual, "A"))
                                            eMensaje = Replace(eMensaje, "[45]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))
                                            eMensaje = Replace(eMensaje, "[61]", ValNull(registroDS.Tables(0).Rows(0)!equipo, "A"))
                                            eMensaje = Replace(eMensaje, "[62]", Format(registroDS.Tables(0).Rows(0)!fecha, "dd/MMM/yyyy HH:mm:ss"))
                                            eMensaje = Replace(eMensaje, "[63]", Format(registroDS.Tables(0).Rows(0)!hasta, "dd/MMM/yyyy HH:mm:ss"))
                                            eMensaje = Replace(eMensaje, "[64]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!hasta, DateAndTime.Now)))
                                            eMensaje = Replace(eMensaje, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                                        ElseIf elmensaje!evento = 304 Or elmensaje!evento = 307 Then
                                            eMensaje = Replace(eMensaje, "[80]", ValNull(registroDS.Tables(0).Rows(0)!carga, "A"))
                                            eMensaje = Replace(eMensaje, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                            eMensaje = Replace(eMensaje, "[61]", ValNull(registroDS.Tables(0).Rows(0)!nequipo, "A"))
                                            eMensaje = Replace(eMensaje, "[81]", Format(registroDS.Tables(0).Rows(0)!fecha, "dd/MMM/yyyy HH:mm:ss"))
                                            eMensaje = Replace(eMensaje, "[82]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!fecha, DateAndTime.Now)))
                                            eMensaje = Replace(eMensaje, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                                            eMensaje = Replace(eMensaje, "[84]", ValNull(registroDS.Tables(0).Rows(0)!texto, "A"))


                                        End If



                                        Dim antes As String = "ÃÀÁÄÂÈÉËÊÌÍÏÎÒÓÖÔÙÚÜÛãàáäâèéëêìíïîòóöôùúüûÑñÇç"
                                        Dim ahora As String = "AAAAAEEEEIIIIOOOOUUUUaaaaaeeeeiiiioooouuuunncc"
                                        For i = 0 To antes.Length - 1
                                            eMensaje = Replace(eMensaje, antes(i), ahora(i))
                                        Next
                                        eMensaje = Replace(eMensaje, ";", " ")
                                        eMensaje = Replace(eMensaje, "\", "-")
                                        eMensaje = Replace(eMensaje, "/", "-")
                                        eMensaje = Replace(eMensaje, "[90]", "")
                                        eMensaje = Replace(eMensaje, System.Environment.NewLine, " ")


                                        'Se cambian los caracteres especiales

                                    Else
                                        If registroDS.Tables(0).Rows(0)!tipoalerta = 101 Then
                                            eMensaje = "REPORTE " & nroReporte & " TIEMPO ESPERA EXCED"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 102 Then
                                            eMensaje = "REPORTE " & nroReporte & " TIEMPO REPARAC EXCED"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 103 Then
                                            eMensaje = "REPORTE " & nroReporte & " TIEMPO INFORME EXCED"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 201 Then
                                            eMensaje = "BAJO RATE EN " & laMaquina
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 202 Then
                                            eMensaje = "SOBRE RATE EN " & laMaquina
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 203 Then
                                            eMensaje = laMaquina & " NO SE DETECTAN PIEZAS"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 301 Then
                                            eMensaje = "SALTO DE OPERACION"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 302 Then
                                            eMensaje = "TIEMPO DE STOCK VENCIDO"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 303 Then
                                            eMensaje = "TIEMPO DE PROCESO VENCIDO"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 304 Then
                                            eMensaje = "TIEMPO DE ENTREGA VENCIDO"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 305 Then
                                            eMensaje = "TIEMPO DE STOCK POR VENCER"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 306 Then
                                            eMensaje = "TIEMPO DE PROCESO POR VENCER"
                                        ElseIf registroDS.Tables(0).Rows(0)!tipoalerta = 307 Then
                                            eMensaje = "TIEMPO DE ENTREGA POR VENCER"
                                        End If
                                        agregarLOG("La alerta " & registroDS.Tables(0).Rows(0)!idalerta & " no tiene un mensaje de MMCall definido se tomó el mensaje por defecto", nroReporte, 2)

                                    End If

                                End If
                            End If
                        End If
                        If generarMensaje And canales.Length > 0 Then
                            Dim mmcalls As String()
                            Dim tempArray As String()
                            Dim totalItems = 0
                            If canales.Length > 0 Then
                                Dim arreCanales = canales.Split(New Char() {";"c})
                                For i = LBound(arreCanales) To UBound(arreCanales)
                                    'Redimensionamos el Array temporal y preservamos el valor  
                                    ReDim Preserve mmcalls(totalItems + i)
                                    mmcalls(totalItems + i) = arreCanales(i)
                                Next
                                tempArray = mmcalls
                                totalItems = mmcalls.Length

                                Dim x As Integer, y As Integer
                                Dim z As Integer

                                For x = 0 To UBound(mmcalls)
                                    z = 0
                                    For y = 0 To UBound(mmcalls) - 1
                                        'Si el elemento del array es igual al array temporal  
                                        If mmcalls(x) = tempArray(z) And y <> x Then
                                            'Entonces Eliminamos el valor duplicado  
                                            mmcalls(y) = ""
                                        End If
                                        z = z + 1
                                    Next y
                                Next x
                                canales = ""
                            End If
                            If maximo_largo_mmcall = 0 Then maximo_largo_mmcall = 40
                            eMensaje = Microsoft.VisualBasic.Strings.Left(eMensaje, maximo_largo_mmcall).Trim
                            mensajeGenerado = False
                            For i = 0 To UBound(mmcalls)
                                If mmcalls(i).Length > 0 Then
                                    'Se valida el reloj
                                    Dim posReloj = Strings.InStr(mmcalls(i), "number=")
                                    Dim iReloj = Strings.Mid(mmcalls(i), posReloj + 7)
                                    If Not valReloj(iReloj) Then
                                        agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: ID sin licencia", nroReporte, 9)
                                        Continue For
                                    End If
                                    respuestaWS = ""

                                    Try
                                        If validarURI(mmcalls(i) & "&message=" & eMensaje) Then
                                            Dim fr As System.Net.HttpWebRequest
                                            Dim targetURI As New Uri(mmcalls(i) & "&message=" & eMensaje)

                                            fr = DirectCast(HttpWebRequest.Create(targetURI), System.Net.HttpWebRequest)
                                            If (fr.GetResponse().ContentLength > 0) Then
                                                Dim str As New System.IO.StreamReader(fr.GetResponse().GetResponseStream())
                                                respuestaWS = str.ReadToEnd
                                                str.Close()
                                            End If
                                            mensajeGenerado = respuestaWS = "success" And Not mensajeGenerado
                                            If mensajeGenerado Then
                                                audiosGen = audiosGen + 1
                                            Else
                                                audiosNGen = audiosNGen + 1
                                                agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: " & respuestaWS, nroReporte, 9)
                                            End If
                                        Else
                                            audiosNGen = audiosNGen + 1
                                            agregarLOG("" & mmcalls(i) & "&message=" & eMensaje & " error: la dirección no es válida", nroReporte, 9)
                                        End If
                                    Catch ex As System.Net.WebException
                                        agregarLOG("Servicio de MMCall: " & mmcalls(i) & "&message=" & eMensaje & " error: " & ex.Message,  , 9)
                                        audiosNGen = audiosNGen + 1
                                        miError = ex.Message
                                    End Try
                                End If
                            Next
                        End If
                        'If mensajeGenerado Then
                        cadSQL = "UPDATE " & rutaBD & ".mensajes SET estatus = 'Z', enviada = NOW()  WHERE id = " & elmensaje!id
                        If optimizar Then
                            If elmensaje!cuenta > 1 Then
                                cadSQL = "UPDATE " & rutaBD & ".mensajes SET estatus = 'Z', enviada = NOW() WHERE canal = 3 AND lista = " & elmensaje!lista & " AND estatus = '" & idProceso & "'"
                                If mantenerPrioridad Then
                                    cadSQL = cadSQL & " AND prioridad = " & elmensaje!prioridad
                                End If
                            End If
                        End If
                        regsAfectados = consultaACT(cadSQL)
                        'End If
                    Next
                    If audiosGen > 0 Or audiosNGen > 0 Then
                        agregarLOG("Se generaron " & audiosGen & " mensaje(s) a MMCall y no se generaron " & audiosNGen & " mensaje(s) a MMCall ")
                    End If
                End If
            End If
        End If
        Application.Exit()
    End Sub
End Module