Imports MySql.Data.MySqlClient
Imports System.IO.Ports
Imports System.IO
Imports System.Text
Imports System.Net.Mail
Imports System.Net
Imports System.ComponentModel
Imports System.Data
Imports System.Windows.Forms
Imports DevExpress.XtraCharts
Imports DevExpress.XtraGauges.Win
Imports DevExpress.XtraGauges.Win.Base
Imports DevExpress.XtraGauges.Win.Gauges.Circular
Imports DevExpress.XtraGauges.Core.Model
Imports DevExpress.XtraGauges.Core.Base
Imports DevExpress.XtraGauges.Core.Drawing
Imports System.Drawing
Imports System.Drawing.Imaging


Public Class Form1

    Dim Estado As Integer = 0
    Dim procesandoAudios As Boolean = False
    Dim eSegundos = 0
    Dim procesandoEscalamientos As Boolean
    Dim procesandoRepeticiones As Boolean
    Dim estadoPrograma As Boolean
    Dim MensajeLlamada = ""
    Dim errorCorreos As String = ""
    Dim cad_consolidado As String = ""
    Dim bajo_color As String
    Dim medio_color As String
    Dim alto_color As String
    Dim escaladas_color As String
    Dim noatendio_color As String
    Dim alto_etiqueta As String
    Dim escaladas_etiqueta As String
    Dim noatendio_etiqueta As String
    Dim bajo_hasta As Integer
    Dim medio_hasta As Integer

    Public be_log_activar As Boolean = False
    Dim filtroAdicional As String
    Dim filtroFechas As String = ""
    Dim filtroOEE As String = ""

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim argumentos As String() = Environment.GetCommandLineArgs()
        estadoPrograma=True
        'cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
        'enviarReportes()
        If Process.GetProcessesByName _
          (Process.GetCurrentProcess.ProcessName).Length > 1 Then
        ElseIf argumentos.Length <= 1 Then
            MsgBox("String connection missing", MsgBoxStyle.Critical, "SIGMA")
        Else
            cadenaConexion = argumentos(1)
            Dim idProceso = Process.GetCurrentProcess.Id

            idProceso = Process.GetCurrentProcess.Id



            estadoPrograma = True
            enviarReportes()

        End If
        Application.Exit()
    End Sub

    Private Sub enviarReportes()
        'Se envía correo

        Dim cadSQL As String = "Select * FROM " & rutaBD & ".control WHERE fecha = '" & Format(Now, "yyyyMMddHH") & "' AND tipo = 5"
        Dim readerDS As DataSet = consultaSEL(cadSQL)
        If readerDS.Tables(0).Rows.Count > 0 Then
            Exit Sub
        End If
        Dim regsAfectados = 0
        'Escalada 4
        Dim miError As String = ""
        Dim correo_cuenta As String
        Dim correo_puerto As String
        Dim correo_ssl As Boolean
        Dim correo_clave As String
        Dim correo_host As String
        Dim rutaFiles As String
        Dim be_envio_reportes As Boolean = False

        cadSQL = "SELECT * FROM " & rutaBD & ".configuracion"
        readerDS = consultaSEL(cadSQL)
        If readerDS.Tables(0).Rows.Count > 0 Then
            Dim reader As DataRow = readerDS.Tables(0).Rows(0)
            be_idioma = ValNull(reader!idioma_defecto, "N")
            etiquetas()
            correo_cuenta = ValNull(reader!correo_cuenta, "A")
            correo_clave = ValNull(reader!correo_clave, "A")
            correo_puerto = ValNull(reader!correo_puerto, "A")
            correo_ssl = ValNull(reader!correo_ssl, "A") = "S"
            be_envio_reportes = ValNull(reader!be_envio_reportes, "A") = "S"
            correo_host = ValNull(reader!correo_host, "A")
            rutaFiles = ValNull(reader!ruta_archivos_enviar, "A")
            alto_etiqueta = ValNull(reader!alto_etiqueta, "A")
            escaladas_etiqueta = ValNull(reader!escaladas_etiqueta, "A")
            noatendio_etiqueta = ValNull(reader!noatendio_etiqueta, "A")
            cad_consolidado = ValNull(reader!cad_consolidado, "A")
            alto_color = ValNull(reader!alto_color, "A")
            medio_color = ValNull(reader!medio_color, "A")
            bajo_color = ValNull(reader!bajo_color, "A")
            escaladas_color = ValNull(reader!escaladas_color, "A")
            noatendio_color = ValNull(reader!noatendio_color, "A")
            bajo_hasta = ValNull(reader!bajo_hasta, "N")
            medio_hasta = ValNull(reader!medio_hasta, "N")
            be_log_activar = ValNull(reader!be_log_activar, "A") = "S"

        End If
        If be_envio_reportes Then
            If bajo_hasta = 0 Then bajo_hasta = 50
            If medio_hasta = 0 Then medio_hasta = 75
            If alto_etiqueta.Length = 0 Then alto_etiqueta = traduccion(2)
            If escaladas_etiqueta.Length = 0 Then escaladas_etiqueta = traduccion(2)
            If noatendio_etiqueta.Length = 0 Then noatendio_etiqueta = traduccion(3)
            alto_color = "#" & alto_color
            escaladas_color = "#" & escaladas_color
            noatendio_color = "#" & noatendio_color
            If alto_color.Length = 0 Then alto_color = System.Drawing.Color.LimeGreen.ToString
            If escaladas_color.Length = 0 Then escaladas_color = System.Drawing.Color.OrangeRed.ToString
            If noatendio_color.Length = 0 Then noatendio_color = System.Drawing.Color.Tomato.ToString

            If rutaFiles.Length = 0 Then
                rutaFiles = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Else
                rutaFiles = Strings.Replace(rutaFiles, "/", "\")
                If Not My.Computer.FileSystem.DirectoryExists(rutaFiles) Then
                    Try
                        My.Computer.FileSystem.CreateDirectory(rutaFiles)
                    Catch ex As Exception
                        rutaFiles = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    End Try
                End If
            End If
            If rutaFiles <> Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) Then
                For Each foundFile As String In My.Computer.FileSystem.GetFiles(
  rutaFiles, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.png")
                    Try
                        File.Delete(foundFile)
                    Catch ex2 As Exception

                    End Try

                    'Se mueven los archivos a otra carpeta
                Next

                For Each foundFile As String In My.Computer.FileSystem.GetFiles(
  rutaFiles, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.csv")
                    Try
                        File.Delete(foundFile)
                    Catch ex2 As Exception

                    End Try

                    'Se mueven los archivos a otra carpeta
                Next

            End If
            If Not estadoPrograma Then
                Exit Sub
            End If
            cadSQL = "Select * FROM " & rutaBD & ".cat_correos WHERE estatus = 'A' AND tipo < 2"
            'Se preselecciona la voz
            Dim indice = 0

            Dim mensajesDS As DataSet = consultaSEL(cadSQL)
            Dim mensajeGenerado = False
            Dim tMensajes = 0

            If mensajesDS.Tables(0).Rows.Count > 0 Then

                Dim enlazado = False
                Dim errorCorreo = ""
                Dim smtpServer As New SmtpClient()
                Try
                    smtpServer.Credentials = New Net.NetworkCredential(correo_cuenta, correo_clave)
                    smtpServer.Port = correo_puerto
                    smtpServer.Host = correo_host '"smtp.live.com" '"smtp.gmail.com"
                    smtpServer.EnableSsl = correo_ssl
                    enlazado = True
                Catch ex As Exception
                    errorCorreo = ex.Message
                End Try
                If enlazado Then
                    For Each elmensaje In mensajesDS.Tables(0).Rows
                        Dim envio = elmensaje!extraccion.Split(New Char() {";"c})
                        'Se busca si hay uno del día y hra
                        If envio(2).Length > 0 And envio(3).Length > 0 Then
                            Dim enviarDia As Boolean = False
                            Dim diaSemana = DateAndTime.Weekday(Now)
                            Dim cadFrecuencia As String = traduccion(5)
                            If envio(2) = "T" Then
                                enviarDia = True
                            ElseIf envio(2) = "LV" And diaSemana >= 2 And diaSemana <= 6 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(6)
                            ElseIf envio(2) = "L" And diaSemana = 2 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(7)
                            ElseIf envio(2) = "M" And diaSemana = 3 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(8)
                            ElseIf envio(2) = "MI" And diaSemana = 4 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(9)
                            ElseIf envio(2) = "J" And diaSemana = 5 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(10)
                            ElseIf envio(2) = "V" And diaSemana = 6 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(11)
                            ElseIf envio(2) = "S" And diaSemana = 7 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(12)
                            ElseIf envio(2) = "D" And diaSemana = 1 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(13)
                            ElseIf envio(2) = "1M" And Val(Today.Day) = 1 Then
                                enviarDia = True
                                cadFrecuencia = traduccion(14)
                            ElseIf envio(2) = "UM" And Val(Today.Day) = Date.DaysInMonth(Today.Year, Today.Month) Then
                                enviarDia = True
                                cadFrecuencia = traduccion(15)
                            End If

                            'eemv
                            'enviarDia = True


                            If enviarDia Then
                                Dim enviar As Boolean = False
                                Dim hora = Val(Format(Now, "HH"))
                                If envio(3) = "T" Then
                                    enviar = True
                                    cadFrecuencia = cadFrecuencia & traduccion(16)
                                ElseIf Val(envio(3)) = Val(hora) Then
                                    cadFrecuencia = cadFrecuencia & IIf(Val(hora) = 1, traduccion(18), traduccion(19) & Val(hora) & traduccion(17))
                                    enviar = True
                                End If


                                'eemv
                                'enviar = True


                                If enviar Then
                                    Dim mail As New MailMessage
                                    Try
                                        Dim cuerpo As String = ValNull(elmensaje!cuerpo, "A")
                                        Dim titulo As String = ValNull(elmensaje!titulo, "A")
                                        Dim ordenPareto As Integer = ValNull(elmensaje!orden, "N")
                                        If titulo.Length = 0 Then titulo = traduccion(20)
                                        If cuerpo.Length = 0 Then cuerpo = traduccion(21)

                                        mail.From = New MailAddress(correo_cuenta) 'TextBox1.Text & "@gmail.com")
                                        Dim mails As String = ValNull(elmensaje!para, "A")
                                        Dim mailsV As String() = mails.Split(New Char() {";"c})
                                        For Each cuenta In mailsV
                                            If cuenta.Length > 0 Then
                                                cuenta = Strings.Replace(cuenta, vbCrLf, "")
                                                cuenta = Strings.Replace(cuenta, vbLf, "")
                                                mail.To.Add(cuenta)
                                            End If
                                        Next
                                        mails = ValNull(elmensaje!copia, "A")
                                        mailsV = mails.Split(New Char() {";"c})
                                        For Each cuenta In mailsV
                                            If cuenta.Length > 0 Then
                                                cuenta = Strings.Replace(cuenta, vbCrLf, "")
                                                cuenta = Strings.Replace(cuenta, vbLf, "")
                                                mail.CC.Add(cuenta)
                                            End If
                                        Next
                                        mails = ValNull(elmensaje!oculta, "A")
                                        mailsV = mails.Split(New Char() {";"c})
                                        For Each cuenta In mailsV
                                            If cuenta.Length > 0 Then
                                                cuenta = Strings.Replace(cuenta, vbCrLf, "")
                                                cuenta = Strings.Replace(cuenta, vbLf, "")
                                                mail.Bcc.Add(cuenta)
                                            End If
                                        Next
                                        mail.Subject = titulo
                                        errorCorreos = ""
                                        cuerpo = cuerpo & Environment.NewLine & traduccion(22)

                                        cadSQL = "SELECT a.reporte, b.nombre, b.grafica, b.file_name, b.grafica FROM " & rutaBD & ".det_correo a INNER JOIN " & rutaBD & ".int_listados b ON a.reporte = b.id AND idioma = " & be_idioma & " WHERE a.correo = " & elmensaje!id & " ORDER BY b.orden"
                                        mensajesDS = consultaSEL(cadSQL)
                                        If mensajesDS.Tables(0).Rows.Count > 0 Then
                                            For Each reporte In mensajesDS.Tables(0).Rows
                                                Dim miReporte = generarReporte(reporte!reporte, reporte!nombre, reporte!file_name, envio(0), envio(1), rutaFiles, reporte!grafica, ordenPareto, elmensaje!consulta, elmensaje!tipo)
                                                If miReporte = -1 Then
                                                    cuerpo = cuerpo & Environment.NewLine & reporte!nombre & traduccion(23) & errorCorreos
                                                Else
                                                    If My.Computer.FileSystem.FileExists(rutaFiles & "\" & reporte!file_name & ".csv") Then
                                                        cuerpo = cuerpo & Environment.NewLine & reporte!nombre
                                                        Dim archivo As Attachment = New Attachment(rutaFiles & "\" & reporte!file_name & ".csv")
                                                        mail.Attachments.Add(archivo)
                                                    End If
                                                    If My.Computer.FileSystem.FileExists(rutaFiles & "\" & reporte!file_name & ".png") Then

                                                        Dim archivo As Attachment = New Attachment(rutaFiles & "\" & reporte!file_name & ".png")
                                                        mail.Attachments.Add(archivo)
                                                    End If
                                                End If
                                            Next
                                        End If
                                        cuerpo = cadFrecuencia & Environment.NewLine & Environment.NewLine & cuerpo
                                        mail.Body = cuerpo
                                        smtpServer.Send(mail)

                                        tMensajes = tMensajes + 1
                                        mensajeGenerado = True
                                    Catch ex As Exception
                                        agregarLOG(traduccion(25) & ex.Message, 0, 9)
                                    End Try
                                Else
                                    mensajeGenerado = True
                                End If
                            End If
                        End If
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_correos SET ultimo_envio = '" & Format(Now, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & elmensaje!id)

                    Next
                End If
                If enlazado Then
                    If tMensajes > 0 Then
                        agregarLOG(traduccion(26).Replace("campo_0", tMensajes))
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".control (fecha, mensajes, tipo) VALUES ('" & Format(Now, "yyyyMMddHH") & "', " & tMensajes & ", 5)")
                    Else
                        agregarLOG(traduccion(27))
                    End If

                Else
                    agregarLOG(traduccion(25) & errorCorreo, 0, 9)
                End If
                smtpServer.Dispose()
            End If
        End If
    End Sub

    Function generarReporte(idReporte As Integer, reporte As String, fName As String, periodo As String, nperiodos As Integer, ruta As String, graficar As String, ordenPareto As Integer, consulta As Long, tipoReporte As Integer) As Integer
        generarReporte = 0

        Dim archivoSaliente = ruta & "\" & fName & ".csv"
        Dim archivoImagen = ruta & "\" & fName & ".png"

        Try
            File.Delete(archivoSaliente)
            File.Delete(archivoSaliente)


        Catch ex As Exception

        End Try

        Try
            File.Delete(archivoImagen)
            File.Delete(archivoImagen)


        Catch ex As Exception

        End Try

        archivoSaliente = archivoSaliente.Replace("\", "\\")

        Dim eDesde = Now()
        Dim eHasta = Now()
        Dim ePeriodo = nperiodos
        Dim diaSemana = DateAndTime.Weekday(Now)
        Dim intervalo = DateInterval.Second
        Dim cadPeriodo As String = nperiodos & traduccion(28)
        If periodo = 1 Then
            intervalo = DateInterval.Minute
            cadPeriodo = nperiodos & traduccion(29)
        ElseIf periodo = 2 Then
            intervalo = DateInterval.Hour
            cadPeriodo = nperiodos & traduccion(30)
        ElseIf periodo = 3 Then
            intervalo = DateInterval.Day
            cadPeriodo = nperiodos & traduccion(31)
        ElseIf periodo = 4 Then
            intervalo = DateInterval.Day
            ePeriodo = 6
            cadPeriodo = nperiodos & traduccion(32)
        ElseIf periodo = 5 Then
            intervalo = DateInterval.Month
            cadPeriodo = nperiodos & traduccion(33)
        ElseIf periodo = 6 Then
            intervalo = DateInterval.Year
            cadPeriodo = nperiodos & traduccion(34)
        ElseIf periodo = 10 Then
            eDesde = CDate(Format(Now, "yyyy/MM/dd") & " 00:00:00")
            cadPeriodo = traduccion(35)
        ElseIf periodo = 11 Then
            cadPeriodo = traduccion(36)
            If diaSemana = 0 Then
                eDesde = CDate(Format(DateAdd(DateInterval.Day, -6, Now), "yyyy/MM/dd") & " 00:00:00")
            Else
                eDesde = CDate(Format(DateAdd(DateInterval.Day, (diaSemana - 2) * -1, Now), "yyyy/MM/dd") & " 00:00:00")
            End If
        ElseIf periodo = 12 Then
            cadPeriodo = traduccion(37)
            eDesde = CDate(Format(Now, "yyyy/MM") & "/01 00:00:00")
        ElseIf periodo = 13 Then
            cadPeriodo = traduccion(38)
            eDesde = CDate(Format(Now, "yyyy") & "/01/01 00:00:00")
        ElseIf periodo = 20 Then
            cadPeriodo = traduccion(39)
            eDesde = CDate(Format(DateAdd(DateInterval.Day, -1, Now), "yyyy/MM/dd") & " 00:00:00")
            eHasta = CDate(Format(DateAdd(DateInterval.Day, -1, Now), "yyyy/MM/dd") & " 23:59:59")
        ElseIf periodo = 21 Then
            cadPeriodo = traduccion(40)
            Dim dayDiff As Integer = Date.Today.DayOfWeek - DayOfWeek.Monday
            eDesde = CDate(Format(Today.AddDays(-dayDiff), "yyyy/MM/dd") & " 00:00:00")
            eDesde = DateAdd(DateInterval.Day, -7, CDate(eDesde))
            eHasta = DateAdd(DateInterval.Day, 6, CDate(eDesde))
        ElseIf periodo = 22 Then
            cadPeriodo = traduccion(41)
            eDesde = CDate(Format(DateAdd(DateInterval.Month, -1, Now), "yyyy/MM") & "/01 00:00:00")
            eHasta = CDate(Format(DateAdd(DateInterval.Day, -1, CDate(Format(Now, "yyyy/MM") & "/01")), "yyyy/MM/dd") & " 23:59:59")
        End If
        If periodo < 10 Then eDesde = DateAdd(intervalo, ePeriodo * -1, eDesde)
        If periodo = 3 Then
            eDesde = Format(eDesde, "yyyy/MM/dd ") & "00:00:00"
            eHasta = CDate(Format(DateAdd(DateInterval.Day, -1, Now), "yyyy/MM/dd") & " 23:59:59")
        End If
        Dim fDesdeSF = Format(eDesde, "yyyy/MM/dd")
        Dim fHastaSF = Format(eHasta, "yyyy/MM/dd")
        filtroOEE = " AND i.dia >= '" & fDesdeSF & "' AND i.dia <= '" & fHastaSF & "' "
        Dim comillas = Microsoft.VisualBasic.Strings.Left(Chr(34), 1)

        Dim inicial = ""
        Dim cabecera = ""
        Dim registros = ""

        Dim Leer As Boolean = False

        generarFIltro(consulta)

        Dim sentencia = "SELECT i.id, i.dia, a.numero, IFNULL(b.nombre, '" & traduccion(100) & "'), i.turno, IFNULL(c.nombre, '" & traduccion(100) & "'), i.equipo, IFNULL(d.nombre, '" & traduccion(100) & "'), d.referencia, i.parte, i.produccion, i.sacos, i.tarimas, i.tiempo_disponible, i.paro, IFNULL(e.nombre, '" & traduccion(100) & "'), i.bloque_inicia, i.bloque_finaliza FROM " & rutaBD & ".lecturas_cortes i LEFT JOIN " & rutaBD & ".lotes a ON i.orden = a.id LEFT JOIN " & rutaBD & ".cat_turnos b ON i.turno = b.id LEFT JOIN " & rutaBD & ".cat_maquinas c ON i.equipo = c.id LEFT JOIN " & rutaBD & ".cat_partes d ON i.parte = d.id LEFT JOIN " & rutaBD & ".cat_usuarios e ON i.operador = e.id WHERE i.id > 0 " & filtroOEE

        cabecera = Chr(34) & traduccion(79) & Chr(34) & "," & Chr(34) & traduccion(139) & Chr(34) & "," & Chr(34) & traduccion(140) & Chr(34) & "," & Chr(34) & traduccion(126) & Chr(34) & "," & Chr(34) & traduccion(69) & Chr(34) & "," & Chr(34) & traduccion(141) & Chr(34) & "," & Chr(34) & traduccion(124) & Chr(34) & "," & Chr(34) & traduccion(142) & Chr(34) & "," & Chr(34) & traduccion(122) & Chr(34) & "," & Chr(34) & traduccion(143) & Chr(34) & "," & Chr(34) & traduccion(144) & Chr(34) & "," & Chr(34) & traduccion(128) & Chr(34) & "," & Chr(34) & traduccion(129) & Chr(34) & "," & Chr(34) & traduccion(130) & Chr(34) & "," & Chr(34) & traduccion(143) & Chr(34) & "," & Chr(34) & traduccion(144) & Chr(34) & "," & Chr(34) & traduccion(121) & Chr(34) & "," & Chr(34) & traduccion(145) & Chr(34) & "," & Chr(34) & traduccion(146) & Chr(34)

        inicial = Chr(34) & reporte & " (" & cadPeriodo & ")" & Chr(34) & Environment.NewLine
        inicial = inicial & Chr(34) & traduccion(110) & ": " & Format(Now(), "ddd, dd-MMM-yyyy HH:mm:ss") & Chr(34) & Environment.NewLine
        inicial = inicial & Chr(34) & traduccion(111) & ": " & Format(eDesde, "dd/MMM/yyyy HH:mm:ss") & " " & traduccion(112) & ": " & Format(eHasta, "dd/MMM/yyyy HH:mm:ss") & Chr(34) & Environment.NewLine

        Dim mensajesDS As DataSet
        Dim adicional = ""



        If idReporte = 99 Then
            mensajesDS = consultaSEL(sentencia)
            If mensajesDS.Tables(0).Rows.Count > 0 Then
                Dim objWriter As New System.IO.StreamWriter(archivoSaliente, False, System.Text.Encoding.UTF8)
                objWriter.WriteLine(inicial)
                objWriter.WriteLine(cabecera)
                Dim linea = 0
                For Each registro In mensajesDS.Tables(0).Rows
                    linea = linea + 1
                    objWriter.WriteLine(Chr(34) & linea & Chr(34) & "," & Chr(34) & registro.item(0) & Chr(34) & "," & Chr(34) & registro.item(1) & Chr(34) & "," & Chr(34) & ValNull(registro.item(2), "A") & Chr(34) & "," & Chr(34) & ValNull(registro.item(3), "A") & Chr(34) & "," & Chr(34) & registro.item(4) & Chr(34) & "," & Chr(34) & registro.item(5) & Chr(34) & "," & Chr(34) & registro.item(6) & Chr(34) & "," & Chr(34) & registro.item(7) & Chr(34) & "," & Chr(34) & registro.item(8) & Chr(34) & "," & Chr(34) & registro.item(9) & Chr(34) & "," & Chr(34) & registro.item(10) & Chr(34) & "," & Chr(34) & registro.item(11) & Chr(34) & "," & Chr(34) & registro.item(12) & Chr(34) & "," & Chr(34) & registro.item(13) & Chr(34) & "," & Chr(34) & registro.item(14) & Chr(34) & "," & Chr(34) & registro.item(15) & Chr(34) & "," & Chr(34) & registro.item(16) & Chr(34) & "," & Chr(34) & registro.item(17) & Chr(34))
                Next
                objWriter.WriteLine(traduccion(134) & ": " & linea)
                If adicional.Length > 0 Then objWriter.WriteLine(adicional)
                objWriter.WriteLine(traduccion(136))
                objWriter.Close()

            End If
        Else

            Dim cadSQL = "SELECT * FROM " & rutaBD & ".pu_graficos WHERE (usuario = 1 OR usuario = 0) AND grafico = " & 100 + idReporte & " ORDER BY usuario DESC LIMIT 1"

            Dim tHaving = ""

            Dim config As DataSet = consultaSEL(cadSQL)
            If config.Tables(0).Rows.Count > 0 Then
                If config.Tables(0).Rows(0)!incluir_ceros = "N" Then
                    tHaving = " HAVING piezas_m > 0 "
                End If

                Dim ordenDatos = " 6 DESC"
                If config.Tables(0).Rows(0)!orden = 1 Then
                    ordenDatos = " 7 DESC"
                ElseIf config.Tables(0).Rows(0)!orden = 2 Then
                    ordenDatos = " 8 DESC"
                End If
                If config.Tables(0).Rows(0)!orden_grafica = "N" Then
                    ordenDatos = " 6 "
                    If config.Tables(0).Rows(0)!orden = 1 Then
                        ordenDatos = " 7 "
                    ElseIf config.Tables(0).Rows(0)!orden = 2 Then
                        ordenDatos = " 8 "
                    End If

                ElseIf config.Tables(0).Rows(0)!orden_grafica = "A" Then
                    ordenDatos = " 4 "
                End If


                Dim cadTabla = "cat_turnos"
                Dim cadReferencia = "'' AS referencia"
                Dim cadCampoResumen = "i.turno"
                Dim cadTitulo = traduccion(123)
                Dim cadTituloReferencia = traduccion(103)

                If idReporte = 2 Then
                    cadTabla = "cat_maquinas"
                    cadReferencia = "'' AS referencia"
                    cadCampoResumen = "i.equipo"
                    cadTitulo = traduccion(124)
                ElseIf idReporte = 7 Then
                    cadTabla = "cat_partes"
                    cadReferencia = "a.referencia"
                    cadCampoResumen = "i.parte"
                    cadTitulo = traduccion(122)
                ElseIf idReporte = 8 Then
                    cadTabla = "cat_usuarios"
                    cadReferencia = "a.referencia"
                    cadCampoResumen = "i.operador"
                    cadTitulo = traduccion(121)
                End If
                sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.nombre, '" & traduccion(100) & "') AS nombre, " & cadReferencia & ", IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, IFNULL(i.disponible, 0) AS disponible_m, IFNULL(i.paros, 0) AS paros_m, 0 AS porcentaje FROM " & rutaBD & "." & cadTabla & " a LEFT JOIN (SELECT " & cadCampoResumen & ", SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY " & cadCampoResumen & ") AS i ON " & cadCampoResumen & " = a.id " & tHaving & " ORDER BY " & ordenDatos

                If idReporte = 6 Then
                    cadTitulo = traduccion(126)
                    cadTituloReferencia = traduccion(127)
                    sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.numero, '" & traduccion(100) & "') AS nombre, b.notas AS referencia, IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, IFNULL(i.disponible, 0) AS disponible_m, IFNULL(i.paros, 0) AS paros_m, 0 AS porcentaje FROM " & rutaBD & ".lotes a LEFT JOIN " & rutaBD & ".equipos_objetivo b ON a.id = b.lote LEFT JOIN (SELECT i.orden, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0  " & filtroOEE & " GROUP BY i.orden) AS i ON i.orden = a.id " & tHaving & " ORDER BY " & ordenDatos
                ElseIf idReporte = 3 Then
                    cadTitulo = traduccion(51)
                    cadReferencia = ""
                    sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, DATE_FORMAT(i.dia, '%Y/%m/%d') AS nombre, '' AS referencia, SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, SUM(i.tiempo_disponible) AS disponible_m, SUM(i.paro) AS paros_m, 0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                ElseIf idReporte = 4 Then
                    cadReferencia = ""
                    sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, CONCAT(YEAR(i.dia), '/', WEEK(i.dia)) AS nombre, STR_TO_DATE(CONCAT(DATE_FORMAT(i.dia,'%x/%v'), ' Monday'), '%x/%v %W') AS referencia, SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, SUM(i.tiempo_disponible) AS disponible_m, SUM(i.paro) AS paros_m, 0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                ElseIf idReporte = 5 Then
                    cadTitulo = traduccion(54)
                    cadReferencia = ""
                    sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, CONCAT(YEAR(i.dia), '/', MONTH(i.dia)) AS nombre, '' AS referencia, SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, SUM(i.tiempo_disponible) AS disponible_m, SUM(i.paro) AS paros_m, 0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                End If

                If tipoReporte = 1 Then
                    idReporte = idReporte - 100
                    Dim cadTiempo = "IFNULL(i.disponible / 3600, 0) AS disponible_m, IFNULL(i.paros / 3600, 0) AS paros_m, "
                    Dim cadTiempo2 = "SUM(i.tiempo_disponible / 3600) AS disponible_m, SUM(i.paro / 3600) AS paros_m, "
                    ordenDatos = " 9 DESC"
                    If config.Tables(0).Rows(0)!incluir_ceros = "N" Then
                        tHaving = " HAVING disponible_m > 0 "
                    End If
                    If config.Tables(0).Rows(0)!orden_grafica = "N" Then
                        ordenDatos = " 9 "
                    ElseIf (config.Tables(0).Rows(0)!orden_grafica = "A") Then
                        ordenDatos = " 1 "
                    End If
                    If config.Tables(0).Rows(0)!orden = 1 Then
                        cadTiempo = "IFNULL(i.disponible / 60, 0) AS disponible_m, IFNULL(i.paros / 60, 0) AS paros_m, "
                        cadTiempo2 = "SUM(i.tiempo_disponible / 60) AS disponible_m, SUM(i.paro / 60) AS paros_m, "
                    ElseIf config.Tables(0).Rows(0)!orden = 2 Then
                        cadTiempo = "IFNULL(i.disponible, 0) AS disponible_m, IFNULL(i.paros, 0) AS paros_m, "
                        cadTiempo2 = "SUM(i.tiempo_disponible) AS disponible_m, SUM(i.paro) AS paros_m, "
                    End If
                    sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.nombre, '" & traduccion(100) & "') AS nombre, " & cadReferencia & ", IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, " & cadTiempo & "0 AS porcentaje FROM " & rutaBD & ".cat_turnos a LEFT JOIN (SELECT i.turno, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY i.turno) AS i ON i.turno = a.id " & tHaving & " ORDER BY " & ordenDatos
                    If idReporte = 2 Then
                        sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.nombre, '" & traduccion(100) & "') AS nombre, " & cadReferencia & ", IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, " & cadTiempo & "0 AS porcentaje FROM " & rutaBD & ".cat_maquinas a LEFT JOIN (SELECT i.equipo, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY i.equipo) AS i ON i.equipo = a.id " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 3 Then
                        sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, DATE_FORMAT(i.dia, '%Y/%m/%d') AS nombre, " & cadReferencia & ", SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, " + cadTiempo2 + "0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 4 Then
                        sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, CONCAT(YEAR(i.dia), '/', WEEK(i.dia)) AS nombre, STR_TO_DATE(CONCAT(DATE_FORMAT(i.dia,'%x/%v'), ' Monday'), '%x/%v %W') AS referencia, SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, " + cadTiempo2 + "0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 5 Then
                        sentencia = "SELECT 0 AS id, 0 AS orden, 1 AS filtro, CONCAT(YEAR(i.dia), '/', MONTH(i.dia)) AS nombre, " & cadReferencia & ", SUM(i.produccion) AS piezas_m, SUM(i.sacos) AS sacos_m, SUM(i.tarimas) AS tarimas_m, " + cadTiempo2 + "0 AS porcentaje FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY nombre " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 6 Then
                        sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.numero, '" & traduccion(100) & "') AS nombre, b.notas AS referencia, IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, " & cadTiempo & "0 AS porcentaje FROM " & rutaBD & ".lotes a LEFT JOIN " & rutaBD & ".equipos_objetivo b ON a.id = b.lote LEFT JOIN (SELECT i.orden, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY i.orden) AS i ON i.orden = a.id " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 7 Then
                        sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.nombre, '" & traduccion(100) & "') AS nombre, " & cadReferencia & ", IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, " & cadTiempo & "0 AS porcentaje FROM " & rutaBD & ".cat_partes a LEFT JOIN (SELECT i.parte, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY i.parte) AS i ON i.parte = a.id " & tHaving & " ORDER BY " & ordenDatos
                    ElseIf idReporte = 8 Then
                        sentencia = "SELECT a.id, 0 AS orden, 1 AS filtro, IFNULL(a.nombre, '" & traduccion(100) & "') AS nombre, " & cadReferencia & ", IFNULL(i.piezas, 0) AS piezas_m, IFNULL(i.sacos, 0) AS sacos_m, IFNULL(i.tarimas, 0) AS tarimas_m, " & cadTiempo & "0 AS porcentaje FROM " & rutaBD & ".cat_usuarios a LEFT JOIN (SELECT i.operador, SUM(i.paro) AS paros, SUM(i.produccion) AS piezas, SUM(i.tiempo_disponible) AS disponible, SUM(i.sacos) AS sacos, SUM(i.tarimas) AS tarimas FROM " & rutaBD & ".lecturas_cortes i WHERE i.id > 0 " & filtroOEE & " GROUP BY i.operador) AS i ON i.operador = a.id " & tHaving & " ORDER BY " & ordenDatos
                    End If
                End If
                cabecera = Chr(34) & traduccion(79) & Chr(34) & "," & Chr(34) & traduccion(125) & Chr(34) & "," & Chr(34) & cadTitulo & Chr(34) & "," & Chr(34) & cadTituloReferencia & Chr(34) & "," & Chr(34) & traduccion(128) & Chr(34) & "," & Chr(34) & traduccion(129) & Chr(34) & "," & Chr(34) & traduccion(130) & Chr(34) & "," & Chr(34) & traduccion(131) & Chr(34) & "," & Chr(34) & traduccion(132) & Chr(34) & "," & Chr(34) & traduccion(109) & Chr(34) & "," & Chr(34) & traduccion(107) & Chr(34)

                Dim campoSumar = "piezas_m"

                mensajesDS = consultaSEL(sentencia)
                adicional = ""
                If mensajesDS.Tables(0).Rows.Count > 0 Then
                    Dim view_o As DataView = New DataView(mensajesDS.Tables(0))
                    Dim objWriter As New System.IO.StreamWriter(archivoSaliente, False, System.Text.Encoding.UTF8)
                    objWriter.WriteLine(inicial)
                    objWriter.WriteLine(cabecera)
                    Dim totalPareto = 0
                    If tipoReporte = 1 Then
                        campoSumar = "disponible_m"
                    ElseIf config.Tables(0).Rows(0)!orden = 1 Then
                        campoSumar = "sacos_m"
                    ElseIf config.Tables(0).Rows(0)!orden = 2 Then
                        campoSumar = "trimas_m"
                    End If
                    If config.Tables(0).Rows(0)!maximo_barras > 0 And config.Tables(0).Rows(0)!maximo_barras < mensajesDS.Tables(0).Rows.Count Or config.Tables(0).Rows(0)!maximo_barraspct > 0 And config.Tables(0).Rows(0)!maximo_barraspct < 100 Then
                        'Se calcula el total del Pareto


                        For Each elmensaje As DataRow In mensajesDS.Tables(0).Rows
                            totalPareto = totalPareto + elmensaje.Item(campoSumar)
                        Next
                        Dim limitar = 0
                        Dim agrupado = ""
                        Dim pcAcum = 0
                        Dim pct = config.Tables(0).Rows(0)!maximo_barraspct / 100
                        Dim i = 0
                        For Each elmensaje In mensajesDS.Tables(0).Rows
                            i = i + 1
                            pcAcum = pcAcum + elmensaje.Item(campoSumar)
                            If pcAcum / totalPareto >= pct Then
                                limitar = i
                                Exit For
                            End If
                        Next

                        If config.Tables(0).Rows(0)!maximo_barras > 0 Then
                            If limitar > config.Tables(0).Rows(0)!maximo_barras Or limitar = 0 Then
                                limitar = config.Tables(0).Rows(0)!maximo_barras
                            End If
                        End If

                        If limitar + 1 >= mensajesDS.Tables(0).Rows.Count And config.Tables(0).Rows(0)!agrupar = "S" Then
                            limitar = 0
                        ElseIf limitar >= mensajesDS.Tables(0).Rows.Count Then
                            limitar = 0
                        End If
                        If limitar > 0 Then
                            For j = 0 To limitar - 1

                                mensajesDS.Tables(0).Rows(i)!orden = j + 1
                            Next
                            If config.Tables(0).Rows(0)!agrupar = "S" Then

                                Dim faltante1 = 0
                                Dim faltante2 = 0
                                Dim faltante3 = 0
                                Dim faltante4 = 0
                                Dim faltante5 = 0
                                Dim totalAgr = 0
                                For j = limitar To mensajesDS.Tables(0).Rows.Count - 1
                                    mensajesDS.Tables(0).Rows(j)!filtro = 0
                                    faltante1 = faltante1 + mensajesDS.Tables(0).Rows(j)!piezas_m
                                    faltante2 = faltante2 + mensajesDS.Tables(0).Rows(j)!sacos_m
                                    faltante3 = faltante3 + mensajesDS.Tables(0).Rows(j)!tarimas_m
                                    faltante4 = faltante4 + mensajesDS.Tables(0).Rows(j)!disponible_m
                                    faltante5 = faltante5 + mensajesDS.Tables(0).Rows(j)!paros_m
                                Next
                                totalAgr = mensajesDS.Tables(0).Rows.Count - limitar
                                Dim row As DataRow = mensajesDS.Tables(0).NewRow

                                row("id") = "0"
                                row("nombre") = IIf(ValNull(config.Tables(0).Rows(0)!agrupar_texto, "A") = "", traduccion(61), config.Tables(0).Rows(0)!agrupar_texto) & " (" & totalAgr & ")"
                                row("referencia") = ""
                                row("piezas_m") = faltante1
                                row("sacos_m") = faltante2
                                row("tarimas_m") = faltante3
                                row("disponible_m") = faltante4
                                row("paros_m") = faltante5
                                row("porcentaje") = 0
                                row("filtro") = 1
                                If config.Tables(0).Rows(0)!agrupar_posicion = "F" Then
                                    row("orden") = limitar + 1
                                ElseIf config.Tables(0).Rows(0)!agrupar_posicion = "P" Then
                                    row("orden") = 0
                                End If
                                mensajesDS.Tables(0).Rows.Add(row)
                            Else
                                adicional = traduccion(135)
                            End If
                        End If
                        If config.Tables(0).Rows(0)!agrupar_posicion = "N" Then
                            view_o.Sort = IIf(config.Tables(0).Rows(0)!orden_grafica = "M", " " & campoSumar & " DESC", IIf(config.Tables(0).Rows(0)!orden_grafica = "N", " " & campoSumar, "nombre"))
                        Else
                            view_o.Sort = "orden ASC"
                        End If
                    Else
                        view_o.Sort = IIf(config.Tables(0).Rows(0)!orden_grafica = "M", " " & campoSumar & " DESC", IIf(config.Tables(0).Rows(0)!orden_grafica = "N", " " & campoSumar, "nombre"))
                    End If
                    view_o.RowFilter = "filtro = 1"
                    totalPareto = 0
                    For Each registro In mensajesDS.Tables(0).Rows
                        If registro!filtro = 1 Then
                            totalPareto = totalPareto + registro.Item(campoSumar)
                        End If
                    Next
                    Dim acumPareto As Double = 0
                    Dim linea = 0
                    Dim tablaGRafico As DataTable = view_o.ToTable()
                    For Each registro In tablaGRafico.Rows
                        If registro!filtro = 1 Then
                            linea = linea + 1
                            acumPareto = acumPareto + registro.Item(campoSumar) / totalPareto * 100
                            registro!porcentaje = acumPareto
                            If linea = mensajesDS.Tables(0).Rows.Count Then
                                acumPareto = 100
                            End If
                            objWriter.WriteLine(Chr(34) & linea & Chr(34) & "," & Chr(34) & registro!id & Chr(34) & "," & Chr(34) & registro!nombre & Chr(34) & "," & Chr(34) & registro!referencia & Chr(34) & "," & Chr(34) & registro!piezas_m & Chr(34) & "," & Chr(34) & registro!sacos_m & Chr(34) & "," & Chr(34) & registro!tarimas_m & Chr(34) & "," & Chr(34) & registro!disponible_m & Chr(34) & "," & Chr(34) & registro!paros_m & Chr(34) & "," & Chr(34) & registro.Item(campoSumar) / totalPareto * 100 & Chr(34) & "," & Chr(34) & acumPareto & Chr(34))
                        End If
                    Next
                    objWriter.WriteLine(traduccion(134) & ": " & linea)
                    If adicional.Length > 0 Then objWriter.WriteLine(adicional)
                    objWriter.WriteLine(traduccion(136))
                    objWriter.Close()

                    If graficar = "S" Then


                        Dim indicador01 = IIf(tipoReporte = 0, traduccion(137), traduccion(138))
                        Dim indicador02 = traduccion(109)

                        Dim titulosSeries = config.Tables(0).Rows(0)!textos_adicionales.Split(New Char() {";"c})
                        If titulosSeries.length = 1 Then
                            indicador01 = titulosSeries(0)
                        End If
                        If titulosSeries.length = 2 Then
                            indicador01 = titulosSeries(0)
                            indicador02 = titulosSeries(1)
                        End If


                        ChartControl1.Series.Clear()
                        ChartControl1.Titles.Clear()
                        Dim Titulo As New ChartTitle()
                        Titulo.Text = config.Tables(0).Rows(0)!titulo & Strings.Space(10)
                        Dim miFuente = New Drawing.Font("Lucida Sans", 10, FontStyle.Regular)
                        Dim miFuenteAlto = New Drawing.Font("Lucida Sans", 16, FontStyle.Bold)
                        Dim miFuenteEjes = New Drawing.Font("Lucida Sans", 11, FontStyle.Regular)

                        Titulo.Font = miFuenteAlto
                        Dim series1 As New Series(indicador01, ViewType.Bar)

                        ChartControl1.Series.Add(series1)
                        series1.DataSource = tablaGRafico
                        series1.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True
                        series1.View.Color = Color.SkyBlue
                        series1.ArgumentScaleType = ScaleType.Qualitative
                        series1.ArgumentDataMember = "nombre"
                        series1.ValueScaleType = ScaleType.Numerical
                        series1.ValueDataMembers.AddRange(New String() {campoSumar})
                        series1.Label.BackColor = Color.DarkBlue
                        series1.Label.TextColor = Color.White
                        series1.Label.Font = miFuente
                        series1.Label.TextPattern = "{V:F1}"

                        If config.Tables(0).Rows(0)!grueso_spiline > 0 Then
                            Dim series2 As New Series(indicador02, ViewType.Spline)

                            ChartControl1.Series.Add(series2)
                            series2.DataSource = tablaGRafico
                            series2.LabelsVisibility = DevExpress.Utils.DefaultBoolean.True
                            series2.View.Color = Color.Green

                            series2.ArgumentScaleType = ScaleType.Qualitative
                            series2.ArgumentDataMember = "nombre"
                            series2.ValueScaleType = ScaleType.Numerical
                            series2.ValueDataMembers.AddRange(New String() {"porcentaje"})
                            series2.Label.BackColor = Color.DarkBlue
                            series2.Label.TextColor = Color.White
                            series2.Label.Font = miFuente
                            series2.Label.TextPattern = "{V:F1}"

                            'Obtener el titulo en Y
                            Dim tituloY As String = ""
                            Dim titulosY = ValNull(config.Tables(0).Rows(0)!texto_y, "A")
                            Try

                                Dim titulosYArreglo = titulosY.Split(New Char() {";"c})
                                If titulosYArreglo.length = 1 Then
                                    tituloY = titulosYArreglo(0)
                                Else
                                    tituloY = titulosYArreglo(ordenPareto)
                                End If
                            Catch ex As Exception
                                tituloY = titulosY
                            End Try
                            CType(series2.View, SplineSeriesView).LineStyle.Thickness = config.Tables(0).Rows(0)!grueso_spiline
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Visibility = DevExpress.Utils.DefaultBoolean.True
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Label.Font = miFuenteEjes
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.GridSpacingAuto = False
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.GridSpacing = 1
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Title.Text = tituloY
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Title.Font = miFuenteAlto
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.GridLines.Visible = False
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Tickmarks.Visible = False
                            CType(ChartControl1.Diagram, XYDiagram).AxisY.Tickmarks.MinorVisible = False

                            Dim myAxisY As New SecondaryAxisY(traduccion(63))
                            CType(ChartControl1.Diagram, XYDiagram).SecondaryAxesY.Clear()
                            CType(ChartControl1.Diagram, XYDiagram).SecondaryAxesY.Add(myAxisY)
                            CType(series2.View, LineSeriesView).AxisY = myAxisY
                            myAxisY.Title.Text = config.Tables(0).Rows(0)!texto_z
                            myAxisY.Title.Visible = True
                            myAxisY.Label.Font = miFuenteEjes
                            myAxisY.Title.Font = miFuenteAlto
                            myAxisY.GridLines.Visible = False
                            myAxisY.Tickmarks.Visible = False
                            myAxisY.Tickmarks.MinorVisible = False
                            myAxisY.Title.TextColor = Color.Green
                            myAxisY.Label.TextColor = Color.Green
                            myAxisY.Color = Color.Green
                        End If
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.GridLines.Visible = False
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.Label.Font = miFuenteEjes
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.Title.Text = Strings.Space(5) & config.Tables(0).Rows(0)!texto_x & Strings.Space(10)
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.GridLines.Visible = False
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.Title.Font = miFuenteAlto
                        CType(ChartControl1.Diagram, XYDiagram).AxisX.Title.Visibility = DevExpress.Utils.DefaultBoolean.True
                        If config.Tables(0).Rows(0)!overlap = "R" Then
                            CType(ChartControl1.Diagram, XYDiagram).AxisX.Label.ResolveOverlappingOptions.AllowStagger = False
                            CType(ChartControl1.Diagram, XYDiagram).AxisX.Label.ResolveOverlappingOptions.AllowRotate = True
                        Else
                            CType(ChartControl1.Diagram, XYDiagram).AxisX.Label.ResolveOverlappingOptions.AllowStagger = True
                            CType(ChartControl1.Diagram, XYDiagram).AxisX.Label.ResolveOverlappingOptions.AllowRotate = False
                        End If


                        ChartControl1.Titles.Add(Titulo)
                        Dim Titulo2 As New ChartTitle()

                        Titulo2.Font = miFuente
                        Titulo2.Text = traduccion(56) & cadPeriodo
                        ChartControl1.Titles.Add(Titulo2)
                        Dim Titulo3 As New ChartTitle()
                        Titulo3.Font = miFuente
                        Titulo3.Text = traduccion(57) & Format(Now, "ddd dd-MMM-yyyy HH:mm:ss")
                        ChartControl1.Titles.Add(Titulo3)
                        Dim Titulo4 As New ChartTitle()
                        Titulo4.Font = miFuente
                        Titulo4.Text = traduccion(58) & Format(eDesde, "dd-MMM-yyyy HH:mm:ss") & traduccion(59) &
                                                            Format(eHasta, "dd-MMM-yyyy HH:mm:ss")
                        ChartControl1.Titles.Add(Titulo4)
                        ChartControl1.Width = 1000
                        ChartControl1.Height = 700

                        If config.Tables(0).Rows(0)!ver_leyenda = "S" Then
                            ChartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True
                        Else
                            ChartControl1.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False
                        End If
                        Try
                            Dim rutaImagen = Microsoft.VisualBasic.Strings.Replace(archivoImagen, "\", "\\")
                            SaveChartImageToFile(ChartControl1, ImageFormat.Png, rutaImagen)
                            Dim image As Image = GetChartImage(ChartControl1, ImageFormat.Png)
                            image.Save(rutaImagen)

                        Catch ex As Exception
                            agregarLOG(traduccion(60) & ex.Message, 7, 0)
                        End Try
                    End If
                End If
            End If
        End If
    End Function


    Function calcularPromedio(tiempo As Integer) As String
        calcularPromedio = ""
        tiempo = Math.Round(tiempo, 0)
        Dim horas = tiempo / 3600
        Dim minutos = (tiempo Mod 3600) / 60
        Dim segundos = tiempo Mod 60
        If segundos > 30 Then
            minutos = minutos + 1
        End If
        If minutos = 0 And horas = 0 Then
            minutos = 1
        End If
        calcularPromedio = Format(Math.Floor(horas), "00") & ":" & Format(Math.Floor(minutos), "00")
    End Function

    Private Sub agregarLOG(cadena As String, Optional reporte As Integer = 0, Optional tipo As Integer = 0, Optional aplicacion As Integer = 80)
        If Not be_log_activar Then Exit Sub
        'tipo 0: Info
        'tipo 2: Advertencia
        'tipo 9: Error
        Dim regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".log (aplicacion, tipo, proceso, texto) VALUES (" & aplicacion & ", " & tipo & ", " & reporte & ", '" & Microsoft.VisualBasic.Strings.Left(cadena, 250) & "')")
    End Sub

    Private Function GetChartImage(ByVal chart As ChartControl, ByVal format As ImageFormat) As Image
        ' Create an image.  
        Dim image As Image = Nothing

        ' Create an image of the chart.  
        Using s As New MemoryStream()
            chart.ExportToImage(s, format)
            image = System.Drawing.Image.FromStream(s)
        End Using

        ' Return the image.  
        Return image
    End Function

    Private Function GetGaugeImage(ByVal chart As GaugeControl, ByVal format As ImageFormat) As Image
        ' Create an image.  
        Dim image As Image = Nothing

        ' Create an image of the chart.  
        Using s As New MemoryStream()
            chart.ExportToImage(s, format)
            image = System.Drawing.Image.FromStream(s)
        End Using

        ' Return the image.  
        Return image
    End Function

    Private Sub SaveChartImageToFile(ByVal chart As ChartControl, ByVal format As ImageFormat, ByVal fileName As String)
        ' Create an image in the specified format from the chart  
        ' and save it to the specified path.  
        chart.ExportToImage(fileName, format)
    End Sub

    Private Sub SaveGaugeImageToFile(ByVal chart As GaugeControl, ByVal format As ImageFormat, ByVal fileName As String)
        ' Create an image in the specified format from the chart  
        ' and save it to the specified path.  
        chart.ExportToImage(fileName, format)
    End Sub

    Sub etiquetas()
        Dim general = consultaSEL("SELECT cadena FROM " & rutaBD & ".det_idiomas_back WHERE idioma = " & IIf(be_idioma = 0, 1, be_idioma) & " AND modulo = 5 ORDER BY linea")
        Dim cadenaTrad = ""
        If general.Tables(0).Rows.Count > 0 Then
            For Each cadena In general.Tables(0).Rows
                cadenaTrad = cadenaTrad & cadena!cadena
            Next
        End If
        traduccion = cadenaTrad.Split(New Char() {";"c})
        Label1.Text = traduccion(0)
        cad_consolidado = traduccion(1)
    End Sub


    Sub generarFIltro(consulta)
        Dim general = consultaSEL("SELECT * FROM " & rutaBD & ".consultas_cab WHERE id = " & consulta)
        Dim cadenaTrad = ""
        If general.Tables(0).Rows.Count > 0 Then

            If general.Tables(0).Rows(0)!filtromaq = "N" Then

                filtroOEE = filtroOEE & " AND i.equipo IN (SELECT valor FROM " & rutaBD & ".consultas_det WHERE consulta = " & consulta & " AND tabla = 20) "

            End If
            If general.Tables(0).Rows(0)!filtrotec = "N" Then

                filtroOEE = filtroOEE & " AND i.operador IN (SELECT valor FROM " & rutaBD & ".consultas_det WHERE consulta = " & consulta & " AND tabla = 50) "
            End If
            If general.Tables(0).Rows(0)!filtronpar = "N" Then

                filtroOEE = filtroOEE & " AND i.parte IN (SELECT valor FROM " & rutaBD & ".consultas_det WHERE consulta = " & consulta & " AND tabla = 60) "
            End If
            If general.Tables(0).Rows(0)!filtrotur = "N" Then

                filtroOEE = filtroOEE & " AND i.turno IN (SELECT valor FROM " & rutaBD & ".consultas_det WHERE consulta = " & consulta & " AND tabla = 70) "
            End If
            If general.Tables(0).Rows(0)!filtroord = "N" Then

                filtroOEE = filtroOEE & " AND i.orden IN (SELECT valor FROM " & rutaBD & ".consultas_det WHERE consulta = " & consulta & " AND tabla = 80) "
            End If
        End If

    End Sub

End Class

