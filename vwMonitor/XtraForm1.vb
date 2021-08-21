Imports DevExpress.XtraEditors
Imports DevExpress.Skins
Imports MySql.Data.MySqlClient
Imports System.Speech.Synthesis
Imports System.IO.Ports
Imports System.IO
Imports System.Text
Imports System.Net.Mail
Imports System.Net.Http
Imports System.Net
Imports System.ComponentModel



Public Class XtraForm1
    Dim Estado As Integer = 0
    Dim leyendoLog As Boolean = False
    Dim enMonitor As Boolean = False
    Dim procesandoAudios As Boolean = False
    Dim reenviar As Boolean = False
    Dim procesandoEscalamientos As Boolean
    Dim procesandoRepeticiones As Boolean
    Dim estadoPrograma As Boolean
    Dim eSegundos = 0
    Dim MensajeLlamada = ""
    Dim procesandoMensajes As Boolean = False
    Dim enviandoReportes As Boolean = False
    Dim errorCorreos As String
    Dim depurando As Boolean = False, primerSensor As Boolean = True
    Dim revisandoSensores As Boolean = False
    Dim incluyeHoyos = False
    Dim be_log_lineas As Integer
    Dim be_idioma As Integer
    Dim be_log_activar As Boolean = False
    Dim be_turno_actual As Long = 0
    Dim be_audios_activar As Boolean = False
    Dim be_tiempo_audios As Long = 0
    Dim be_hibrido_alarmar_ubicacion As Boolean = False
    Dim be_hibrido_alarmar_reparacion As Boolean = False
    Dim be_audios_escalamiento As Integer = 0
    Dim area_change = 0
    Dim tipo_change = 0
    Dim entroReportes As Boolean = False
    Dim idProceso
    Dim dosBotProcesando As Boolean = False
    Dim cincoBotProcesando As Boolean = False
    Dim fallasPLCProcesando As Boolean = False
    Dim botCheckList As Boolean = False
    Dim tDisponible As Integer = 0
    Dim tMantto As Integer = 0
    Dim modulos(20) As Integer
    Dim tipoANDON As Integer = 0
    Dim contarTotal = 0
    Dim operador As Long = 0

    Private Sub XtraForm1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim argumentos As String() = Environment.GetCommandLineArgs()
        If Process.GetProcessesByName _
          (Process.GetCurrentProcess.ProcessName).Length > 1 Then

            Dim cadSQL As String = "SELECT idioma_defecto FROM " & rutaBD & ".configuracion"
            Dim reader As DataSet = consultaSEL(cadSQL)
            If errorBD.Length > 0 Then
                agregarLOG("No connection with MySQL. Error: " & errorBD, 9, 0)
                Application.Exit()
            Else
                be_idioma = ValNull(reader.Tables(0).Rows(0)!idioma_defecto, "N")
                etiquetas()
            End If
            XtraMessageBox.Show(traduccion(0), traduccion(1), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
        Else
            idProceso = Process.GetCurrentProcess.Id
            If argumentos.Length > 1 Then
                cadenaConexion = argumentos(1).ToUpper
            End If
            If argumentos.Length > 2 Then
                cadenaConexionMMCALL = argumentos(2).ToUpper
            End If
            If argumentos.Length > 3 Then
                cadenaConexionFALLAS = argumentos(3).ToUpper
            End If

            If cadenaConexion = "" Then
                cadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
            Else
                Dim baseCadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
                Dim arreParametros = baseCadenaConexion.Split(New Char() {";"c})
                Dim arreConexion = cadenaConexion.Split(New Char() {";"c})

                If arreConexion.Length > 0 Then
                    cadenaConexion = ""
                    For i = LBound(arreParametros) To UBound(arreParametros)
                        Dim variablesValores1 = arreParametros(i).Split(New Char() {"="c})
                        Dim encontrado As Integer = -1
                        For j = LBound(arreConexion) To UBound(arreConexion)
                            Dim variablesValores2 = arreConexion(j).Split(New Char() {"="c})
                            If variablesValores1(0).ToUpper = variablesValores2(0).ToUpper Then
                                encontrado = j
                                Exit For
                            End If
                        Next
                        If encontrado = -1 Then
                            cadenaConexion = cadenaConexion & arreParametros(i).ToLower & ";"
                        Else
                            cadenaConexion = cadenaConexion & arreConexion(encontrado).ToLower & ";"
                        End If
                    Next
                End If
            End If
            If cadenaConexionMMCALL = "" Then
                cadenaConexionMMCALL = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
            Else
                Dim baseCadenaConexion = "server=127.0.0.1;user id=root;password=usbw;port=3307;Convert Zero Datetime=True;Allow User Variables=True"
                Dim arreParametros = baseCadenaConexion.Split(New Char() {";"c})
                Dim arreConexion = cadenaConexionMMCALL.Split(New Char() {";"c})

                If arreConexion.Length > 0 Then
                    cadenaConexionMMCALL = ""
                    For i = LBound(arreParametros) To UBound(arreParametros)
                        Dim variablesValores1 = arreParametros(i).Split(New Char() {"="c})
                        Dim encontrado As Integer = -1
                        For j = LBound(arreConexion) To UBound(arreConexion)
                            Dim variablesValores2 = arreConexion(j).Split(New Char() {"="c})
                            If variablesValores1(0).ToUpper = variablesValores2(0).ToUpper Then
                                encontrado = j
                                Exit For
                            End If
                        Next
                        If encontrado = -1 Then
                            cadenaConexionMMCALL = cadenaConexionMMCALL & arreParametros(i).ToLower & ";"
                        Else
                            cadenaConexionMMCALL = cadenaConexionMMCALL & arreConexion(encontrado).ToLower & ";"
                        End If
                    Next
                End If
            End If
            If cadenaConexionFALLAS = "" Then
                cadenaConexionFALLAS = "server=127.0.0.1;user id=root2;password=usbw;port=3306;Convert Zero Datetime=True;Allow User Variables=True"
            Else
                Dim baseCadenaConexion = "server=127.0.0.1;user id=root2;password=usbw;port=3306;Convert Zero Datetime=True;Allow User Variables=True"
                Dim arreParametros = baseCadenaConexion.Split(New Char() {";"c})
                Dim arreConexion = cadenaConexionFALLAS.Split(New Char() {";"c})

                If arreConexion.Length > 0 Then
                    cadenaConexionFALLAS = ""
                    For i = LBound(arreParametros) To UBound(arreParametros)
                        Dim variablesValores1 = arreParametros(i).Split(New Char() {"="c})
                        Dim encontrado As Integer = -1
                        For j = LBound(arreConexion) To UBound(arreConexion)
                            Dim variablesValores2 = arreConexion(j).Split(New Char() {"="c})
                            If variablesValores1(0).ToUpper = variablesValores2(0).ToUpper Then
                                encontrado = j
                                Exit For
                            End If
                        Next
                        If encontrado = -1 Then
                            cadenaConexionFALLAS = cadenaConexionFALLAS & arreParametros(i).ToLower & ";"
                        Else
                            cadenaConexionFALLAS = cadenaConexionFALLAS & arreConexion(encontrado).ToLower & ";"
                        End If
                    Next
                End If
            End If
            Dim cadSQL As String = "SELECT idioma_defecto FROM " & rutaBD & ".configuracion"
            Dim reader As DataSet = consultaSEL(cadSQL)
            If errorBD.Length > 0 Then
                agregarLOG("No connection with MySQL. Error: " & errorBD, 9, 0)
                Application.Exit()
            Else
                be_idioma = ValNull(reader.Tables(0).Rows(0)!idioma_defecto, "N")
                etiquetas()
            End If
            TextEdit1.Text = cadenaConexion & " | " & cadenaConexionMMCALL & " | " & cadenaConexionFALLAS
            'actualizarBD()
        End If
        For i = 0 To 19
            modulos(i) = 1
        Next
        estadoPrograma = True 'validarLicencia()

    End Sub

    Sub iniciarPantalla()
        Dim regsAfectados As Integer = 0
        ListBoxControl1.Items.Clear()

        'Se escribe en la base de datos
        regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET ejecutando_desde = NOW()")
        If errorBD.Length > 0 Then
            'Error en la base de datos
            agregarLOG(traduccion(6) & rutaBD & traduccion(29) + errorBD, 9, 0)
        ElseIf regsAfectados = 0 Then
            regsAfectados = consultaACT("INSERT INTO configuracion (ejecutando_desde, revisar_cada) VALUES ('" & Format(horaDesde, "yyyy/MM/dd HH:mm:ss") & "', 60)")
        End If
        BarManager1.Items(0).Caption = traduccion(7) + Format(horaDesde, "ddd, dd-MMM-yyyy HH:mm:ss")

    End Sub

    Private Sub XtraForm1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged
        ListBoxControl1.Width = Me.Width - 30
        TextEdit1.Width = ListBoxControl1.Width
        TextEdit1.Left = ListBoxControl1.Left
        GroupControl1.Width = ListBoxControl1.Width
        ListBoxControl1.Height = Me.Height - 250
        TextEdit1.Top = ListBoxControl1.Height + ListBoxControl1.Top + 6
        SimpleButton3.Left = Me.Width - SimpleButton3.Width - 20
        SimpleButton2.Left = Me.Width - SimpleButton2.Width - 20

    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        If XtraMessageBox.Show(traduccion(8), traduccion(9), MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.No Then
            Dim totalRegs As Integer = ListBoxControl1.Items.Count
            ListBoxControl1.Items.Clear()
            ListBoxControl1.Items.Add(Format(Now, "dd-MMM-yyyy HH:mm:ss") & ": " + traduccion(10) & totalRegs & traduccion(11) & Format(horaDesde, "dd-MMM-yyyy HH:mm:ss"))
            horaDesde = Now
            ContarLOG()
        End If
    End Sub

    Private Sub SimpleButton3_Click(sender As Object, e As EventArgs) Handles SimpleButton3.Click
        autenticado = False
        Dim Forma As New XtraForm2
        Forma.Text = traduccion(12)
        Forma.ShowDialog()
        If autenticado Then
            If XtraMessageBox.Show(traduccion(13), traduccion(12), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.No Then
                Estado = 1
                SimpleButton3.Visible = False
                SimpleButton2.Visible = True
                ContextMenuStrip1.Items(1).Enabled = False
                ContextMenuStrip1.Items(2).Enabled = True
                estadoPrograma = False
                agregarLOG(traduccion(14) & usuarioCerrar, 9, 0)
            End If
        End If
    End Sub

    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        If XtraMessageBox.Show(traduccion(15), traduccion(16), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.No Then
            Estado = 1
            SimpleButton3.Visible = True
            SimpleButton2.Visible = False
            ContextMenuStrip1.Items(1).Enabled = True
            ContextMenuStrip1.Items(2).Enabled = False
            'enviarCorreos()
            estadoPrograma = True
            agregarLOG(traduccion(17), 9, 0)
        End If
    End Sub

    Private Sub agregarLOG(cadena As String, tipo As Integer, reporte As Integer, Optional aplicacion As Integer = 1)
        If Not be_log_activar Then Exit Sub
        'Se agrega a la base de datos
        'tipo 1: Info
        'tipo 2: Incongruencia en los datos (usuario)
        'tipo 9: Error crítico de Base de datos sigma
        Dim regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".log (aplicacion, tipo, texto) VALUES (0, " & tipo & ", '" & Microsoft.VisualBasic.Strings.Left(cadena, 250) & "')")

    End Sub

    Private Sub ContarLOG()
        If be_log_lineas = 0 Then be_log_lineas = 1000
        If ListBoxControl1.Items.Count > be_log_lineas Then
            For i = ListBoxControl1.Items.Count - 1 To be_log_lineas Step -1
                ListBoxControl1.Items.RemoveAt(i)
            Next
        End If
        BarStaticItem2.Caption = IIf(ListBoxControl1.Items.Count = 0, traduccion(18), IIf(ListBoxControl1.Items.Count = 1, traduccion(19), ListBoxControl1.Items.Count & traduccion(20)))
    End Sub

    Private Sub HyperlinkLabelControl1_Click(sender As Object, e As EventArgs) Handles HyperlinkLabelControl1.Click
        System.Diagnostics.Process.Start("www.mmcallmexico.mx")
    End Sub

    Private Sub ComboBoxEdit2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBoxEdit2.SelectedIndexChanged
        Dim MiFuente As Font = New System.Drawing.Font("Lucida Sans", 9, FontStyle.Regular)

        If ComboBoxEdit2.SelectedIndex = 0 Then

            ListBoxControl1.Font = MiFuente

        ElseIf ComboBoxEdit2.SelectedIndex = 1 Then
            MiFuente = New System.Drawing.Font("Lucida Sans", 6, FontStyle.Regular)
            ListBoxControl1.Font = MiFuente
        ElseIf ComboBoxEdit2.SelectedIndex = 2 Then
            MiFuente = New System.Drawing.Font("Lucida Sans", 7, FontStyle.Regular)
            ListBoxControl1.Font = MiFuente

        ElseIf ComboBoxEdit2.SelectedIndex = 3 Then
            MiFuente = New System.Drawing.Font("Lucida Sans", 11, FontStyle.Regular)
            ListBoxControl1.Font = MiFuente
        ElseIf ComboBoxEdit2.SelectedIndex = 4 Then
            MiFuente = New System.Drawing.Font("Lucida Sans", 13, FontStyle.Regular)
            ListBoxControl1.Font = MiFuente
        ElseIf ComboBoxEdit2.SelectedIndex = 5 Then
            MiFuente = New System.Drawing.Font("Lucida Sans", 15, FontStyle.Regular)
            ListBoxControl1.Font = MiFuente
        End If
    End Sub

    Private Sub revisaFlag_Tick(sender As Object, e As EventArgs) Handles revisaFlag.Tick
        If enMonitor Or Not estadoPrograma Then Exit Sub
        horaDesde = Now
        Dim cadSQL As String = "SELECT audios_escalamiento, tiempo_audios, correo_prueba, idioma_defecto, hibrido_alarmar_ubicacion, hibrido_alarmar_reparacion, turno_oee, turno_secuencia, mapa_solicitud, ruta_programa_mapa, be_log_lineas, be_log_activar, audios_activar, area_change, tipo_change, listado01, fallas_plc FROM " & rutaBD & ".configuracion"
        Dim reader As DataSet = consultaSEL(cadSQL)
        Dim regsAfectados = 0
        Dim ruta_programa_mapa As String
        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) & errorBD, 9, 0)

        Else
            be_log_activar = ValNull(reader.Tables(0).Rows(0)!be_log_activar, "A") = "S"
            be_audios_activar = ValNull(reader.Tables(0).Rows(0)!audios_activar, "A") = "S"
            be_tiempo_audios = ValNull(reader.Tables(0).Rows(0)!tiempo_audios, "N")
            be_audios_escalamiento = ValNull(reader.Tables(0).Rows(0)!audios_escalamiento, "N")
            be_log_lineas = ValNull(reader.Tables(0).Rows(0)!be_log_lineas, "N")
            be_turno_actual = ValNull(reader.Tables(0).Rows(0)!turno_oee, "N")

            be_hibrido_alarmar_ubicacion = ValNull(reader.Tables(0).Rows(0)!hibrido_alarmar_ubicacion, "A") = "S"
            be_hibrido_alarmar_reparacion = ValNull(reader.Tables(0).Rows(0)!hibrido_alarmar_reparacion, "A") = "S"

            ruta_programa_mapa = ValNull(reader.Tables(0).Rows(0)!ruta_programa_mapa, "A")
            area_change = ValNull(reader.Tables(0).Rows(0)!area_change, "N")
            tipo_change = ValNull(reader.Tables(0).Rows(0)!tipo_change, "N")
            sensores.Enabled = modulos(5) = 1
            If be_log_lineas = 0 Then be_log_lineas = 1000
            'If ValNull(reader.Tables(0).Rows(0)!mapa_solicitud, "A") = "S" Then
            ' Try
            ' regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET mapa_solicitud = 'P'")
            ' ruta_programa_mapa = Microsoft.VisualBasic.Strings.Replace(ruta_programa_mapa, "/", "\")
            ' Shell(ruta_programa_mapa, AppWinStyle.MinimizedNoFocus)
            '
            '            Catch ex As Exception
            '            agregarSolo(traduccion(22))
            '            End Try
            '            ElseIf ValNull(reader.Tables(0).Rows(0)!mapa_solicitud, "A") = "Z" Then
            '            regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET mapa_ultimo = NOW(), mapa_solicitud = 'A'")
            '            agregarSolo(traduccion(23))
            '
            '        End If

            'If ValNull(reader.Tables(0).Rows(0)!listado01, "A") = "S" Then
            ' Try
            ' Shell(Application.StartupPath & "\listados.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "programacion" & Chr(34),' AppWinStyle.MinimizedNoFocus)
            '       agregarSolo(traduccion(182))
            ' Catch ex As Exception
            ' agregarSolo(traduccion(183))
            ' End Try
            ' End If

            If ValNull(reader.Tables(0).Rows(0)!correo_prueba, "A") = "S" Then

                Try
                    agregarSolo(traduccion(178) & " " & Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "test" & Chr(34))
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET correo_prueba = 'T'")
                    Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "test" & Chr(34), AppWinStyle.MinimizedNoFocus)
                Catch ex As Exception
                    agregarLOG(traduccion(28) & "mensajes.exe" & traduccion(29) & ex.Message, 7, 0)
                End Try



            End If
        End If

        calcularRevision()
        enMonitor = True
        'revisaFlag.Enabled = False
        'revisarEventos()
        'cancelarAlertas()
        'paseaStock()
        'asignarCarga()
        'calcularEstimado()
        depurar()
        enMonitor = False

        enviar_mensajes()

        'revisaFlag.Enabled = True

    End Sub
    Private Sub enviar_mensajes()

        If Not estadoPrograma Then Exit Sub
        Dim cadSQL = "SELECT canal FROM " & rutaBD & ".mensajes WHERE estatus = 'E' GROUP BY canal"
        Dim falla As DataSet = consultaSEL(cadSQL)
        Dim AppFuncion = ""

        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                If Not estadoPrograma Then Exit Sub
                sinEventos.Enabled = False
                sinEventos.Enabled = True
                Try
                    If lotes!canal = 0 Then
                        AppFuncion = "voz.exe"
                        Shell(Application.StartupPath & "\voz.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                        agregarSolo(traduccion(24))
                    ElseIf lotes!canal = 1 Then
                        AppFuncion = "sms.exe"
                        Shell(Application.StartupPath & "\sms.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                        agregarSolo(traduccion(25))
                    ElseIf lotes!canal = 2 Then
                        AppFuncion = "mensajes.exe"
                        Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                        agregarSolo(traduccion(177))
                    ElseIf lotes!canal = 3 Then
                        AppFuncion = "mmcall.exe"
                        Shell(Application.StartupPath & "\mmcall.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                        agregarSolo(traduccion(26))
                    ElseIf lotes!canal = 4 Then
                        AppFuncion = "log.exe"
                        Shell(Application.StartupPath & "\log.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                        agregarSolo(traduccion(27))
                    End If
                Catch ex As Exception
                    agregarLOG(traduccion(28) & AppFuncion & traduccion(29) & ex.Message, 7, 0)
                End Try

            Next
        End If

        'If be_audios_activar Then
        ' Dim cadAdic = ")"
        ' If be_audios_escalamiento = 1 Or be_audios_escalamiento Or 3 Then
        ' cadAdic = " OR (TIME_TO_SEC(TIMEDIFF(NOW(), audios)) >= " & be_tiempo_audios & "))"
        ' End If
        ' cadSQL = "SELECT id FROM " & rutaBD & ".reportes WHERE estatus = 0 AND (generar_audio <> 'Z'" & cadAdic & " LIMIT 1"
        ' falla = consultaSEL(cadSQL)
        ' If falla.Tables(0).Rows.Count > 0 Then
        ' AppFuncion = "voz.exe"
        ' Shell(Application.StartupPath & "\voz.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
        ' agregarSolo(traduccion(30))
        ' End If
        ' End If
        If Not estadoPrograma Then Exit Sub
    End Sub

    Private Sub cancelarAlertas()
        Dim regsAfectados = 0
        Dim veces = 0
        Dim pases = 0

        'Alertas ANDON
        Dim cadSQL = "SELECT a.*, c.informar_resolucion, c.evento AS tipoalerta FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".reportes b ON a.proceso = b.id AND ((b.alarmado_atender = 'S' AND b.estatus > 0) OR (b.alarmado_atendido = 'S' AND b.estatus > 10) OR (b.alarmado = 'S' AND b.estatus > 100)) WHERE a.estatus <> 9"

        Dim falla As DataSet = consultaSEL(cadSQL)
        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                Dim cadAdic = "UPDATE " & rutaBD & ".reportes SET alarmado_atender = 'Y', generar_audio = 'Z' WHERE id = " & lotes!proceso
                If lotes!tipoalerta = 102 Then
                    cadAdic = "UPDATE " & rutaBD & ".reportes SET alarmado_atendido = 'Y' WHERE id = " & lotes!proceso
                ElseIf lotes!tipoalerta = 103 Then
                    cadAdic = "UPDATE " & rutaBD & ".reportes SET alarmado = 'Y' WHERE id = " & lotes!proceso
                    'ElseIf lotes!tipoalerta = 104 Then
                    '    cadAdic = "UPDATE " & rutaBD & ".reportes SET alarmado_atencion = 'Y' WHERE id = " & lotes!proceso
                End If
                regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET estatus = 9, fin = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicio))" & IIf(ValNull(lotes!informar_resolucion, "A") = "S", ", informado = 'S'", "") & " WHERE id = " & lotes!id & ";" & cadAdic & ";UPDATE " & rutaBD & ".mensajes SET estatus = 'Z' WHERE alarma = " & lotes!id)
                If ValNull(lotes!informar_resolucion, "A") = "S" Then
                    'Se informa a los involucrados
                    regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, tipo, proceso, alarma, lista) SELECT a.alerta, b.canal, 8, a.proceso, a.id, b.lista FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".mensajes b ON a.id = b.alarma WHERE a.id = " & lotes!id & " and a.estatus = 9  GROUP BY a.alerta, b.canal, a.proceso, a.id, b.lista;")
                End If
                agregarLOG(traduccion(147) & lotes!proceso, 0, lotes!proceso)
            Next
        End If

        'Alertas OEE
        cadSQL = "SELECT a.*, b.alarmado_manual, b.alarmado_bajo, b.alarmado_alto, b.rate_efecto, b.oee_efecto, c.informar_resolucion, c.evento AS tipoalerta FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".relacion_maquinas_lecturas b ON a.proceso = b.equipo AND ((b.alarmado_bajo = 'S' AND b.rate_efecto <> 'B' AND c.evento = 201) OR (b.alarmado_alto = 'S' AND b.rate_efecto <> 'A' AND c.evento = 202) OR (b.alarmado_manual = 'N' AND c.evento = 203) OR (b.alarmado_ftq = 'S' AND b.ftq_efecto <> 'B' AND c.evento = 204) OR (b.alarmado_dis = 'S' AND b.dis_efecto <> 'B' AND c.evento = 205) OR (b.alarmado_efi = 'S' AND b.efi_efecto <> 'B' AND c.evento = 206) OR (b.alarmado_oee = 'S' AND b.oee_efecto <> 'B' AND c.evento = 207)) WHERE a.estatus <> 9"

        falla = consultaSEL(cadSQL)
        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                Dim cadAdic = ""
                Dim actualizar = False
                If lotes!tipoalerta = 201 And lotes!alarmado_bajo = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_bajo = 'N', rate_tendencia_baja = NULL WHERE equipo = " & lotes!proceso
                    actualizar = True
                ElseIf lotes!tipoalerta = 202 And lotes!alarmado_alto = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_alto = 'N', rate_tendencia_alta = NULL WHERE AND equipo = " & lotes!proceso
                    actualizar = True
                ElseIf lotes!tipoalerta = 203 And lotes!alarmado_manual = "N" Then
                    actualizar = True
                ElseIf lotes!tipoalerta = 204 And lotes!alarmado_alto = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_ftq = 'N', ftq_tendencia_baja = NULL WHERE AND equipo = " & lotes!proceso
                    actualizar = True
                ElseIf lotes!tipoalerta = 205 And lotes!alarmado_alto = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_dis = 'N', dis_tendencia_baja = NULL WHERE AND equipo = " & lotes!proceso
                    actualizar = True
                ElseIf lotes!tipoalerta = 206 And lotes!alarmado_alto = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_efi = 'N', efi_tendencia_baja = NULL WHERE AND equipo = " & lotes!proceso
                    actualizar = True
                ElseIf lotes!tipoalerta = 207 And lotes!alarmado_alto = "S" Then
                    cadAdic = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_oee = 'N', oee_tendencia_baja = NULL WHERE AND equipo = " & lotes!proceso
                    actualizar = True
                End If
                If actualizar Then
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET estatus = 9, fin = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicio))" & IIf(ValNull(lotes!informar_resolucion, "A") = "S", ", informado = 'S'", "") & " WHERE id = " & lotes!id & ";" & cadAdic & ";UPDATE " & rutaBD & ".mensajes SET estatus = 'Z' WHERE alarma = " & lotes!id)
                    If ValNull(lotes!informar_resolucion, "A") = "S" Then
                        'Se informa a los involucrados
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, tipo, proceso, alarma, lista) SELECT a.alerta, b.canal, 8, a.proceso, a.id, b.lista FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".mensajes b ON a.id = b.alarma WHERE a.id = " & lotes!id & " and a.estatus = 9  GROUP BY a.alerta, b.canal, a.proceso, a.id, b.lista;")
                    End If
                    agregarLOG(traduccion(146) & lotes!proceso, 0, lotes!proceso)
                End If

            Next
        End If

        'Alertas WIP/Lotes
        cadSQL = "SELECT a.*, c.informar_resolucion, c.evento AS tipoalerta FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".lotes b ON a.proceso = b.id AND (b.alarma_tse_paso = 'S' OR b.alarma_tpe_paso = 'S') WHERE ISNULL(a.fin) AND (c.evento = 302 OR c.evento = 303 OR c.evento = 305 OR c.evento = 306)"

        falla = consultaSEL(cadSQL)
        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                Dim cadAdic = ""
                cadAdic = "UPDATE " & rutaBD & ".lotes SET " & IIf(lotes!tipoalerta = 302 Or lotes!tipoalerta = 305, "alarma_tse_paso = 'N', alarma_tse = 'N', alarma_tse_p = 'N' ", "alarma_tpe_paso = 'N', alarma_tpe = 'N', alarma_tpe_p = 'N'") & " WHERE id = " & lotes!proceso
                regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET estatus = 9, fin = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicio))" & IIf(ValNull(lotes!informar_resolucion, "A") = "S", ", informado = 'S'", "") & " WHERE id = " & lotes!id & ";" & cadAdic & ";UPDATE " & rutaBD & ".mensajes SET estatus = 'Z' WHERE alarma = " & lotes!id)
                If ValNull(lotes!informar_resolucion, "A") = "S" Then
                    'Se informa a los involucrados
                    regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, tipo, proceso, alarma, lista) SELECT a.alerta, b.canal, 8, a.proceso, a.id, b.lista FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".mensajes b ON a.id = b.alarma WHERE a.id = " & lotes!id & " and a.estatus = 9  GROUP BY a.alerta, b.canal, a.proceso, a.id, b.lista;")
                End If
                agregarLOG(traduccion(146) & lotes!proceso, 0, lotes!proceso)

            Next
        End If

        'Alertas WIP/Cargas
        cadSQL = "SELECT a.*, c.informar_resolucion, c.transcurrido, c.evento, b.estatus AS estatuscarga, b.carga, b.fecha AS fechacarga, b.completada, b.estatus FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id LEFT JOIN " & rutaBD & ".cargas b ON a.proceso = b.id WHERE ISNULL(a.fin) AND (c.evento = 304 OR c.evento = 307)"

        falla = consultaSEL(cadSQL)
        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                Dim cadAdic = ""
                Dim mensajeTexto = ""
                If Not lotes!completada.Equals(System.DBNull.Value) Then
                    If ValNull(lotes!completada, "A") = "S" Then
                    ElseIf DateAndTime.DateAdd(DateInterval.Second, lotes!transcurrido, lotes!fechacarga) > Now() And lotes!evento = 304 Then
                        mensajeTexto = traduccion(31)
                    ElseIf ValNull(lotes!carga, "A") = "" Then
                        mensajeTexto = traduccion(32)
                    ElseIf ValNull(lotes!estatuscarga, "A") = "I" Then
                        mensajeTexto = traduccion(33)
                    End If
                    cadAdic = "UPDATE " & rutaBD & ".cargas SET alarma_rep_p = 'N', alarma_rep_paso = 'N', alarma_rep = 'N' WHERE id = " & lotes!proceso
                ElseIf ValNull(lotes!carga, "A") = "" Then
                    mensajeTexto = traduccion(32)
                End If
                If mensajeTexto.Length > 0 Then
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET estatus = 9, fin = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicio))" & IIf(ValNull(lotes!informar_resolucion, "A") = "S", ", informado = 'S'", "") & " WHERE id = " & lotes!id & ";" & cadAdic & ";UPDATE " & rutaBD & ".mensajes SET estatus = 'Z' WHERE alarma = " & lotes!id)

                    If ValNull(lotes!informar_resolucion, "A") = "S" Then

                        'Se informa a los involucrados
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, tipo, proceso, alarma, lista, texto) SELECT a.alerta, b.canal, 8, a.proceso, a.id, b.lista, '" & mensajeTexto & "' FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".mensajes b ON a.id = b.alarma WHERE a.id = " & lotes!id & " and a.estatus = 9  GROUP BY a.alerta, b.canal, a.proceso, a.id, b.lista;")
                    End If
                    agregarLOG(traduccion(146) & lotes!proceso, 0, lotes!proceso)
                End If
            Next
        End If

        crearMensajes()
    End Sub

    Function tiempoValido(desde, hasta, maquina)
        incluyeHoyos = False
        If hasta.Equals(System.DBNull.Value) Then
            tiempoValido = 0
        Else
            tiempoValido = DateAndTime.DateDiff(DateInterval.Second, desde, hasta)
        End If
        If tiempoValido = 0 Then Exit Function
        Dim diaSemana = 0
        Dim fechaEspecifica As Boolean = False
        Dim maquinaEspecifico As Boolean = False
        Dim completado = False
        Dim fechaEstimada = desde
        Dim horaDesde = Format(fechaEstimada, "HH:mm:ss")
        Dim primerDia = True
        Dim MaximoDias = 14
        Dim diasContados = 0
        tiempoValido = 0
        Do While Not completado
            diasContados = diasContados + 1
            Dim rangosPositivosD(0) As String
            Dim rangosPositivosH(0) As String
            Dim rangosNegativosD(0) As String
            Dim rangosNegativosH(0) As String
            Dim rangoPositivo = 0
            Dim rangoNegativo = 0
            horaDesde = Format(fechaEstimada, "HH:mm:ss")
            'Recorrido por día
            diaSemana = DateAndTime.Weekday(fechaEstimada) - 1
            Dim cadSQL = "SELECT desde, hasta, dia, maquina, tipo FROM " & rutaBD & ".horarios WHERE clase = 0 AND (dia = " & diaSemana & " OR (dia = 9 AND fecha = '" & Format(fechaEstimada, "yyyy/MM/dd") & "')) AND (maquina = 0 OR maquina = " & maquina & ") AND hasta > '" & horaDesde & "' ORDER BY tipo DESC, maquina DESC, dia DESC, desde, hasta"
            Dim horarios As DataSet = consultaSEL(cadSQL)
            maquinaEspecifico = False
            fechaEspecifica = False
            Dim segundos = 0
            Dim primerRegistro = True
            Dim holgura = 0
            Dim combinacion = 0
            Dim sumando As Boolean = True
            Dim continuar = False
            If horarios.Tables(0).Rows.Count > 0 Then
                For Each rango In horarios.Tables(0).Rows

                    If primerRegistro Then
                        primerRegistro = False
                        If rango!tipo = "S" Then
                            rangoPositivo = rangoPositivo + 1
                            If fechaEstimada.date = desde.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                Else
                                    rangosPositivosD(rangoPositivo - 1) = horaDesde
                                End If
                            Else
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                            End If
                            rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            maquinaEspecifico = rango!maquina <> 0
                            fechaEspecifica = rango!dia = 9
                            If maquinaEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf maquinaEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        End If
                    Else
                        If sumando And rango!tipo = "N" Then
                            sumando = False
                            rangoNegativo = rangoNegativo + 1
                            If fechaEstimada.date = desde.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                Else
                                    rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                    rangosNegativosH(rangoNegativo - 1) = horaDesde
                                End If
                            Else
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                            End If
                            rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            maquinaEspecifico = rango!maquina <> 0
                            fechaEspecifica = rango!dia = 9
                            If maquinaEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf maquinaEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        Else
                            If combinacion = 1 Then
                                continuar = rango!maquina <> 0 And rango!dia <> 9
                            ElseIf combinacion = 2 Then
                                continuar = rango!maquina <> 0
                            ElseIf combinacion = 3 Then
                                continuar = rango!dia = 9
                            Else
                                continuar = True
                            End If
                        End If
                        If continuar Then
                            If sumando Then
                                rangoPositivo = rangoPositivo + 1
                                ReDim Preserve rangosPositivosD(rangoPositivo)
                                ReDim Preserve rangosPositivosH(rangoPositivo)
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            Else
                                rangoNegativo = rangoNegativo + 1
                                ReDim Preserve rangosNegativosD(rangoNegativo)
                                ReDim Preserve rangosNegativosH(rangoNegativo)
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            End If
                        End If
                    End If
                Next
            End If
            'Se crear un registro único por día
            If rangoPositivo > 0 Then
                Dim arreDefD(0) As String
                Dim arreDefH(0) As String
                Dim arreDefP(0) As String
                Dim totalItems = 1
                arreDefD(0) = rangosPositivosD(0)
                arreDefH(0) = rangosPositivosH(0)
                arreDefP(0) = "S"
                For i = 0 To rangoPositivo - 1
                    If rangosPositivosD(i) > arreDefH(totalItems - 1) Then
                        totalItems = totalItems + 1
                        ReDim Preserve arreDefD(totalItems)
                        ReDim Preserve arreDefH(totalItems)
                        ReDim Preserve arreDefP(totalItems)
                        arreDefD(totalItems - 1) = rangosPositivosD(i)
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                        arreDefP(totalItems - 1) = "S"
                    ElseIf rangosPositivosH(i) > arreDefH(totalItems - 1) Then
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                    End If
                Next
                If rangoNegativo > 0 Then
                    For i = 0 To rangoNegativo - 1
                        For j = 0 To totalItems - 1
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefP(j) = "N"
                            End If
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefD(j) Then
                                arreDefD(j) = rangosNegativosH(i)
                            End If
                            If rangosNegativosD(i) >= arreDefD(j) And rangosNegativosD(i) < arreDefH(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                            If rangosNegativosD(i) > arreDefD(j) And rangosNegativosH(i) < arreDefH(j) Then
                                totalItems = totalItems + 1
                                ReDim Preserve arreDefD(totalItems)
                                ReDim Preserve arreDefH(totalItems)
                                ReDim Preserve arreDefP(totalItems)
                                arreDefD(totalItems - 1) = rangosNegativosH(i)
                                arreDefH(totalItems - 1) = arreDefH(j)
                                arreDefP(totalItems - 1) = "S"
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                        Next
                    Next
                End If
                If totalItems > 0 Then
                    Dim swap1 = 0
                    Dim swap2 = 0
                    Dim swap3 = 0
                    For i = 0 To totalItems - 1
                        For j = 0 To totalItems - 2
                            If arreDefD(j) > arreDefD(j + 1) Then
                                swap1 = arreDefD(j)
                                swap2 = arreDefH(j)
                                swap3 = arreDefP(j)
                                arreDefD(j) = arreDefD(j + 1)
                                arreDefH(j) = arreDefH(j + 1)
                                arreDefP(j) = arreDefP(i + 1)
                                arreDefD(j + 1) = swap1
                                arreDefH(j + 1) = swap2
                                arreDefP(j + 1) = swap3
                            End If
                        Next
                    Next

                    Dim tiempoDisponible = 0
                    For i = 0 To totalItems - 1
                        If Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)) < hasta Then
                            If arreDefP(i) = "S" Then
                                tiempoValido = tiempoValido + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)))
                            Else
                                incluyeHoyos = True
                            End If
                        ElseIf Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)) >= hasta Then
                            If arreDefP(i) = "S" Then
                                If Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)) < hasta Then
                                    tiempoValido = tiempoValido + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(hasta, "yyyy/MM/dd HH:mm:ss")))
                                End If

                            Else
                                incluyeHoyos = True
                            End If
                            completado = True
                        End If
                    Next
                    If Not completado Then
                        fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                        fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
                    End If
                End If
            Else
                If Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 23:59:59") >= hasta Then
                    completado = True
                Else
                    fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                    fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
                End If
            End If
            If diasContados > MaximoDias Then completado = True
        Loop
        'If tiempoValido = 0 Then
        ' tiempoValido = DateAndTime.DateDiff(DateInterval.Second, desde, hasta)
        ' End If
    End Function

    Function tiempoValido2(desde, hasta, maquina)
        tiempoValido2 = DateAndTime.DateDiff(DateInterval.Second, desde, hasta)
        incluyeHoyos = False
        If tiempoValido2 = 0 Then Exit Function
        Dim diaSemana = 0
        Dim fechaEspecifica As Boolean = False
        Dim maquinaEspecifico As Boolean = False
        Dim completado = False
        Dim fechaEstimada = desde
        Dim horaDesde = Format(fechaEstimada, "HH:mm:ss")
        Dim primerDia = True
        Dim MaximoDias = 14
        Dim diasContados = 0
        Do While Not completado
            diasContados = diasContados + 1
            Dim rangosPositivosD(0) As String
            Dim rangosPositivosH(0) As String
            Dim rangosNegativosD(0) As String
            Dim rangosNegativosH(0) As String
            Dim rangoPositivo = 0
            Dim rangoNegativo = 0
            horaDesde = Format(fechaEstimada, "HH:mm:ss")
            'Recorrido por día
            diaSemana = DateAndTime.Weekday(fechaEstimada) - 1
            Dim cadSQL = "SELECT desde, hasta, dia, maquina, tipo FROM " & rutaBD & ".horarios WHERE (dia = " & diaSemana & " OR (dia = 9 AND fecha = '" & Format(fechaEstimada, "yyyy/MM/dd") & "')) AND (maquina = 0 OR maquina = " & maquina & ") AND hasta > '" & horaDesde & "' ORDER BY tipo DESC, maquina DESC, dia DESC, desde, hasta"
            Dim horarios As DataSet = consultaSEL(cadSQL)
            maquinaEspecifico = False
            fechaEspecifica = False
            Dim segundos = 0
            Dim primerRegistro = True
            Dim holgura = 0
            Dim combinacion = 0
            Dim sumando As Boolean = True
            Dim continuar = False
            If horarios.Tables(0).Rows.Count > 0 Then
                For Each rango In horarios.Tables(0).Rows

                    If primerRegistro Then
                        primerRegistro = False
                        If rango!tipo = "S" Then
                            rangoPositivo = rangoPositivo + 1
                            If fechaEstimada.date = desde.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                Else
                                    rangosPositivosD(rangoPositivo - 1) = horaDesde
                                End If
                            Else
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                            End If
                            rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            maquinaEspecifico = rango!maquina <> 0
                            fechaEspecifica = rango!dia = 9
                            If maquinaEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf maquinaEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        End If
                    Else
                        If sumando And rango!tipo = "N" Then
                            sumando = False
                            rangoNegativo = rangoNegativo + 1
                            If fechaEstimada.date = desde.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                Else
                                    rangosNegativosH(rangoNegativo - 1) = horaDesde
                                End If
                            Else
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                            End If
                            rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            maquinaEspecifico = rango!maquina <> 0
                            fechaEspecifica = rango!dia = 9
                            If maquinaEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf maquinaEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        Else
                            If combinacion = 1 Then
                                continuar = rango!maquina <> 0 And rango!dia <> 9
                            ElseIf combinacion = 2 Then
                                continuar = rango!maquina <> 0
                            ElseIf combinacion = 3 Then
                                continuar = rango!dia = 9
                            Else
                                continuar = True
                            End If
                        End If
                        If continuar Then
                            If sumando Then
                                rangoPositivo = rangoPositivo + 1
                                ReDim Preserve rangosPositivosD(rangoPositivo)
                                ReDim Preserve rangosPositivosH(rangoPositivo)
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            Else
                                rangoNegativo = rangoNegativo + 1
                                ReDim Preserve rangosNegativosD(rangoNegativo)
                                ReDim Preserve rangosNegativosH(rangoNegativo)
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            End If
                        End If
                    End If
                Next
            End If
            'Se crear un registro único por día
            If rangoPositivo > 0 Then
                Dim arreDefD(0) As String
                Dim arreDefH(0) As String
                Dim arreDefP(0) As String
                Dim totalItems = 1
                arreDefD(0) = rangosPositivosD(0)
                arreDefH(0) = rangosPositivosH(0)
                arreDefP(0) = "S"
                For i = 0 To rangoPositivo - 1
                    If rangosPositivosD(i) > arreDefD(totalItems - 1) Then
                        totalItems = totalItems + 1
                        ReDim Preserve arreDefD(totalItems)
                        ReDim Preserve arreDefH(totalItems)
                        ReDim Preserve arreDefP(totalItems)
                        arreDefD(totalItems - 1) = rangosPositivosD(i)
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                        arreDefP(totalItems - 1) = "S"
                    ElseIf rangosPositivosH(i) > arreDefH(totalItems - 1) Then
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                    End If
                Next
                If rangoNegativo > 0 Then
                    For i = 0 To rangoNegativo - 1
                        For j = 0 To totalItems - 1
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefP(j) = "N"
                            End If
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefD(j) Then
                                arreDefD(j) = rangosNegativosH(i)
                            End If
                            If rangosNegativosD(i) >= arreDefD(j) And rangosNegativosD(i) < arreDefH(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                            If rangosNegativosD(i) > arreDefD(j) And rangosNegativosH(i) < arreDefH(j) Then
                                totalItems = totalItems + 1
                                ReDim Preserve arreDefD(totalItems)
                                ReDim Preserve arreDefH(totalItems)
                                ReDim Preserve arreDefP(totalItems)
                                arreDefD(totalItems - 1) = rangosNegativosH(i)
                                arreDefH(totalItems - 1) = arreDefH(j)
                                arreDefP(totalItems - 1) = "S"
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                        Next
                    Next
                End If
                If totalItems > 0 Then
                    Dim swap1 = 0
                    Dim swap2 = 0
                    Dim swap3 = 0
                    For i = 0 To totalItems - 1
                        For j = 0 To totalItems - 2
                            If arreDefD(j) > arreDefD(j + 1) Then
                                swap1 = arreDefD(j)
                                swap2 = arreDefH(j)
                                swap3 = arreDefP(j)
                                arreDefD(j) = arreDefD(j + 1)
                                arreDefH(j) = arreDefH(j + 1)
                                arreDefP(j) = arreDefP(i + 1)
                                arreDefD(j + 1) = swap1
                                arreDefH(j + 1) = swap2
                                arreDefP(j + 1) = swap3
                            End If
                        Next
                    Next


                    tiempoValido2 = 0
                    Dim tiempoDisponible = 0
                    For i = 0 To totalItems - 1
                        If Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)) <= hasta Then
                            If arreDefP(i) = "S" Then
                                tiempoValido2 = tiempoValido2 + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)))
                            End If
                        ElseIf Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)) > hasta Then
                            If arreDefP(i) = "S" Then
                                tiempoValido2 = tiempoValido2 + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(hasta, "yyyy/MM/dd HH:mm:ss")))
                            End If
                            completado = True
                        Else
                            incluyeHoyos = True
                        End If
                    Next
                    If Not completado Then
                        fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                        fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
                    End If
                End If
            Else
                completado = True
                fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
            End If
            If diasContados > MaximoDias Then completado = True
        Loop
    End Function


    Function calcularFechaEstimada(fecha, tiempoNecesario, proceso)
        calcularFechaEstimada = fecha
        If tiempoNecesario = 0 Then Exit Function
        Dim diaSemana = 0
        Dim fechaEspecifica As Boolean = False
        Dim procesoEspecifico As Boolean = False
        Dim completado = False
        Dim fechaEstimada = fecha
        Dim horaDesde = Format(fechaEstimada, "HH:mm:ss")
        Dim primerDia = True
        Dim MaximoDias = 14
        Dim diasContados = 0
        Do While Not completado
            diasContados = diasContados + 1
            Dim rangosPositivosD(0) As String
            Dim rangosPositivosH(0) As String
            Dim rangosNegativosD(0) As String
            Dim rangosNegativosH(0) As String
            Dim rangoPositivo = 0
            Dim rangoNegativo = 0
            horaDesde = Format(fechaEstimada, "HH:mm:ss")
            'Recorrido por día
            diaSemana = DateAndTime.Weekday(fechaEstimada) - 1
            Dim cadSQL = "SELECT desde, hasta, dia, proceso, tipo FROM " & rutaBD & ".horarios WHERE (dia = " & diaSemana & " OR (dia = 9 AND fecha = '" & Format(fechaEstimada, "yyyy/MM/dd") & "')) AND (proceso = 0 OR proceso = " & proceso & ") AND hasta > '" & horaDesde & "' ORDER BY tipo DESC, proceso DESC, dia DESC, desde, hasta"
            Dim horarios As DataSet = consultaSEL(cadSQL)
            procesoEspecifico = False
            fechaEspecifica = False
            Dim segundos = 0
            Dim primerRegistro = True
            Dim holgura = 0
            Dim combinacion = 0
            Dim sumando As Boolean = True
            Dim continuar = False
            If horarios.Tables(0).Rows.Count > 0 Then
                For Each rango In horarios.Tables(0).Rows

                    If primerRegistro Then
                        primerRegistro = False
                        If rango!tipo = "S" Then
                            rangoPositivo = rangoPositivo + 1
                            If fechaEstimada.date = fecha.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                Else
                                    rangosPositivosD(rangoPositivo - 1) = horaDesde
                                End If
                            Else
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                            End If
                            rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            procesoEspecifico = rango!proceso <> 0
                            fechaEspecifica = rango!dia = 9
                            If procesoEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf procesoEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        End If
                    Else
                        If sumando And rango!tipo = "N" Then
                            sumando = False
                            rangoNegativo = rangoNegativo + 1
                            If fechaEstimada.date = fecha.date Then
                                If rango!desde.ToString > horaDesde Then
                                    rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                Else
                                    rangosNegativosH(rangoNegativo - 1) = horaDesde
                                End If
                            Else
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                            End If
                            rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            procesoEspecifico = rango!proceso <> 0
                            fechaEspecifica = rango!dia = 9
                            If procesoEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf procesoEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        Else
                            If combinacion = 1 Then
                                continuar = rango!proceso <> 0 And rango!dia <> 9
                            ElseIf combinacion = 2 Then
                                continuar = rango!proceso <> 0
                            ElseIf combinacion = 3 Then
                                continuar = rango!dia = 9
                            Else
                                continuar = True
                            End If
                        End If
                        If continuar Then
                            If sumando Then
                                rangoPositivo = rangoPositivo + 1
                                ReDim Preserve rangosPositivosD(rangoPositivo)
                                ReDim Preserve rangosPositivosH(rangoPositivo)
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            Else
                                rangoNegativo = rangoNegativo + 1
                                ReDim Preserve rangosNegativosD(rangoNegativo)
                                ReDim Preserve rangosNegativosH(rangoNegativo)
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString
                            End If
                        End If
                    End If
                Next
            End If
            'Se crear un registro único por día
            If rangoPositivo > 0 Then
                Dim arreDefD(0) As String
                Dim arreDefH(0) As String
                Dim arreDefP(0) As String
                Dim totalItems = 1
                arreDefD(0) = rangosPositivosD(0)
                arreDefH(0) = rangosPositivosH(0)
                arreDefP(0) = "S"
                For i = 0 To rangoPositivo - 1
                    If rangosPositivosD(i) > arreDefD(totalItems - 1) Then
                        totalItems = totalItems + 1
                        ReDim Preserve arreDefD(totalItems)
                        ReDim Preserve arreDefH(totalItems)
                        ReDim Preserve arreDefP(totalItems)
                        arreDefD(totalItems - 1) = rangosPositivosD(i)
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                        arreDefP(totalItems - 1) = "S"
                    ElseIf rangosPositivosH(i) > arreDefH(totalItems - 1) Then
                        arreDefH(totalItems - 1) = rangosPositivosH(i)
                    End If
                Next
                If rangoNegativo > 0 Then
                    For i = 0 To rangoNegativo - 1
                        For j = 0 To totalItems - 1
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefP(j) = "N"
                            End If
                            If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefD(j) Then
                                arreDefD(j) = rangosNegativosH(i)
                            End If
                            If rangosNegativosD(i) >= arreDefD(j) And rangosNegativosD(i) < arreDefH(j) And rangosNegativosH(i) >= arreDefH(j) Then
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                            If rangosNegativosD(i) > arreDefD(j) And rangosNegativosH(i) < arreDefH(j) Then
                                totalItems = totalItems + 1
                                ReDim Preserve arreDefD(totalItems)
                                ReDim Preserve arreDefH(totalItems)
                                ReDim Preserve arreDefP(totalItems)
                                arreDefD(totalItems - 1) = rangosNegativosH(i)
                                arreDefH(totalItems - 1) = arreDefH(j)
                                arreDefP(totalItems - 1) = "S"
                                arreDefH(j) = rangosNegativosD(i)
                            End If
                        Next
                    Next
                End If
                If totalItems > 0 Then
                    Dim swap1 = 0
                    Dim swap2 = 0
                    Dim swap3 = 0
                    For i = 0 To totalItems - 1
                        For j = 0 To totalItems - 2
                            If arreDefD(j) > arreDefD(j + 1) Then
                                swap1 = arreDefD(j)
                                swap2 = arreDefH(j)
                                swap3 = arreDefP(j)
                                arreDefD(j) = arreDefD(j + 1)
                                arreDefH(j) = arreDefH(j + 1)
                                arreDefP(j) = arreDefP(i + 1)
                                arreDefD(j + 1) = swap1
                                arreDefH(j + 1) = swap2
                                arreDefP(j + 1) = swap3
                            End If
                        Next
                    Next


                    Dim tiempoDisponible = 0
                    For i = 0 To totalItems - 1
                        If arreDefP(i) = "S" Then
                            horaDesde = arreDefD(i)
                            tiempoDisponible = DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i)))
                            If tiempoDisponible > tiempoNecesario Then
                                'Se cubrió
                                completado = True

                                fechaEstimada = DateAdd(DateInterval.Second, tiempoNecesario, fechaEstimada)
                                tiempoNecesario = 0
                            Else
                                tiempoNecesario = tiempoNecesario - tiempoDisponible
                                fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " " & arreDefH(i))
                            End If
                        End If
                    Next
                    If Not completado Then
                        fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                        fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
                    End If
                End If
            Else
                fechaEstimada = DateAdd(DateInterval.Day, 1, fechaEstimada)
                fechaEstimada = Convert.ToDateTime(Format(fechaEstimada, "yyyy/MM/dd") & " 00:00:00")
            End If
            If diasContados > MaximoDias Then completado = True
        Loop
        If tiempoNecesario > 0 Then
            'Es la misma fecha, se debe sumar obligatoriamente
            fechaEstimada = DateAdd(DateInterval.Second, tiempoNecesario, fechaEstimada)
        End If
        calcularFechaEstimada = fechaEstimada

    End Function

    Private Sub revisarEventos()
        'Se alarman los reportes cuyo de llenado ya pasó
        'Dim regsAfectados = consultaACT("UPDATE " & rutabd & ".reportes a, " & rutabd & ".configuracion b SET a.alarmado = 'S' WHERE TIME_TO_SEC(TIMEDIFF(NOW(), a.inicio_reporte)) >= b.tiempo_reporte AND b.tiempo_reporte > 0 AND a.estatus = 100")
        If Not estadoPrograma Then Exit Sub

        Dim cadSQL = "SELECT * FROM " & rutaBD & ".int_eventos WHERE monitor = 'S' AND (ISNULL(revisado) OR TIME_TO_SEC(TIMEDIFF(NOW(), revisado)) >= revision) ORDER BY prioridad"
        Dim eventos As DataSet = consultaSEL(cadSQL)
        If eventos.Tables(0).Rows.Count > 0 Then
            For Each evento In eventos.Tables(0).Rows
                If Not estadoPrograma Then Exit Sub
                If evento!alerta >= 101 And evento!alerta <= 307 Then
                    'Alertas referenciadas a la tabla reportes
                    alertarReporte(evento!alerta)
                    Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".int_eventos SET revisado = NOW() WHERE alerta = " & evento!alerta)
                End If
            Next
            crearMensajes()
        End If

    End Sub


    Sub alertarReporte(evento As Integer)
        Dim regsAfectados = 0
        Dim veces = 0
        Dim pases = 0
        Dim cadSQL As String = ""
        Dim filtroANDON = " AND a.origen = 0 "
        If be_hibrido_alarmar_ubicacion Then
            filtroANDON = ""
        End If
        Dim filtroANDONRep = " AND a.origen = 0 "
        If be_hibrido_alarmar_ubicacion Then
            filtroANDONRep = ""
        End If
        If evento = 101 Then
            cadSQL = "SELECT a.id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR a.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.maquina IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND (z.area = 'S' OR a.area IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 3 AND alerta = z.id)) AND (z.falla = 'S' OR a.falla IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 4 AND alerta = z.id)) AND evento = 101 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.fecha)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".reportes a WHERE a.alarmado_atender = 'N' AND a.estatus = 0 " & filtroANDON & " HAVING idalerta > 0"
        ElseIf evento = 102 Then
            cadSQL = "SELECT a.id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR a.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.maquina IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND (z.area = 'S' OR a.area IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 3 AND alerta = z.id)) AND (z.falla = 'S' OR a.falla IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 4 AND alerta = z.id)) AND evento = 102 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.inicio_atencion)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".reportes a WHERE a.alarmado_atendido = 'N' AND a.estatus = 10 " & filtroANDONRep & " HAVING idalerta > 0"
        ElseIf evento = 103 Then
            cadSQL = "Select a.id, 0 As lote, (Select id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR a.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.maquina IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND (z.area = 'S' OR a.area IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 3 AND alerta = z.id)) AND (z.falla = 'S' OR a.falla IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 4 AND alerta = z.id)) AND evento = 103 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.inicio_reporte)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".reportes a LEFT JOIN " & rutaBD & ".cat_lineas b ON a.linea = b.id LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.maquina = c.id LEFT JOIN " & rutaBD & ".cat_areas d ON a.area = d.id LEFT JOIN " & rutaBD & ".cat_fallas e ON a.falla = e.id WHERE a.alarmado = 'N' AND a.estatus = 100 " & filtroANDONRep & " HAVING idalerta > 0"
        ElseIf evento = 201 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 201 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.rate_tendencia_baja)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_bajo = 'N' HAVING idalerta > 0"
        ElseIf evento = 202 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 202 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.rate_tendencia_alta)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_alto = 'N' HAVING idalerta > 0"
        ElseIf evento = 203 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 203 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.ultima_produccion)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_manual = 'N' HAVING idalerta > 0"
        ElseIf evento = 204 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 204 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.ftq_tendencia_baja)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_ftq = 'N' HAVING idalerta > 0"
        ElseIf evento = 205 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 205 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.dis_tendencia_baja)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_dis = 'N' HAVING idalerta > 0"
        ElseIf evento = 206 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 206 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.efi_tendencia_baja)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_efi = 'N' HAVING idalerta > 0"
        ElseIf evento = 207 Then
            cadSQL = "SELECT a.equipo AS id, 0 AS lote, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.linea = 'S' OR c.linea IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 1 AND alerta = z.id)) AND (z.maquina = 'S' OR a.equipo IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 2 AND alerta = z.id)) AND evento = 207 AND estatus = 'A' AND IFNULL(TIME_TO_SEC(TIMEDIFF(NOW(), a.oee_tendencia_baja)), 0) >= transcurrido ORDER BY prioridad DESC, linea, maquina LIMIT 1) AS idalerta FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' AND c.oee = 'S' LEFT JOIN " & rutaBD & ".cat_lineas b ON c.linea = b.id AND b.estatus = 'A' WHERE a.alarmado_oee = 'N' HAVING idalerta > 0"
        ElseIf evento = 301 Then
            cadSQL = "SELECT a.id, a.lote, a.proceso, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 301 AND estatus = 'A' ORDER BY prioridad DESC, proceso LIMIT 1) AS idalerta FROM " & rutaBD & ".lotes_historia a INNER JOIN " & rutaBD & ".lotes b ON a.lote = b.id LEFT JOIN " & rutaBD & ".det_rutas c ON a.ruta_detalle_anterior = c.id LEFT JOIN " & rutaBD & ".det_rutas d ON a.ruta_detalle = d.id LEFT JOIN " & rutaBD & ".cat_partes e ON b.parte = e.id  WHERE a.alarma_so = 'S' HAVING idalerta > 0"

        ElseIf evento = 302 Then
            cadSQL = "SELECT a.id, 0 AS lote, a.proceso, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 302 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.hasta)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".lotes a WHERE estado = 20 AND estatus = 'A' AND alarma_tse <> 'S' HAVING idalerta > 0"
        ElseIf evento = 303 Then
            cadSQL = "SELECT a.id, 0 AS lote, a.proceso, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 303 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.hasta)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".lotes a WHERE estado = 50 AND estatus = 'A' AND alarma_tpe <> 'S' HAVING idalerta > 0"
        ElseIf evento = 304 Then
            cadSQL = "SELECT a.id, a.carga, 0 AS lote, a.alarma, a.alarma_rep, IFNULL((SELECT SUM(cantidad) FROM " & rutaBD & ".programacion WHERE carga = a.id AND estatus = 'A'), 0) AS piezas, (SELECT COUNT(*) FROM " & rutaBD & ".lotes WHERE carga = a.id) AS avance, b.proceso, (SELECT id FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 304 AND estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(NOW(), a.fecha)) >= transcurrido ORDER BY prioridad DESC, linea, maquina, area, falla LIMIT 1) AS idalerta FROM " & rutaBD & ".cargas a LEFT JOIN " & rutaBD & ".cat_maquinas b ON a.equipo = b.id AND b.estatus = 'A' WHERE a.estatus = 'A'  AND a.completada = 'N'"
            'Anticipaciones

        ElseIf evento = 305 Then

            cadSQL = "SELECT a.id, 0 AS lote, a.proceso FROM " & rutaBD & ".lotes a WHERE a.estado = 20 AND estatus = 'A' AND a.alarma_tse_p <> 'S' AND TIME_TO_SEC(TIMEDIFF(a.hasta, NOW())) > 0 AND TIME_TO_SEC(TIMEDIFF(a.hasta, NOW())) <= (SELECT MAX(tiempo0) FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 305 AND estatus = 'A')"
        ElseIf evento = 306 Then

            cadSQL = "SELECT a.id, 0 AS lote, a.proceso FROM " & rutaBD & ".lotes a WHERE a.estado = 50 AND estatus = 'A' AND a.alarma_tpe_p <> 'S' AND TIME_TO_SEC(TIMEDIFF(a.hasta, NOW())) > 0 AND TIME_TO_SEC(TIMEDIFF(a.hasta, NOW())) <= (SELECT MAX(tiempo0) FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR a.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 306 AND estatus = 'A')"
        ElseIf evento = 307 Then
            cadSQL = "SELECT a.id, a.carga, 0 AS lote, a.alarma, a.alarma_rep_p, IFNULL((SELECT SUM(cantidad) FROM " & rutaBD & ".programacion WHERE carga = a.id AND estatus = 'A'), 0) AS piezas, (SELECT COUNT(*) FROM " & rutaBD & ".lotes WHERE carga = a.id) AS avance, b.proceso FROM " & rutaBD & ".cargas a LEFT JOIN " & rutaBD & ".cat_maquinas b ON a.equipo = b.id AND b.estatus = 'A' WHERE a.estatus = 'A' AND TIME_TO_SEC(TIMEDIFF(a.fecha, NOW())) > 0 AND TIME_TO_SEC(TIMEDIFF(a.fecha, NOW())) <= (SELECT MAX(tiempo0) FROM " & rutaBD & ".cat_alertas z WHERE (z.proceso = 'S' OR b.proceso IN (SELECT proceso FROM " & rutaBD & ".relacion_alertas_operaciones WHERE tipo = 5 AND alerta = z.id)) AND evento = 307 AND estatus = 'A') AND a.completada = 'N'"
        End If
        If Not estadoPrograma Then Exit Sub

        Dim falla As DataSet = consultaSEL(cadSQL)
        If falla.Tables(0).Rows.Count > 0 Then
            For Each lotes In falla.Tables(0).Rows
                If evento = 307 Then
                    If lotes!piezas > lotes!avance And lotes!alarma = "S" And lotes!alarma_rep_p <> "S" Then
                    Else
                        Continue For
                    End If
                ElseIf evento = 304 Then
                    If lotes!piezas > lotes!avance And lotes!alarma = "S" And lotes!alarma_rep <> "S" Then
                    Else
                        Continue For
                    End If
                End If
                If Not estadoPrograma Then Exit Sub
                Dim procesoID = lotes!id
                Dim loteID = lotes!lote
                If evento >= 305 And evento <= 307 Then
                    cadSQL = "SELECT * FROM " & rutaBD & ".cat_alertas WHERE evento = " & evento & " AND estatus = 'A' AND (proceso = 0 OR proceso = " & procesoID & ") ORDER BY tiempo0 DESC LIMIT 1"
                Else
                    cadSQL = "SELECT * FROM " & rutaBD & ".cat_alertas WHERE id = " & ValNull(lotes!idalerta, "N")
                End If
                Dim alerta As DataSet = consultaSEL(cadSQL)
                Dim uID = 0

                If alerta.Tables(0).Rows.Count > 0 Then

                    Dim idAlerta = alerta.Tables(0).Rows(0)!id
                    Dim fechaDesde
                    Dim crearReporte As Boolean = True
                    Dim solapable As Boolean = False

                    'Se pregunta si hay un rperte activo y si es solapable
                    If ValNull(alerta.Tables(0).Rows(0)!solapar, "A") = "S" Then
                        solapable = True
                    Else
                        cadSQL = "SELECT * FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND estatus <> 9 "
                        Dim solapar As DataSet = consultaSEL(cadSQL)
                        crearReporte = solapar.Tables(0).Rows.Count = 0
                    End If
                    If crearReporte Then
                        Dim porAcumulacion = False

                        If ValNull(alerta.Tables(0).Rows(0)!acumular, "A") = "S" Then
                            crearReporte = False
                            porAcumulacion = True
                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".alarmas (alerta, proceso, prioridad) VALUES (" & idAlerta & ", " & procesoID & ", " & alerta.Tables(0).Rows(0)!prioridad & ")")

                            If ValNull(alerta.Tables(0).Rows(0)!acumular_inicializar, "A") = "S" Then
                                If alerta.Tables(0).Rows(0)!acumular_tiempo > 0 Then
                                    fechaDesde = DateAdd(DateInterval.Second, alerta.Tables(0).Rows(0)!acumular_tiempo * -1, Now)
                                    cadSQL = "SELECT COUNT(*) as cuenta FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND estatus = 0 AND inicio >= '" & Format(fechaDesde, "yyyy/MM/dd HH:mm:ss") & "'"
                                Else
                                    cadSQL = "SELECT COUNT(*) as cuenta FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND estatus = 0"
                                End If
                            Else
                                If alerta.Tables(0).Rows(0)!acumular_tiempo > 0 Then
                                    fechaDesde = DateAdd(DateInterval.Second, alerta.Tables(0).Rows(0)!acumular_tiempo * -1, Now)
                                    cadSQL = "SELECT COUNT(*) as cuenta FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND inicio >= '" & Format(fechaDesde, "yyyy/MM/dd HH:mm:ss") & "'"
                                Else
                                    cadSQL = "SELECT COUNT(*) as cuenta FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta
                                End If
                            End If

                            Dim acumulado = 0
                            Dim acum As DataSet = consultaSEL(cadSQL)
                            If acum.Tables(0).Rows.Count > 0 Then
                                acumulado = acum.Tables(0).Rows(0)!cuenta
                            End If
                            If acumulado >= alerta.Tables(0).Rows(0)!acumular_veces Then
                                veces = acumulado + 1
                                If alerta.Tables(0).Rows(0)!acumular_tiempo > 0 Then
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET acumulada = 'S', activada = NOW(), estatus = 1 WHERE alerta = " & idAlerta & " AND estatus = 0 AND inicio >= '" & Format(fechaDesde, "yyyy/MM/dd HH:mm:ss") & "'")
                                Else
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET acumulada = 'S', activada = NOW(), estatus = 1 WHERE alerta = " & idAlerta & " AND estatus = 0")
                                End If
                                crearReporte = True
                            End If
                        Else
                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".alarmas (alerta, proceso, prioridad, estatus, activada) VALUES (" & idAlerta & ", " & procesoID & ", " & alerta.Tables(0).Rows(0)!prioridad & ", 1, NOW())")
                        End If

                        If crearReporte Then
                            If ValNull(alerta.Tables(0).Rows(0)!llamada, "A") = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) SELECT alerta, 0, proceso, " & alerta.Tables(0).Rows(0)!prioridad & ", " & alerta.Tables(0).Rows(0)!lista & ", id FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND fase = 0")
                            End If
                            If ValNull(alerta.Tables(0).Rows(0)!sms, "A") = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) Select alerta, 1, proceso, " & alerta.Tables(0).Rows(0)!prioridad & ", " & alerta.Tables(0).Rows(0)!lista & ", id FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND fase = 0")
                            End If
                            If ValNull(alerta.Tables(0).Rows(0)!correo, "A") = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) Select alerta, 2, proceso, " & alerta.Tables(0).Rows(0)!prioridad & ", " & alerta.Tables(0).Rows(0)!lista & ", id FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND fase = 0")
                            End If
                            If ValNull(alerta.Tables(0).Rows(0)!mmcall, "A") = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) Select alerta, 3, proceso, " & alerta.Tables(0).Rows(0)!prioridad & ", " & alerta.Tables(0).Rows(0)!lista & ", id FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND fase = 0")
                            End If
                            If ValNull(alerta.Tables(0).Rows(0)!log, "A") = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) Select alerta, 4, proceso, " & alerta.Tables(0).Rows(0)!prioridad & ", " & alerta.Tables(0).Rows(0)!lista & ", id FROM " & rutaBD & ".alarmas WHERE alerta = " & idAlerta & " AND fase = 0")
                            End If
                            If evento = 301 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET fase = 1, estatus = 9, fin = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicio)) WHERE estatus = 1 And fase = 0")
                            Else
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET fase = 1 WHERE alerta = " & idAlerta & " AND estatus = 1 And fase = 0")
                            End If

                            pases = pases + 1
                        End If

                        If evento = 101 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes Set alarmado_atender = 'S', alarmado_atender_desde = NOW() WHERE id = " & procesoID)
                        ElseIf evento = 102 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes SET alarmado_atendido = 'S', alarmado_atendido_desde = NOW() WHERE id = " & procesoID)
                        ElseIf evento = 103 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes SET alarmado = 'S', alarmado_informe_desde = NOW() WHERE id = " & procesoID)
                        ElseIf evento = 201 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_bajo = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 202 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_alto = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 203 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_manual = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 204 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_ftq = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 205 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_dis = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 206 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_efi = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 207 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET alarmado_oee = 'S' WHERE equipo = " & procesoID)
                        ElseIf evento = 301 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes_historia SET alarma_so = '0', alarma_so_rep = NOW() WHERE id = " & procesoID & ";UPDATE " & rutaBD & ".lotes SET alarmas = alarmas + 1 WHERE id = " & loteID)
                        ElseIf evento = 302 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET alarma_tse = 'S', alarmas = alarmas + 1 WHERE id = " & procesoID)
                        ElseIf evento = 303 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET alarma_tpe = 'S', alarmas = alarmas + 1 WHERE id = " & procesoID)
                        ElseIf evento = 304 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET alarma_rep = 'S' WHERE id = " & procesoID)
                        ElseIf evento = 305 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET alarma_tse_p = 'S' WHERE id = " & procesoID)
                        ElseIf evento = 306 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET alarma_tpe_p = 'S' WHERE id = " & procesoID)
                        ElseIf evento = 307 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET alarma_repexc_p = 'S' WHERE id = " & procesoID)
                        End If

                    Else
                        agregarLOG(traduccion(148).Replace("campo_0", procesoID).Replace("campo_1", idAlerta), 2, procesoID)
                    End If
                End If
            Next
        End If
        Dim cadAgregar = ""
        If pases > 0 And evento = 101 Then
            cadAgregar = traduccion(34)
            If pases > 1 Then
                cadAgregar = traduccion(35) & pases & traduccion(36)
            End If
        ElseIf pases > 0 And evento = 102 Then
            cadAgregar = traduccion(37)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(38)
            End If
        ElseIf pases > 0 And evento = 103 Then
            cadAgregar = traduccion(40)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(41)
            End If
        ElseIf pases > 0 And evento = 201 Then
            cadAgregar = traduccion(42)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(43)
            End If
        ElseIf pases > 0 And evento = 202 Then
            cadAgregar = traduccion(149)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(150)
            End If
        ElseIf pases > 0 And evento = 203 Then
            cadAgregar = traduccion(151)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(152)
            End If
        ElseIf pases > 0 And evento = 204 Then
            cadAgregar = traduccion(153)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(154)
            End If
        ElseIf pases > 0 And evento = 205 Then
            cadAgregar = traduccion(155)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(156)
            End If

        ElseIf pases > 0 And evento = 206 Then
            cadAgregar = traduccion(157)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(158)
            End If
        ElseIf pases > 0 And evento = 207 Then
            cadAgregar = traduccion(159)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(160)
            End If
        ElseIf pases > 0 And evento = 301 Then
            cadAgregar = traduccion(161)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(162)
            End If
        ElseIf pases > 0 And evento = 302 Then
            cadAgregar = traduccion(163)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(164)
            End If
        ElseIf pases > 0 And evento = 303 Then
            cadAgregar = traduccion(165)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(166)
            End If
        ElseIf pases > 0 And evento = 304 Then
            cadAgregar = traduccion(167)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(168)
            End If
        ElseIf pases > 0 And evento = 305 Then
            cadAgregar = traduccion(169)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(164)
            End If
        ElseIf pases > 0 And evento = 306 Then
            cadAgregar = traduccion(170)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(166)
            End If
        ElseIf pases > 0 And evento = 307 Then
            cadAgregar = traduccion(171)
            If pases > 1 Then
                cadAgregar = traduccion(39) & pases & traduccion(168)
            End If
        End If

        If pases > 0 Then agregarSolo(cadAgregar)
    End Sub

    Private Sub calcularRevision()
        Dim cadSQL As String = "SELECT revisar_cada FROM " & rutaBD & ".configuracion"
        Dim reader As DataSet = consultaSEL(cadSQL)
        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + errorBD, 9, 0)
        Else
            If reader.Tables(0).Rows.Count > 0 Then
                If ValNull(reader.Tables(0).Rows(0)!revisar_cada, "N") = 0 Then
                    eSegundos = 60
                    Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET revisar_cada = 60")
                    If revisaFlag.Interval <> 1000 Then
                        revisaFlag.Interval = 1000
                        revisaFlag.Enabled = False
                        revisaFlag.Enabled = True
                    End If

                Else
                    eSegundos = ValNull(reader.Tables(0).Rows(0)!revisar_cada, "N")
                    If revisaFlag.Interval <> eSegundos * 1000 Then
                        revisaFlag.Interval = eSegundos * 1000
                        revisaFlag.Enabled = False
                        revisaFlag.Enabled = True
                    End If
                End If

            End If
        End If
        BarManager1.Items(1).Caption = traduccion(44).Replace("campo_0", eSegundos)
    End Sub

    Function calcularTiempo(Seg) As String
        calcularTiempo = ""
        If Seg < 60 Then
            calcularTiempo = Seg & traduccion(45)
        ElseIf Seg < 3600 Then
            calcularTiempo = Math.Round(Seg / 60, 1) & traduccion(46)
        Else
            calcularTiempo = Math.Round(Seg / 3600, 1) & traduccion(47)
        End If
    End Function


    Function calcularTiempoCad(Seg) As String
        calcularTiempoCad = "-"
        Dim horas = Math.Floor(Seg / 3600)
        Dim minutos = Math.Floor((Seg Mod 3600) / 60)
        Dim segundos = (Seg Mod 3600) Mod 60
        calcularTiempoCad = horas & ":" & Format(minutos, "00") & ":" & Format(segundos, "00")
    End Function


    Sub escalamientos()
        Dim regsAfectados = 0
        Dim cadSQL = ""

        Dim limite = 0
        'Escalada 5
        cadSQL = "SELECT a.*, b.evento, b.escrep5, b.tiempo5, b.prioridad, b.lista5, b.escalar5, b.llamada5, sms5, log5, mmcall5, correo5 FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE b.escalar1 <> 'N' AND b.escalar2 <> 'N' AND b.escalar3 <> 'N' AND b.escalar4 <> 'N' AND b.escalar5 <> 'N' AND ((a.estatus = 5) OR (a.estatus >= 5 AND a.estatus < 9 AND b.repetir5 = 'S')) AND (a.escalamientos5 <= b.veces5 OR b.veces5 = 0)"
        Dim alertaDS As DataSet = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then

            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id
                'Se verifica que no se haya repetido antes
                limite = 0
                If alerta!escalada5.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada4, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = alerta!tiempo5
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada5, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = IIf(alerta!escrep5 = 0, alerta!tiempo5, alerta!escrep5)
                End If
                Dim tiempoCad = ""
                If segundos >= limite Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If ValNull(alerta!escalar5, "A") = "T" And alerta!estatus = 5 Then
                        'Se valida si se repite el mesaje para el nivel anterior
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, alarma, lista, tipo, canal, proceso, prioridad) SELECT alerta, alarma, lista, 15, canal, proceso, prioridad FROM " & rutaBD & ".mensajes WHERE alarma = " & alerta!id & " AND tipo <= 4")
                    End If
                    If alerta!estatus = 5 And alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.escalado = 5, a.generar_audio = 'E' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If
                    If ValNull(alerta!llamada5, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista5 & ", id, 5 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms5, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista5 & ", id, 5 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo5, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista5 & ", id, 5 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall5, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista5 & ", id, 5 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log5, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista5 & ", id, 5 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(174).Replace("campo_0", procesoID) & "5", 0, procesoID)
                    Dim cadAdic = "fase = 15, "
                    If alerta!escalamientos5 > 0 Then
                        cadAdic = ""
                    End If
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET " & cadAdic & "escalada5 = NOW(), escalamientos5 = escalamientos5 + 1, estatus = 6 WHERE id = " & uID)
                End If
            Next
        End If

        cadSQL = "SELECT a.*, b.evento, b.escrep4, b.tiempo4, b.prioridad, b.lista4, b.escalar4, b.llamada4, sms4, log4, mmcall4, correo4 FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE b.escalar1 <> 'N' AND b.escalar2 <> 'N' AND b.escalar3 <> 'N' AND b.escalar4 <> 'N' AND ((a.estatus = 4) OR (a.estatus >= 4 AND a.estatus < 9 AND b.repetir4 = 'S')) AND (a.escalamientos4 <= b.veces4 OR b.veces4 = 0)"
        alertaDS = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then

            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id
                'Se verifica que no se haya repetido antes
                limite = 0
                If alerta!escalada4.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada3, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = alerta!tiempo4
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada4, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = IIf(alerta!escrep4 = 0, alerta!tiempo4, alerta!escrep4)
                End If
                Dim tiempoCad = ""
                If segundos >= limite Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If ValNull(alerta!escalar4, "A") = "T" And alerta!estatus = 4 Then
                        'Se valida si se repite el mesaje para el nivel anterior
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, alarma, lista, tipo, canal, proceso, prioridad) SELECT alerta, alarma, lista, 14, canal, proceso, prioridad FROM " & rutaBD & ".mensajes WHERE alarma = " & alerta!id & " AND tipo <= 3")
                    End If
                    If alerta!estatus = 4 And alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.escalado = 4, a.generar_audio = 'E' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If
                    If ValNull(alerta!llamada4, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista4 & ", id, 4 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms4, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista4 & ", id, 4 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo4, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista4 & ", id, 4 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall4, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista4 & ", id, 4 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log4, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista4 & ", id, 4 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(174).Replace("campo_0", procesoID) & "4", 0, procesoID)
                    Dim cadAdic = "fase = 14, "
                    If alerta!escalamientos4 > 0 Then
                        cadAdic = ""
                    End If
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET " & cadAdic & "escalada4 = NOW(), escalamientos4 = escalamientos4 + 1, estatus = 5 WHERE id = " & alerta!id)
                End If
            Next
        End If

        cadSQL = "SELECT a.*, b.evento, b.escrep3, b.tiempo3, b.prioridad, b.lista3, b.escalar3, b.llamada3, sms3, log3, mmcall3, correo3 FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE b.escalar1 <> 'N' AND b.escalar2 <> 'N' AND b.escalar3 <> 'N' AND ((a.estatus = 3) OR (a.estatus >= 3 AND a.estatus < 9 AND b.repetir3 = 'S')) AND (a.escalamientos3 <= b.veces3 OR b.veces3 = 0)"
        alertaDS = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then

            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id

                limite = 0
                'Se verifica que no se haya repetido antes
                If alerta!escalada3.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada2, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = alerta!tiempo3
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada3, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = IIf(alerta!escrep3 = 0, alerta!tiempo3, alerta!escrep3)
                End If
                Dim tiempoCad = ""
                If segundos >= limite Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If ValNull(alerta!escalar3, "A") = "T" And alerta!estatus = 3 Then
                        'Se valida si se repite el mesaje para el nivel anterior
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, alarma, lista, tipo, canal, proceso, prioridad) SELECT alerta, alarma, lista, 13, canal, proceso, prioridad FROM " & rutaBD & ".mensajes WHERE alarma = " & alerta!id & " AND tipo <= 2")
                    End If
                    If alerta!estatus = 3 And alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.escalado = 3, a.generar_audio = 'E' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If
                    If ValNull(alerta!llamada3, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista3 & ", id, 3 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms3, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista3 & ", id, 3 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo3, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista3 & ", id, 3 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall3, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista3 & ", id, 3 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log3, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista3 & ", id, 3 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(174).Replace("campo_0", procesoID) & "3", 0, procesoID)
                    Dim cadAdic = "fase = 13, "
                    If alerta!escalamientos3 > 0 Then
                        cadAdic = ""
                    End If
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET " & cadAdic & "escalada3 = NOW(), escalamientos3 = escalamientos3 + 1, estatus = 4 WHERE id = " & alerta!id)
                End If
            Next
        End If

        cadSQL = "SELECT a.*, b.evento, b.escrep2, b.tiempo2, b.prioridad, b.lista2, b.escalar2, b.llamada2, sms2, log2, mmcall2, correo2 FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE b.escalar1 <> 'N' AND b.escalar2 <> 'N' AND ((a.estatus = 2) OR (a.estatus >= 2 AND a.estatus < 9 AND b.repetir2 = 'S')) AND (a.escalamientos2 <= b.veces2 OR b.veces2 = 0)"
        alertaDS = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then

            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id
                'Se verifica que no se haya repetido antes
                limite = 0
                If alerta!escalada2.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada1, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = alerta!tiempo2
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada2, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = IIf(alerta!escrep2 = 0, alerta!tiempo2, alerta!escrep2)
                End If
                Dim tiempoCad = ""
                If segundos >= limite Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If ValNull(alerta!escalar2, "A") = "T" And alerta!estatus = 2 Then
                        'Se valida si se repite el mesaje para el nivel anterior
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, alarma, lista, tipo, canal, proceso, prioridad) SELECT alerta, alarma, lista, 12, canal, proceso, prioridad FROM " & rutaBD & ".mensajes WHERE alarma = " & alerta!id & " AND tipo <= 1")
                    End If
                    If alerta!estatus = 2 And alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.escalado = 2, a.generar_audio = 'E' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If
                    If ValNull(alerta!llamada2, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista2 & ", id, 2 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms2, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista2 & ", id, 2 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo2, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista2 & ", id, 2 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall2, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista2 & ", id, 2 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log2, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista2 & ", id, 2 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(174).Replace("campo_0", procesoID) & "2", 0, procesoID)
                    Dim cadAdic = "fase = 12, "
                    If alerta!escalamientos2 > 0 Then
                        cadAdic = ""
                    End If
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET " & cadAdic & "escalada2 = NOW(), escalamientos2 = escalamientos2 + 1, estatus = 3 WHERE id = " & alerta!id)
                End If
            Next
        End If


        cadSQL = "SELECT a.*, b.evento, b.escrep1, b.tiempo1, b.prioridad, b.lista1, b.escalar1, b.llamada1, sms1, log1, mmcall1, correo1 FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE b.escalar1 <> 'N' AND ((a.estatus = 1) OR (a.estatus >= 1 AND a.estatus < 9 AND b.repetir1 = 'S')) AND (a.escalamientos1 < b.veces1 OR b.veces1 = 0)"
        alertaDS = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then

            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id
                'Se verifica que no se haya repetido antes
                limite = 0
                If alerta!escalada1.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = alerta!tiempo1
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!escalada1, "yyyy/MM/dd HH:mm:ss")), Now)
                    limite = IIf(alerta!escrep1 = 0, alerta!tiempo1, alerta!escrep1)
                End If
                Dim tiempoCad = ""
                If segundos >= limite Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If ValNull(alerta!escalar1, "A") = "T" And alerta!estatus = 1 Then
                        'Se valida si se repite el mesaje para el nivel anterior
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, alarma, lista, tipo, canal, proceso, prioridad) SELECT alerta, alarma, lista, 11, canal, proceso, prioridad FROM " & rutaBD & ".mensajes WHERE alarma = " & alerta!id & " AND tipo = 0")
                    End If
                    If alerta!estatus = 1 And alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.escalado = 1, a.generar_audio = 'E' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If
                    If ValNull(alerta!llamada1, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista1 & ", id, 1 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms1, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista1 & ", id, 1 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo1, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista1 & ", id, 1 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall1, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista1 & ", id, 1 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log1, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista1 & ", id, 1 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(174).Replace("campo_0", procesoID) & "1", 0, procesoID)
                    Dim cadAdic = "fase = 11, "
                    If alerta!escalamientos1 > 0 Then
                        cadAdic = ""
                    End If
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET " & cadAdic & "escalada1 = NOW(), escalamientos1 = escalamientos1 + 1, estatus = 2 WHERE id = " & uID)
                End If
            Next
        End If


        cadSQL = "SELECT a.*, b.evento, b.repetir_tiempo, b.repetir_veces, b.prioridad, b.lista, b.llamada, sms, log, mmcall, correo FROM " & rutaBD & ".alarmas a INNER JOIN " & rutaBD & ".cat_alertas b ON a.alerta = b.id AND b.estatus = 'A' WHERE ((a.estatus = 1 AND b.repetir = 'T') OR (a.estatus >= 1 AND a.estatus < 9 AND b.repetir = 'S')) AND b.repetir_tiempo > 0 AND (a.repeticiones < b.repetir_veces OR b.repetir_veces = 0)"
        alertaDS = consultaSEL(cadSQL)
        If alertaDS.Tables(0).Rows.Count > 0 Then
            For Each alerta In alertaDS.Tables(0).Rows
                If Not estadoPrograma Then
                    procesandoEscalamientos = False
                    Exit Sub
                End If
                Dim idAlerta = alerta!alerta
                Dim procesoID = alerta!proceso
                Dim segundos = 0
                Dim activarEscalada As Boolean = False
                Dim uID = alerta!id
                'Se verifica que no se haya repetido antes
                'Application.DoEvents()
                Dim repeticiones As Integer = alerta!repeticiones + 1
                'Se verifica que no se haya repetido antes
                If alerta!repetida.Equals(System.DBNull.Value) Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                Else
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!repetida, "yyyy/MM/dd HH:mm:ss")), Now)
                End If
                Dim tiempoCad = ""
                If segundos >= alerta!repetir_tiempo Then
                    segundos = DateDiff(DateInterval.Second, CDate(Format(alerta!activada, "yyyy/MM/dd HH:mm:ss")), Now)
                    tiempoCad = calcularTiempoCad(segundos)
                    If alerta!evento < 200 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".reportes a SET a.repeticiones = a.repeticiones + 1, a.generar_audio = 'R' WHERE estatus = 0 AND id IN (SELECT proceso FROM " & rutaBD & ".alarmas WHERE id = " & uID & ")")
                    End If

                    If ValNull(alerta!llamada, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 0, proceso, prioridad, " & alerta!lista & ", " & uID & ", 9 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!sms, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 1, proceso, prioridad, " & alerta!lista & ", " & uID & ", 9 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!correo, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 2, proceso, prioridad, " & alerta!lista & ", " & uID & ", 9 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!mmcall, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 3, proceso, prioridad, " & alerta!lista & ", " & uID & ", 9 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    If ValNull(alerta!log, "A") = "S" Then
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma, tipo) SELECT alerta, 4, proceso, prioridad, " & alerta!lista & ", " & uID & ", 9 FROM " & rutaBD & ".alarmas WHERE id = " & uID)
                    End If
                    agregarLOG(traduccion(175).Replace("campo_0", procesoID), 0, procesoID)
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".alarmas SET repeticiones = repeticiones + 1, repetida = NOW(), estatus = 1 WHERE id = " & alerta!id)
                End If
            Next
        End If


        procesandoEscalamientos = False
        BarManager1.Items(1).Caption = traduccion(44).Replace("campo_0", eSegundos)
        crearMensajes()
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Try
            Me.Visible = True
            NotifyIcon1.Visible = False
            Me.WindowState = FormWindowState.Maximized
        Catch ex As Exception

        End Try

    End Sub

    Private Sub XtraForm1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Visible = False
        NotifyIcon1.Visible = True
    End Sub


    Private Sub ReanudarElMonitorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReanudarElMonitorToolStripMenuItem.Click
        If XtraMessageBox.Show(traduccion(15), traduccion(16), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.No Then
            Estado = 1
            SimpleButton3.Visible = True
            SimpleButton2.Visible = False
            ContextMenuStrip1.Items(1).Enabled = True
            ContextMenuStrip1.Items(2).Enabled = False
            estadoPrograma = True
            agregarLOG(traduccion(17), 9, 0)
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If XtraMessageBox.Show(traduccion(13), traduccion(12), MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.No Then
            Estado = 1
            SimpleButton3.Visible = False
            SimpleButton2.Visible = True
            ContextMenuStrip1.Items(1).Enabled = False
            ContextMenuStrip1.Items(2).Enabled = True
            estadoPrograma = False
            agregarLOG(traduccion(48), 9, 0)
        End If
    End Sub

    Private Sub XtraForm1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        Dim f As Form
        f = sender

        'Check if the form is minimized
        If f.WindowState = FormWindowState.Minimized Then
            Me.Visible = False
            NotifyIcon1.Visible = True
        End If

    End Sub

    Private Sub XtraForm1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        autenticado = False
        Dim Forma As New XtraForm2
        Forma.Text = traduccion(140)
        Forma.ShowDialog()
        If autenticado Then
            If XtraMessageBox.Show(traduccion(138), traduccion(12), MessageBoxButtons.YesNo, MessageBoxIcon.Stop) <> DialogResult.No Then
                agregarLOG(traduccion(139) & usuarioCerrar, 9, 0)
            Else
                e.Cancel = True
            End If
        Else
            e.Cancel = True
        End If
    End Sub

    Sub agregarSolo(cadena As String)
        ListBoxControl1.Items.Insert(0, traduccion(49) & " " & Format(Now, "dd-MMM HH:mm:ss") & ": " & cadena)
        ContarLOG()
    End Sub

    Sub depurar()
        'Se depura la BD
        If Format(DateAndTime.Now(), "HHmm") = "0000" And Not depurando Then
            depurando = True
            'Se produce la depuración a las 12 de la noche
            Dim cadSQL As String = "SELECT gestion_meses, gestion_log FROM " & rutaBD & ".configuracion WHERE gestion_meses > 0 AND (ISNULL(gestion_log) OR gestion_log < '" & Format(Now(), "yyyyMM") & "')"
            Dim reader As DataSet = consultaSEL(cadSQL)
            Dim regsAfectados = 0
            Dim eliminados = 0
            Dim mesesAtras = 1
            If reader.Tables(0).Rows.Count > 0 Then
                mesesAtras = reader.Tables(0).Rows(0)!gestion_meses
                If mesesAtras > 0 Then
                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".reportes WHERE cierre_reporte < '" & Format(DateAndTime.DateAdd(DateInterval.Month, mesesAtras * -1, Now()), "yyyy/MM") & "/01 00:00:00' AND estatus = 1000")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".lotes WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Month, mesesAtras * -1, Now()), "yyyy/MM") & "/01 00:00:00'")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".lotes_historia WHERE lote Not IN (SELECT id FROM " & rutaBD & ".lotes)")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE From " & rutaBD & ".ruta_congelada WHERE lote Not IN (SELECT id FROM " & rutaBD & ".lotes)")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".calidad_historia WHERE lote Not IN (SELECT id FROM " & rutaBD & ".lotes)")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".detalleparos WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Month, mesesAtras * -1, Now()), "yyyy/MM") & "/01 00:00:00'")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion SET gestion_log = '" & Format(Now(), "yyyyMM") & "'")
                    eliminados = eliminados + regsAfectados

                    regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".control WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Month,
    mesesAtras * -1, Now()), "yyyyMM") & "0100'")


                    agregarLOG(traduccion(141) & Format(Now(), "MMMM-yyyy") & traduccion(142) & Format(DateAndTime.DateAdd(DateInterval.Month, mesesAtras * -1, Now()), "yyyy/MM") & "/01). " & traduccion(143) & eliminados & traduccion(144), 7, 0)

                End If
            End If
            mesesAtras = 1
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".alarmas WHERE estatus = 9 OR (fin < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -1, Now()), "yyyy/MM/dd") & " 00:00:00')")

            'Elvis
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".lecturas WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -1, Now()), "yyyy/MM/dd") & " 00:00:00' AND estatus = 2")
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".lecturas_cortes WHERE dia < '" & Format(DateAndTime.DateAdd(DateInterval.Month, -24, Now()), "yyyy/MM/dd") & "'")

            'regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".mensajes WHERE enviada < '" & Format(DateAndTime.DateAdd(DateInterval.Month, mesesAtras * -1, Now()), "yyyy/MM") & "/01 00:00:00'")
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".mensajes WHERE enviada < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -1, Now()), "yyyy/MM/dd") & " 00:00:00'")
            eliminados = eliminados + regsAfectados
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".log WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -3, Now()), "yyyy/MM/dd") & " 00:00:00'")
            eliminados = eliminados + regsAfectados
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".lecturas_resumen WHERE desde < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -2, Now()), "yyyy/MM/dd") & " 00:00:00'")
            eliminados = eliminados + regsAfectados
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".mensajes_procesados WHERE fecha < '" & Format(DateAndTime.DateAdd(DateInterval.Day, -1, Now()), "yyyy/MM/dd") & "'")
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".temporal_mmcall WHERE estatus = 1;DELETE FROM " & rutaBD & ".temporal_plc WHERE estatus = 1;")
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".horaxhora WHERE estatus = 'C';DELETE FROM " & rutaBD & ".horaxhora WHERE estatus = 'Z' AND dia < '" & Format(DateAndTime.DateAdd(DateInterval.Month, -6, Now()), "yyyy/MM/dd") & "'")


            eliminados = eliminados + regsAfectados
            depurando = False
        End If
    End Sub

    Private Sub revisarLog_Tick(sender As Object, e As EventArgs) Handles revisarLog.Tick

        'revisarLog.Enabled = False
        If leyendoLog Or Not estadoPrograma Then Exit Sub
        leyendoLog = True
        Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".log SET visto = 'P' WHERE visto = 'N'")
        Dim cadSQL = "SELECT id, texto, aplicacion FROM " & rutaBD & ".log WHERE visto = 'P' ORDER BY id"
        Dim reader = consultaSEL(cadSQL)
        If reader.Tables(0).Rows.Count > 0 Then
            sinEventos.Enabled = False
            sinEventos.Enabled = True
            For Each elmensaje In reader.Tables(0).Rows
                Dim appOrigen = traduccion(49)
                If elmensaje!aplicacion = 30 Then
                    appOrigen = traduccion(50)
                ElseIf elmensaje!aplicacion = 40 Then
                    appOrigen = traduccion(51)
                ElseIf elmensaje!aplicacion = 20 Then
                    appOrigen = traduccion(52)
                ElseIf elmensaje!aplicacion = 60 Then
                    appOrigen = traduccion(53)
                ElseIf elmensaje!aplicacion = 50 Then
                    appOrigen = traduccion(54)
                ElseIf elmensaje!aplicacion = 70 Then
                    appOrigen = traduccion(55)
                ElseIf elmensaje!aplicacion = 80 Then
                    appOrigen = traduccion(56)
                End If
                ListBoxControl1.Items.Insert(0, appOrigen & " " & Format(Now, "dd-MMM HH:mm:ss") & ": " & elmensaje!texto)

            Next
            regsAfectados = consultaACT("UPDATE " & rutaBD & ".log SET visto = 'S' WHERE visto = 'P'")
            ContarLOG()
        End If
        leyendoLog = False
        'revisarLog.Enabled = True
    End Sub

    Private Sub sinEventos_Tick(sender As Object, e As EventArgs) Handles sinEventos.Tick
        agregarSolo(traduccion(57))
    End Sub

    Private Sub escalamiento_Tick(sender As Object, e As EventArgs)

        If procesandoEscalamientos Or Not estadoPrograma Then Exit Sub
        procesandoEscalamientos = True
        escalamientos()
    End Sub

    Private Sub reportes_Tick(sender As Object, e As EventArgs) Handles reportes.Tick
        If enviandoReportes Or Not estadoPrograma Then Exit Sub
        If Format(Now, "mm") >= "05" Then
            entroReportes = False
            Exit Sub
        ElseIf Not entroReportes Then
            entroReportes = True
        Else
            Exit Sub

        End If
        'reportes.Enabled = False
        enviandoReportes = True
        Try

            Shell(Application.StartupPath & "\reportes.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
            agregarSolo(traduccion(58))
        Catch ex As Exception
            agregarLOG(traduccion(59) & ex.Message, 7, 0)
        End Try
        enviandoReportes = False
        'reportes.Enabled = True
    End Sub

    Private Sub arduino_Tick(sender As Object, e As EventArgs)
        If Not estadoPrograma Then Exit Sub

        Dim cadSQL As String = "SELECT ruta_audios, ruta_sms, be_alarmas_llamadas, be_alarmas_sms FROM " & rutaBD & ".configuracion"
        Dim reader As DataSet = consultaSEL(cadSQL)
        Dim regsAfectados = 0
        Dim rutaAudios
        Dim rutaSMS
        Dim be_alarmas_llamadas As Boolean = False
        Dim be_alarmas_sms As Boolean = False

        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + errorBD, 9, 0)

        Else
            rutaAudios = ValNull(reader.Tables(0).Rows(0)!ruta_audios, "A")
            rutaSMS = ValNull(reader.Tables(0).Rows(0)!ruta_sms, "A")
            be_alarmas_llamadas = ValNull(reader.Tables(0).Rows(0)!be_alarmas_llamadas, "A") = "S"
            be_alarmas_sms = ValNull(reader.Tables(0).Rows(0)!be_alarmas_sms, "A") = "S"
        End If
        Dim llamarPrograma As Boolean = False
        If be_alarmas_llamadas Or be_alarmas_sms Then
            If rutaSMS.Length = 0 Then
                rutaSMS = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Else
                rutaSMS = Microsoft.VisualBasic.Strings.Replace(rutaSMS, "/", "\")
            End If
            If Not My.Computer.FileSystem.DirectoryExists(rutaSMS) Then
                rutaSMS = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            If rutaAudios.Length = 0 Then
                rutaAudios = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Else
                rutaAudios = Microsoft.VisualBasic.Strings.Replace(rutaAudios, "/", "\")
            End If
            If Not My.Computer.FileSystem.DirectoryExists(rutaAudios) Then
                rutaAudios = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            If Not My.Computer.FileSystem.DirectoryExists(rutaAudios) Then
                rutaAudios = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If
            Dim LlamadasPendientes = 0
            For Each FoundFile As String In My.Computer.FileSystem.GetFiles(
          rutaAudios, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.wav")
                Dim Numero = Microsoft.VisualBasic.Strings.Left(Path.GetFileName(FoundFile), 10)
                Dim nombreArchivo = Path.GetFileName(FoundFile)
                If IsNumeric(Numero) Then LlamadasPendientes = LlamadasPendientes + 1
            Next
            If LlamadasPendientes > 0 And be_alarmas_llamadas Then
                llamarPrograma = True
            End If
            If Not llamarPrograma Then
                For Each FoundFile As String In My.Computer.FileSystem.GetFiles(
          rutaSMS, Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.txt")
                    Dim Numero = Microsoft.VisualBasic.Strings.Left(Path.GetFileName(FoundFile), 10)
                    Dim nombreArchivo = Path.GetFileName(FoundFile)
                    If IsNumeric(Numero) Then LlamadasPendientes = LlamadasPendientes + 1
                Next
                If LlamadasPendientes > 0 And be_alarmas_llamadas Then
                    llamarPrograma = True
                End If
            End If
        End If
        If llamarPrograma Then
            Try
                Shell(Application.StartupPath & "\arduino.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
                agregarSolo(traduccion(60))
            Catch ex As Exception
                agregarLOG(traduccion(61) & ex.Message, 7, 0)
            End Try
        End If
    End Sub

    Function validarLicencia() As Boolean
        validarLicencia = True

        'HEB// 
        'HEB//Exit Function
        'HEB// 

        validarLicencia = True
        Dim cadSQL = "SELECT licencia FROM " & rutaBD & ".configuracion"
        Dim falla As DataSet = consultaSEL(cadSQL)
        Dim mensajeLicencia As String = ""
        If falla.Tables(0).Rows.Count > 0 Then
            Dim licenciaSIGMA As String = ValNull(falla.Tables(0).Rows(0)!licencia, "A")
            If licenciaSIGMA.Length = 0 Then
                mensajeLicencia = traduccion(62)
            Else
                cadSQL = "SELECT CONCAT(key_number, serial) AS mmcall FROM " & rutaMMCALL & ".locations"
                falla = consultaSEL(cadSQL, cadenaConexionMMCALL)
                If falla.Tables(0).Rows.Count > 0 Then
                    Dim licenciaMMCall As String = ValNull(falla.Tables(0).Rows(0)!mmcall, "A")
                    If licenciaMMCall.Length = 0 Then
                        mensajeLicencia = traduccion(63)
                    Else
                        Dim partes = licenciaSIGMA.Split(New Char() {"-"c})
                        partes(0) = Format(CLng("&H" & partes(0)), "00000000000")
                        partes(1) = Format(CLng("&H" & partes(1)), "00000000000")
                        partes(2) = Format(CLng("&H" & partes(2)), "00000000000")
                        partes(3) = CLng("&H" & partes(3))
                        partes(3) = partes(3).ToString.PadLeft(11 - (40 - Val(Microsoft.VisualBasic.Strings.Right(partes(3), 2))), "0")
                        Dim numeroCompleto As String = partes(0) & partes(1) & partes(2) & partes(3)
                        Dim acumulador = 0
                        For i = 2 To Len(numeroCompleto) - 5
                            acumulador = acumulador + Val(Strings.Mid(numeroCompleto, i, 1))
                        Next
                        Dim totalMutiplicado = acumulador * Val(Microsoft.VisualBasic.Strings.Left(numeroCompleto, 1))
                        If totalMutiplicado >= 1000 Then
                            totalMutiplicado = Val(Microsoft.VisualBasic.Strings.Left(totalMutiplicado, 3))
                        End If
                        If totalMutiplicado <> Val(Microsoft.VisualBasic.Strings.Mid(numeroCompleto, Len(numeroCompleto) - 4, 3)) Then
                            mensajeLicencia = traduccion(64)
                        Else
                            numeroCompleto = numeroCompleto.Substring(1, Len(numeroCompleto) - 6)
                            Dim licencia As String = numeroCompleto.Substring(0, 7)
                            licencia = licencia & numeroCompleto.Substring(8, 4)
                            licencia = licencia & numeroCompleto.Substring(13, 3)
                            licencia = licencia & numeroCompleto.Substring(17, 4)
                            licencia = licencia & numeroCompleto.Substring(22, 2)
                            licencia = licencia & numeroCompleto.Substring(25, 3)
                            licencia = licencia & numeroCompleto.Substring(29, 2)
                            licencia = licencia & numeroCompleto.Substring(32, 2)
                            licencia = licencia & numeroCompleto.Substring(35)
                            Dim izquierda As String = licencia.Substring(0, 5)
                            Dim derecha As String = licencia.Substring(licencia.Length - 4, 4)
                            licencia = licencia.Substring(5, licencia.Length - 9)
                            tipoANDON = Val(izquierda.Substring(1, 1))
                            If derecha.Substring(0, 1) = "1" Then
                                modulos(2) = 1
                            ElseIf derecha.Substring(0, 1) = "2" Then
                                modulos(1) = 1
                            ElseIf derecha.Substring(0, 1) = "3" Then
                                modulos(1) = 1
                                modulos(2) = 1
                            ElseIf derecha.Substring(0, 1) = "4" Then
                                modulos(0) = 1
                            ElseIf derecha.Substring(0, 1) = "5" Then
                                modulos(0) = 1
                                modulos(2) = 1
                            ElseIf derecha.Substring(0, 1) = "6" Then
                                modulos(0) = 1
                                modulos(1) = 1
                            ElseIf derecha.Substring(0, 1) = "7" Then
                                modulos(0) = 1
                                modulos(1) = 1
                                modulos(2) = 1
                            End If
                            If derecha.Substring(1, 1) = "1" Then
                                modulos(5) = 1
                            ElseIf derecha.Substring(1, 1) = "2" Then
                                modulos(4) = 1
                            ElseIf derecha.Substring(1, 1) = "3" Then
                                modulos(4) = 1
                                modulos(5) = 1
                            ElseIf derecha.Substring(1, 1) = "4" Then
                                modulos(3) = 1
                            ElseIf derecha.Substring(1, 1) = "5" Then
                                modulos(3) = 1
                                modulos(5) = 1
                            ElseIf derecha.Substring(1, 1) = "6" Then
                                modulos(3) = 1
                                modulos(4) = 1
                            ElseIf derecha.Substring(1, 1) = "7" Then
                                modulos(3) = 1
                                modulos(4) = 1
                                modulos(5) = 1
                            End If
                            If derecha.Substring(2, 1) = "1" Then
                                modulos(8) = 1
                            ElseIf derecha.Substring(2, 1) = "2" Then
                                modulos(7) = 1
                            ElseIf derecha.Substring(2, 1) = "3" Then
                                modulos(7) = 1
                                modulos(8) = 1
                            ElseIf derecha.Substring(2, 1) = "4" Then
                                modulos(6) = 1
                            ElseIf derecha.Substring(2, 1) = "5" Then
                                modulos(6) = 1
                                modulos(8) = 1
                            ElseIf derecha.Substring(2, 1) = "6" Then
                                modulos(6) = 1
                                modulos(7) = 1
                            ElseIf derecha.Substring(2, 1) = "7" Then
                                modulos(6) = 1
                                modulos(7) = 1
                                modulos(8) = 1
                            End If
                            If derecha.Substring(3, 1) = "1" Then
                                modulos(11) = 1
                            ElseIf derecha.Substring(3, 1) = "2" Then
                                modulos(10) = 1
                            ElseIf derecha.Substring(3, 1) = "3" Then
                                modulos(10) = 1
                                modulos(11) = 1
                            ElseIf derecha.Substring(3, 1) = "4" Then
                                modulos(9) = 1
                            ElseIf derecha.Substring(3, 1) = "5" Then
                                modulos(9) = 1
                                modulos(11) = 1
                            ElseIf derecha.Substring(3, 1) = "6" Then
                                modulos(0) = 1
                                modulos(10) = 1
                            ElseIf derecha.Substring(3, 1) = "7" Then
                                modulos(9) = 1
                                modulos(10) = 1
                                modulos(11) = 1
                            End If

                            Dim claveInterna As String = "CronoEIntelraciVn201i"

                            If claveInterna.Length > licenciaMMCall.Length Then
                                licenciaMMCall = licenciaMMCall.PadLeft(claveInterna.Length - licencia.Length, "0")
                            ElseIf (licenciaMMCall.Length > claveInterna.Length) Then
                                claveInterna = claveInterna & Microsoft.VisualBasic.Strings.StrDup(licencia.Length - claveInterna.Length, "0")
                            End If
                            Dim cadComparar = ""
                            Dim licValida As Boolean = True
                            For i = 0 To licenciaMMCall.Length - 1
                                Dim numero As String = Asc(licenciaMMCall(i)) Xor Asc(claveInterna(i))
                                If (numero.Length = 1) Then
                                    cadComparar = numero
                                ElseIf (numero.Length = 2) Then
                                    cadComparar = Microsoft.VisualBasic.Strings.Mid(numero, 2, 1)
                                ElseIf (numero.Length = 3) Then
                                    cadComparar = Microsoft.VisualBasic.Strings.Mid(numero, 2, 1)
                                End If
                                If cadComparar <> Microsoft.VisualBasic.Strings.Mid(licencia, i + 1, 1) Then
                                    licValida = False
                                    Exit For
                                End If
                            Next
                            Dim anyo = numeroCompleto.Substring(7, 1) & numeroCompleto.Substring(12, 1) & numeroCompleto.Substring(16, 1) & numeroCompleto.Substring(21, 1)
                            Dim mes = numeroCompleto.Substring(24, 1) + numeroCompleto.Substring(28, 1)
                            Dim dia = numeroCompleto.Substring(31, 1) + numeroCompleto.Substring(34, 1)
                            If licValida Or anyo & mes & dia <> "99999999" Then
                                If IsDate(anyo & "/" & mes & "/" & dia) Then
                                    If Format(DateAndTime.DateValue(anyo & "/" & mes & "/" & dia), "yyyyMMdd") < Format(Now.Date, "yyyyMMdd") Then
                                        mensajeLicencia = traduccion(65)
                                    Else
                                        'LabelControl2.Text = LabelControl2.Text & traduccion(66) & Format(DateAndTime.DateValue(anyo & "/" & mes & "/" & dia), "ddd, dd/MMM/yyyy")
                                    End If

                                End If
                            ElseIf Not licValida Then
                                mensajeLicencia = traduccion(67)
                            End If
                        End If
                    End If
                Else
                    mensajeLicencia = traduccion(68)
                End If
            End If
        Else
            mensajeLicencia = traduccion(69)
        End If
        If mensajeLicencia.Length > 0 Then
            XtraMessageBox.Show(mensajeLicencia, traduccion(70), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            estadoPrograma = False
            validarLicencia = False
        End If
    End Function

    Private Sub reenviarMMCALL_Tick(sender As Object, e As EventArgs)

        If reenviar Or Not estadoPrograma Then Exit Sub
        reenviar = True
        If Not estadoPrograma Then Exit Sub
        Try

            Shell(Application.StartupPath & "\repeticiones.exe " & Chr(34) & cadenaConexion & Chr(34), AppWinStyle.MinimizedNoFocus)
        Catch ex As Exception
            agregarLOG(traduccion(71) & ex.Message, 7, 0)
        End Try

        reenviar = False
    End Sub

    Function validarURI(ByVal cadena As String) As Boolean
        Dim validatedUri As System.Uri
        Return Uri.TryCreate(cadena, UriKind.RelativeOrAbsolute, validatedUri)
    End Function

    Private Sub sensores_Tick(sender As Object, e As EventArgs) Handles sensores.Tick
        If miTurno = -1 Then Exit Sub
        If revisandoSensores Or Not estadoPrograma Then Exit Sub
        If modulos(5) = 0 Then Exit Sub

        If primerSensor Then primerSensor = False
        Dim ultimo = DateAndTime.Now()
        revisandoSensores = True

        Dim regsAfectados As Long = 0
        Dim cadAdic = ""
        Dim cadAdic2 = ""
        Dim produccion As Long = 0, produccion_seg = 0, calidad As Long = 0, buffer As Long = 0
        Dim produccion_tc As Double = 0, calidad_tc As Double = 0
        Dim general As DataSet
        incluyeHoyos = False
        Dim cadSQL = "SELECT turno_oee, turno_secuencia, andon_prorrateado FROM " & rutaBD & ".configuracion"
        Dim config As DataSet = consultaSEL(cadSQL)
        Dim aCortar = False
        Dim TTotal As Long = 0
        Dim hayTipo3 As Boolean
        Dim AP As Boolean = False
        If config.Tables(0).Rows.Count > 0 Then


            AP = ValNull(config.Tables(0).Rows(0)!andon_prorrateado, "A") = "S"
            miSecuencia = config.Tables(0).Rows(0)!turno_secuencia
            'Se calcula el turno
            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas SET estatus = 1 WHERE estatus = 0")
            cadSQL = "Select c.ultima_hora, c.paro_emergencia, c.proceso, c.id As maquinaid, c.replanear, c.replanear_desde, c.compaginar, c.compaginar_desde, c.paro_wip, c.oee_estado, c.oee_estado_desde, c.oee_historico_rate, c.oee_historico_rate_reiniciar, c.linea, c.oee_umbral_produccion, c.nombre As nequipo, e.nombre As nparte, e.referencia, e.referencia, f.nombre As ntripulacion, g.numero As norden, d.nombre As nturno, c.oee_turno_actual, c.oee_tripulacion_actual, c.oee_operador_actual, c.oee_lote_actual, c.oee_lote_anterior, c.oee_lote_cambio, c.oee_parte_actual, b.id, b.equipo, b.tipo, b.area As area_sensor, b.clasificacion As clasificacion_sensor, b.kg_saco_parte, c.pieza, b.multiplicador, b.base, Date(a.fecha) As fecha, IFNULL(SUM(a.valor), 0) As totals, d.cambiodia, d.mover, d.inicia, d.termina, d.secuencia FROM " & rutaBD & ".cat_maquinas c INNER JOIN " & rutaBD & ".relacion_procesos_sensores b On b.equipo = c.id And b.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_partes e ON c.oee_parte_actual = e.id AND e.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_tripulacion f ON c.oee_tripulacion_actual = f.id AND f.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_turnos d ON c.oee_turno_actual = d.id AND d.estatus = 'A' LEFT JOIN " & rutaBD & ".lecturas a ON b.sensor = a.sensor AND a.estatus = 1 LEFT JOIN " & rutaBD & ".lotes g ON c.oee_lote_actual = g.id AND g.estatus = 'A' WHERE c.oee = 'S' AND c.estatus = 'A' GROUP BY c.ultima_hora, c.paro_emergencia, c.linea, c.nombre, e.nombre, e.referencia, f.nombre, b.equipo, b.tipo, b.multiplicador, b.kg_saco_parte, c.pieza, b.base, DATE(a.fecha), c.oee_turno_actual, c.oee_tripulacion_actual, c.oee_operador_actual, c.oee_lote_actual, c.oee_lote_anterior, c.oee_lote_cambio, c.oee_parte_actual, d.mover, d.cambiodia"

            Dim capturas = consultaSEL(cadSQL)
            Dim paroEmergencia As Boolean = False
            If capturas.Tables(0).Rows.Count > 0 Then
                Dim equipoActual = 0
                For Each captura In capturas.Tables(0).Rows
                    If ValNull(captura!nturno, "A") = "" Then
                        Dim elvis = 1
                    End If
                    Dim cantidadAntes As Boolean = False
                    hayTipo3 = False
                    Dim cadPlaneado As String = ", planeado = 'N'"
                    produccion_seg = 0
                    produccion_tc = 0
                    calidad_tc = 0
                    'Buscar rate
                    Dim rateEquipo As Double = 1
                    Dim rateBajo As Double = 75
                    Dim rateAlto As Double = 125
                    Dim medTiempo As Long = 0
                    Dim rateUnidad = "PzaxHr"
                    Dim cadSQLParo = ""
                    Dim funcionando As Boolean

                    If captura!maquinaid = 1 Then
                        Dim uno = 1
                    End If

                    cadSQL = "SELECT piezas, bajo, alto, tiempo, unidad FROM " & rutaBD & ".relacion_partes_equipos WHERE (equipo = " & captura!maquinaid & " Or equipo = 0) And (parte = " & captura!oee_parte_actual & " Or parte = 0) ORDER BY parte DESC, equipo DESC LIMIT 1"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        rateEquipo = general.Tables(0).Rows(0)!piezas
                        rateBajo = general.Tables(0).Rows(0)!bajo
                        rateAlto = general.Tables(0).Rows(0)!alto
                        medTiempo = general.Tables(0).Rows(0)!tiempo
                        rateUnidad = ValNull(general.Tables(0).Rows(0)!unidad, "A")
                    End If

                    If rateEquipo = 0 Then rateEquipo = 1
                    Dim TC As Double = 0
                    If rateEquipo = 0 Then
                        TC = 1
                    ElseIf medTiempo = 2 Then
                        TC = 3600 / rateEquipo
                    ElseIf medTiempo = 1 Then
                        TC = 60 / rateEquipo
                        ' TC = rateEquipo / 60
                    ElseIf medTiempo = 0 Then
                        TC = 1 / rateEquipo
                    ElseIf medTiempo = 3 Then
                        TC = 86400 / rateEquipo
                    End If

                    TC = Math.Round(TC, 10)

                    'Buscar Paro

                    Dim conProduccion = False
                    Dim conRateBajo = False
                    cadAdic = ""
                    cadAdic2 = ""
                    Dim paroActual As Long = 0
                    Dim reporteActual As Long = 0
                    Dim paroDesde = ultimo
                    Dim paroCorte = ultimo
                    Dim fechaEstado = ultimo
                    Dim tiempoEstado As Long = 0
                    Dim tiempoParo As Long = 0
                    Dim tiempoParoCorte As Long = 0
                    Dim claseParoActual = 0
                    Dim finaliza_sensor = False
                    Dim piezasWIP = 0
                    Dim cadDetener = ""
                    Dim cadIniciarParo = ""
                    paroEmergencia = captura!paro_emergencia = "S"
                    Dim salioEmergencia As Boolean = False
                    Dim paroCreado = False

                    'Se pregunta si hay un paro activo
                    Dim cadSQLOri = "SELECT id, clase, corte, reporte, finaliza_sensor, wip_piezas, tipo, estado, inicia, finaliza, wip_piezas FROM " & rutaBD & ".detalleparos WHERE maquina = " & captura!maquinaid & " AND estatus = 'A' AND estado = 'C' ORDER BY tipo, inicia LIMIT 1"
                    general = consultaSEL(cadSQLOri)
                    TTotal = 0
                    'Del mismo modo, se busca si hay un paro planeado vencido
                    cadSQL = "SELECT id FROM " & rutaBD & ".detalleparos WHERE maquina = " & captura!maquinaid & " AND desde <= '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' AND hasta > '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' AND estatus = 'A' AND clase = 0 AND estado = 'L' ORDER BY tipo, inicia LIMIT 1"
                    Dim general2 As DataSet = consultaSEL(cadSQL)
                    TTotal = 0
                    If general2.Tables(0).Rows.Count > 0 Then
                        'Hay paro vencido
                        If general.Tables(0).Rows.Count > 0 Then
                            'Si hay un paro activo, se termina
                            If Not general.Tables(0).Rows(0)!corte.Equals(System.DBNull.Value) Then
                                paroCorte = general.Tables(0).Rows(0)!corte
                            Else
                                paroCorte = general.Tables(0).Rows(0)!inicia
                            End If
                            If paroCorte <> ultimo Then
                                tiempoParo = tiempoValido(paroCorte, ultimo, captura!proceso)
                                tiempoParoCorte = tiempoParo
                            End If
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(general.Tables(0).Rows(0)!inicia, ultimo, captura!proceso) & ", finalizo = 1, finalizo_accion = 'O' WHERE id = " & general.Tables(0).Rows(0)!id & ";UPDATE " & rutaBD & ".detalleparos SET hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & general.Tables(0).Rows(0)!id & " AND clase <> 0")

                        End If
                        'Se activa el paro vencido
                        cadDetener = "UPDATE " & rutaBD & ".detalleparos SET estado = 'C', turno = " & config.Tables(0).Rows(0)!turno_oee & ", inicia = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', inicio = 1, turno = " & captura!oee_turno_actual & ", fecha = '" & Format(ultimo, "yyyy/MM/dd") & "' WHERE id = " & general2.Tables(0).Rows(0)!id & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', parada_desde  = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid
                        regsAfectados = consultaACT(cadDetener)

                        general = consultaSEL(cadSQLOri)
                    End If

                    If general.Tables(0).Rows.Count > 0 Then
                        'Si hay un paro planeado o no, se pregunta si ya se venció
                        cadSQL = "SELECT id, inicia, corte FROM " & rutaBD & ".detalleparos WHERE id = " & general.Tables(0).Rows(0)!id & " AND hasta <= '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                        general2 = consultaSEL(cadSQL)
                        TTotal = 0
                        If general2.Tables(0).Rows.Count > 0 Then
                            'Si ya se venció se calcula su monto y se cierra
                            If Not general2.Tables(0).Rows(0)!corte.Equals(System.DBNull.Value) Then
                                paroCorte = general2.Tables(0).Rows(0)!corte
                            Else
                                paroCorte = general2.Tables(0).Rows(0)!inicia
                            End If
                            If paroCorte <> ultimo Then
                                tiempoParoCorte = tiempoValido(paroCorte, ultimo, captura!proceso)
                                tiempoParo = tiempoParo + tiempoParoCorte
                            End If
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(general2.Tables(0).Rows(0)!inicia, ultimo, captura!proceso) & ", finalizo = 1 WHERE id = " & general2.Tables(0).Rows(0)!id & ";UPDATE " & rutaBD & ".detalleparos SET hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & general2.Tables(0).Rows(0)!id & " AND clase <> 0;UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', parada_desde = NULL, estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                            If regsAfectados > 0 Then

                                paroActual = 0
                                reporteActual = 0
                                claseParoActual = 0
                                finaliza_sensor = False
                                piezasWIP = 0
                            End If
                        Else
                            'Se mantiene el paro actual
                            paroActual = general.Tables(0).Rows(0)!id
                            reporteActual = general.Tables(0).Rows(0)!reporte
                            claseParoActual = general.Tables(0).Rows(0)!clase
                            finaliza_sensor = ValNull(general.Tables(0).Rows(0)!finaliza_sensor, "A") = "S"
                            If claseParoActual = 0 Then
                                cadPlaneado = ", planeado = 'S'"
                                piezasWIP = ValNull(general.Tables(0).Rows(0)!wip_piezas, "N")
                            End If

                            paroDesde = general.Tables(0).Rows(0)!inicia

                            If Not general.Tables(0).Rows(0)!corte.Equals(System.DBNull.Value) Then
                                paroCorte = general.Tables(0).Rows(0)!corte
                            Else
                                paroCorte = paroDesde
                            End If

                            Dim pFinalizado = False
                            If reporteActual > 0 Then
                                'Si el paro proviene de un reporte, se pregunta el estatus de éste


                                cadSQL = "SELECT id, fecha, estatus, cierre_atencion FROM " & rutaBD & ".reportes WHERE id = " & reporteActual & " AND afecta_oee = 'S'"

                                Dim miReporte = consultaSEL(cadSQL)

                                If miReporte.Tables(0).Rows.Count > 0 Then
                                    'Se cerró el reporte
                                    If miReporte.Tables(0).Rows(0)!estatus > 10 Then
                                        Dim miTiempo = tiempoValido(paroDesde, miReporte.Tables(0).Rows(0)!cierre_atencion, captura!proceso)
                                        Dim fCierre = DateAndTime.Now
                                        If Not miReporte.Tables(0).Rows(0)!cierre_atencion.Equals(System.DBNull.Value) Then
                                            fCierre = miReporte.Tables(0).Rows(0)!cierre_atencion
                                        End If
                                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', tiempo = " & miTiempo & ", finaliza = '" & Format(fCierre, "yyyy/MM/dd HH:mm:ss") & "', hasta = '" & Format(fCierre, "yyyy/MM/dd HH:mm:ss") & "', finalizo_accion = 'R' WHERE reporte = " & reporteActual & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                        reporteActual = 0
                                        paroActual = 0
                                        claseParoActual = 0
                                        pFinalizado = True
                                        finaliza_sensor = False
                                        piezasWIP = 0

                                    End If
                                End If
                            End If
                            If Not pFinalizado Then
                                'Si el paro aún esta vivo
                                If ValNull(captura!oee_estado, "A") = "S" Then
                                    'Se reactiva la máquina
                                    fechaEstado = captura!oee_estado_desde
                                    tiempoEstado = DateAndTime.DateDiff(DateInterval.Second, fechaEstado, ultimo)
                                    TTotal = tiempoEstado
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado = 'N', oee_estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', oee_estado_cambio = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid)
                                    funcionando = False
                                End If
                                'Se contabiliza el tiempo del paro
                                If paroCorte <> ultimo Then
                                    tiempoParoCorte = tiempoValido(paroCorte, ultimo, captura!proceso)
                                    tiempoParo = tiempoParo + tiempoParoCorte
                                End If
                            End If
                        End If
                    End If

                    produccion_seg = 0
                    produccion = 0
                    calidad = 0
                    buffer = 0
                    Dim piezas As Long = 0
                    Dim fecha = ultimo
                    If captura!totals > 0 And captura!tipo = 0 Then
                        fecha = captura!fecha
                        piezas = captura!totals * captura!multiplicador
                        If captura!base > 0 Then
                            piezas = piezas / captura!base
                        End If
                    End If
                    If ValNull(captura!cambiodia, "A") = "S" Then
                        If ValNull(captura!mover, "N") = 1 And Format(fecha, "HH") >= "00" And Format(fecha, "HH:mm:ss") < captura!termina.ToString Then
                            fecha = DateAdd(DateInterval.Day, -1, fecha)
                        ElseIf ValNull(captura!mover, "N") = 2 And Format(fecha, "HH:mm:ss") >= captura!inicia.ToString And Format(fecha, "HHmmss") <= "235959" Then
                            fecha = DateAdd(DateInterval.Day, 1, fecha)
                        End If
                    End If

                    'If captura!tipo = 3 Then
                    ' hayTipo3 = True
                    ' cadSQL = "SELECT oee_estado, oee_estado_desde FROM " & rutaBD & ".cat_maquinas WHERE id = " & captura!maquinaid
                    ' general = consultaSEL(cadSQL)
                    ' If general.Tables(0).Rows.Count > 0 Then
                    ' If ValNull(general.Tables(0).Rows(0)!oee_estado, "A") = "S" Then
                    ' tiempoEstado = DateAndTime.DateDiff(DateInterval.Second, general.Tables(0).Rows(0)!oee_estado_desde, ultimo)
                    ' TTotal = TTotal + tiempoEstado
                    ' regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F' WHERE equipo = " & captura!maquinaid)
                    'funcionando = True
                    ' End If
                    ' End If
                    ' End If

                    'Buscar objetivo
                    Dim objetivo As Double = 0
                    Dim van As Double = 0
                    Dim reinicio As Long = 0
                    Dim cadTarimas = ", 0, 0"
                    Dim cadTarimas2 = ""
                    cadSQL = "Select id, objetivo, reinicio, lote, van, kg_saco, sacos_tarima, notificada_fin, notificada_inicio, notificada_completada, notificada_cambio FROM " & rutaBD & ".equipos_objetivo WHERE (equipo = " & captura!maquinaid & " Or equipo = 0) And (parte = " & captura!oee_parte_actual & " Or parte = 0) And (fijo = 'S' OR ('" & Format(fecha, "yyyy/MM/dd") & "' >= desde AND '" & Format(fecha, "yyyy/MM/dd") & "' <= hasta AND (turno = 0 OR turno = " & captura!oee_turno_actual & ") AND (lote = 0 OR lote = " & captura!oee_lote_actual & "))) ORDER BY parte DESC, equipo DESC, fijo, lote DESC, turno DESC LIMIT 1"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        If ValNull(captura!kg_saco_parte, "A") = "S" And piezas > 0 And general.Tables(0).Rows(0)!kg_saco > 0 Then
                            piezas = piezas * general.Tables(0).Rows(0)!kg_saco
                        End If
                        objetivo = general.Tables(0).Rows(0)!objetivo
                        van = general.Tables(0).Rows(0)!van + piezas
                        Dim cadNotificar As String = ""
                        Dim enviarMensajeOP As Boolean = False
                        If van >= objetivo And general.Tables(0).Rows(0)!notificada_completada = "N" Then
                            enviarMensajeOP = True
                            cadNotificar = ", notificada_completada = 'S'"
                        End If
                        If captura!oee_lote_cambio = "S" Then
                            mensajeOPCambio(captura!oee_lote_actual, captura!oee_lote_anterior)
                            cadNotificar = cadNotificar & ", notificada_cambio = 'S'"
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_lote_cambio = 'N' WHERE id = " & captura!maquinaid)
                        End If

                        reinicio = general.Tables(0).Rows(0)!reinicio
                        Dim tmp As Double = 0
                        If general.Tables(0).Rows(0)!kg_saco > 0 Then

                            tmp = van / general.Tables(0).Rows(0)!kg_saco
                            If tmp - Math.Floor(tmp) > 0 Then
                                tmp = Math.Floor(tmp) + 1
                            End If
                            cadTarimas = ", " & tmp
                            cadTarimas2 = ", sacos = " & tmp
                        Else
                            cadTarimas = ", 0"
                        End If
                        If general.Tables(0).Rows(0)!sacos_tarima > 0 Then
                            tmp = tmp / general.Tables(0).Rows(0)!sacos_tarima
                            If tmp - Math.Floor(tmp) > 0 Then

                                tmp = Math.Floor(tmp) + 1
                            End If
                            cadTarimas = cadTarimas & ", " & tmp
                            cadTarimas2 = cadTarimas2 & ", tarimas = " & tmp
                        Else
                            cadTarimas = cadTarimas & ", 0"
                        End If

                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".equipos_objetivo Set van = van + " & piezas & cadNotificar & " WHERE id = " & general.Tables(0).Rows(0)!id)


                        If enviarMensajeOP Then
                            mensajeOPCompleta(captura!oee_lote_actual)
                        End If

                    End If
                    If captura!oee_operador_actual <> 0 And operador = 0 Then
                        operador = captura!oee_operador_actual
                    End If
                    'Buscar si ya existe el corte
                    cadSQL = "SELECT id, paro_actual, bloque_inicia, bloque_finaliza FROM " & rutaBD & ".lecturas_cortes WHERE dia = '" & Format(fecha, "yyyy/MM/dd") & "' AND equipo = " & captura!maquinaid & " AND orden = " & captura!oee_lote_actual & " AND parte = " & captura!oee_parte_actual & " AND turno = " & captura!oee_turno_actual & " AND tc = " & TC & " AND tripulacion = " & captura!oee_tripulacion_actual & " AND operador = " & operador & " ORDER BY id DESC LIMIT 1"
                    Dim miEquipo = consultaSEL(cadSQL)
                    Dim cadParo = ""
                    Dim idCorte = 0
                    Dim existeCorte As Boolean = False
                    If miEquipo.Tables(0).Rows.Count > 0 Then
                        cadSQL = "SELECT corte FROM " & rutaBD & ".relacion_maquinas_lecturas WHERE equipo = " & captura!maquinaid & " AND corte = " & miEquipo.Tables(0).Rows(0)!id
                        general = consultaSEL(cadSQL)
                        existeCorte = general.Tables(0).Rows.Count > 0
                    End If
                    If existeCorte Then
                        idCorte = miEquipo.Tables(0).Rows(0)!id
                        If captura!totals > 0 Then
                            If captura!tipo = 0 Then
                                cadAdic = ", produccion = produccion + " & piezas
                                cadAdic = cadAdic & ", produccion_tc = produccion_tc + " & piezas * TC
                                produccion = piezas
                                produccion_tc = piezas * TC
                                produccion_seg = piezas
                                'ElseIf captura!tipo = 4 Then
                                '    cadAdic = ", calidad = calidad + " & piezas
                                '    cadAdic = cadAdic & ", calidad_tc = calidad_tc + " & piezas * TC
                                '    calidad = piezas
                                '    calidad_tc = piezas * TC
                            ElseIf captura!tipo = 1 Then
                                'cantidadAntes = True
                                'cadAdic = ", calidad_antes = calidad_antes + " & piezas & ", calidad = calidad + " & piezas & ", produccion = produccion + " & piezas
                                'cadAdic = cadAdic & ", calidad_tc = calidad_tc + " & piezas * TC & ", produccion_tc = produccion_tc + " & piezas * TC
                                'calidad = piezas
                                'calidad_tc = piezas * TC
                                'produccion = piezas
                                'produccion_tc = piezas * TC
                                'produccion_seg = piezas
                                paroEmergencia = True
                            ElseIf captura!tipo = 2 Then
                                'cantidadAntes = True
                                'cadAdic = ", calidad_antes = calidad_antes + " & piezas & ", calidad = calidad + " & piezas & ", produccion = produccion + " & piezas
                                'cadAdic = cadAdic & ", calidad_tc = calidad_tc + " & piezas * TC & ", produccion_tc = produccion_tc + " & piezas * TC
                                'calidad = piezas
                                'calidad_tc = piezas * TC
                                'produccion = piezas
                                'produccion_tc = piezas * TC
                                'produccion_seg = piezas
                                paroEmergencia = False
                                'ElseIf captura!tipo = 2 Then
                                '    cadAdic = ", buffer = buffer + " & piezas
                                '    buffer = piezas
                                'ElseIf captura!tipo = 3 Then
                                '
                                '                                cadSQL = "SELECT fecha, valor FROM " & rutaBD & ".lecturas WHERE sensor = " & captura!id & " AND estatus = 1 ORDER BY id"
                                '                                Dim regsSensor = consultaSEL(cadSQL)
                                '                                If regsSensor.Tables(0).Rows.Count > 0 Then
                                '                                Dim fechaInicial = miEquipo.Tables(0).Rows(0)!bloque_inicia
                                '                                If Not captura!oee_estado_desde.Equals(System.DBNull.Value) Then
                                '                                fechaInicial = captura!oee_estado_desde
                                '                            End If
                                '                                For Each rSensor In regsSensor.Tables(0).Rows
                                '                                If rSensor!valor = 2 And ValNull(captura!oee_estado, "A") <> "S" Then
                                '                                'Encendido
                                '                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado = 'S', oee_estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "', oee_estado_cambio = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                '                               funcionando = True
                                '                              fechaInicial = rSensor!fecha
                                '                         ElseIf rSensor!valor = 1 And ValNull(captura!oee_estado, "A") = "S" Then
                                '                            'Apagado
                                '                           tiempoEstado = DateAndTime.DateDiff(DateInterval.Second, fechaInicial, rSensor!fecha)
                                '                          TTotal = TTotal + tiempoEstado
                                '                         regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado = 'N', oee_estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "', oee_estado_cambio = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                '                        funcionando = False
                                '                   End If
                                '      Next
                                '                     cadAdic = ", produccion = produccion + " & TTotal & ", produccion_tc = produccion_tc + " & TTotal
                                '                    produccion_seg = TTotal
                                '               End If
                            End If
                            '             ElseIf captura!tipo = 3 And TTotal > 0 Then
                            '                cadAdic = ", produccion = produccion + " & TTotal & ", produccion_tc = produccion_tc + " & TTotal
                            '               produccion_seg = TTotal
                        End If

                        If reporteActual > 0 And AP And tiempoParo > 0 Then
                            tiempoParo = tiempoParo - produccion_tc
                            tiempoParoCorte = tiempoParoCorte - produccion_tc
                        End If

                        cadAdic = cadAdic & ", paro = paro + " & tiempoParo
                        cadSQL = "UPDATE " & rutaBD & ".lecturas_cortes SET tiempo_disponible = " & tiempoValido(miEquipo.Tables(0).Rows(0)!bloque_inicia, ultimo, captura!proceso) & ", bloque_finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', operador = " & operador & " " & cadAdic & cadTarimas2 & " WHERE id = " & miEquipo.Tables(0).Rows(0)!id & cadParo & ";UPDATE " & rutaBD & ".lecturas_cortes SET paro = tiempo_disponible WHERE paro > tiempo_disponible AND id = " & miEquipo.Tables(0).Rows(0)!id
                    Else
                        'NUEVO LECTURA CORTE
                        'Calcular el corte anterior

                        If captura!totals > 0 Then
                            If captura!tipo = 0 Then
                                cadAdic = ", produccion, produccion_tc"
                                cadAdic2 = cadAdic2 & ", " & piezas & ", " & piezas * TC
                                produccion = piezas
                                produccion_seg = piezas
                                produccion_tc = piezas * TC
                                'ElseIf captura!tipo = 4 Then
                                '    cadAdic = ", calidad, calidad_tc"
                                '    cadAdic2 = cadAdic2 & ", " & piezas & ", " & piezas * TC
                                '    calidad = piezas
                                '    calidad_tc = piezas * TC
                            ElseIf captura!tipo = 1 Then
                                paroEmergencia = True
                            ElseIf captura!tipo = 2 Then
                                paroEmergencia = False
                                'cantidadAntes = True
                                'cadAdic = ", calidad, calidad_tc, produccion, produccion_tc, calidad_antes, "
                                'cadAdic2 = cadAdic2 & ", " & piezas & ", " & piezas * TC & ", " & piezas & ", " & piezas * TC & ", " & piezas
                                'calidad = piezas
                                'calidad_tc = piezas * TC
                                'produccion = piezas
                                'produccion_seg = piezas
                                'produccion_tc = piezas * TC

                                'ElseIf captura!tipo = 2 Then
                                '   cadAdic = ", buffer"
                                '  cadAdic2 = cadAdic2 & ", " & piezas
                                ' buffer = piezas
                                'ElseIf captura!tipo = 3 Then
                                '   cadSQL = "SELECT fecha, valor FROM " & rutaBD & ".lecturas WHERE sensor = " & captura!id & " AND estatus = 1 ORDER BY id"
                                '  Dim regsSensor = consultaSEL(cadSQL)
                                ' If regsSensor.Tables(0).Rows.Count > 0 Then
                                'Dim fechaInicial = ultimo
                                ' If Not captura!oee_estado_desde.Equals(System.DBNull.Value) Then
                                ' fechaInicial = captura!oee_estado_desde
                                ' End If
                                '     For Each rSensor In regsSensor.Tables(0).Rows
                                '     If rSensor!valor = 2 And ValNull(captura!oee_estado, "A") <> "S" Then
                                '    'Encendido
                                '   regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado = 'S', oee_estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "', oee_estado_cambio = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                '  funcionando = True
                                ' fechaInicial = rSensor!fecha
                                'ElseIf rSensor!valor = 1 And ValNull(captura!oee_estado, "A") = "S" Then
                                '   'Apagado
                                '  tiempoEstado = DateAndTime.DateDiff(DateInterval.Second, fechaInicial, rSensor!fecha)
                                ' TTotal = TTotal + tiempoEstado
                                'regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_estado = 'N', oee_estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "', oee_estado_cambio = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!maquinaid & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', estado_desde = '" & Format(rSensor!fecha, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                'funcionando = False
                                'End If
                                'Next
                                '               cadAdic = ", produccion, produccion_tc"
                                '              cadAdic2 = cadAdic2 & ", " & TTotal & ", " & TTotal
                                '             produccion_seg = TTotal
                                'End If
                            End If
                            'ElseIf captura!tipo = 3 And TTotal > 0 Then
                            '    cadAdic = ", produccion, produccion_tc"
                            '    cadAdic2 = cadAdic2 & ", " & TTotal & ", " & TTotal
                            '    produccion_seg = TTotal
                        End If

                        If reporteActual > 0 And AP And tiempoParo > 0 Then
                            tiempoParo = tiempoParo - produccion_tc
                            tiempoParoCorte = tiempoParoCorte - produccion_tc
                        End If

                        If paroActual > 0 Then
                            cadAdic = cadAdic & ", paro"
                            cadAdic2 = cadAdic2 & ", " & tiempoParo
                        End If
                        cadSQL = "INSERT " & rutaBD & ".lecturas_cortes (dia, orden, parte, turno, turno_secuencia, equipo, tripulacion" & cadAdic & ", bloque_inicia, bloque_finaliza, tc, operador, sacos, tarimas) VALUES ('" & Format(fecha, "yyyy/MM/dd") & "', " & captura!oee_lote_actual & ", " & captura!oee_parte_actual & ", " & captura!oee_turno_actual & ", " & miSecuencia & ", " & captura!maquinaid & ", " & captura!oee_tripulacion_actual & cadAdic2 & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', " & TC & ", " & operador & cadTarimas & ")" & cadParo
                    End If

                    'Dim primeraVez As Boolean = False
                    regsAfectados = consultaACT(cadSQL)
                    If idCorte = 0 Then
                        cadSQL = "SELECT MAX(id) AS id FROM " & rutaBD & ".lecturas_cortes"
                        general = consultaSEL(cadSQL)
                        If general.Tables(0).Rows.Count > 0 Then
                            idCorte = ValNull(general.Tables(0).Rows(0)!id, "N")
                        End If
                    End If

                    cadAdic2 = ""

                    If paroEmergencia And captura!pieza = "S" And piezas > 0 And captura!tipo = 0 Then
                        paroEmergencia = False
                    End If

                    If paroEmergencia And captura!paro_emergencia = "N" Or Not paroEmergencia And captura!paro_emergencia = "S" Then
                        Dim cadEmergencia = ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', parada_desde = NULL, estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid
                        If Not paroEmergencia Then salioEmergencia = True
                        If paroEmergencia Then
                            cadEmergencia = ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', parada_desde  = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid
                        End If
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET paro_emergencia = '" & IIf(paroEmergencia, "S", "N") & "' WHERE id = " & captura!maquinaid & cadEmergencia)
                        'Enviar mensaje de paro de emergencia
                        mensajeEmergencia(paroEmergencia)
                    End If

                    'Buscar estimados
                    Dim EFI As Double = 100, OEE As Double = 85, FTQ As Double = 100, DIS As Double = 100
                    Dim eEFI As Double = 100, eOEE As Double = 85, eFTQ As Double = 100, eDIS As Double = 100
                    cadSQL = "Select * FROM " & rutaBD & ".estimados WHERE (linea = " & captura!linea & " Or linea = 0) And (equipo = " & captura!maquinaid & " Or equipo = 0) And (fijo = 'S' OR ('" & Format(fecha, "yyyy/MM/dd") & "' >= desde AND '" & Format(fecha, "yyyy/MM/dd") & "' <= hasta )) ORDER BY equipo DESC, linea DESC, fijo, desde LIMIT 1"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        eOEE = general.Tables(0).Rows(0)!oee
                        eEFI = general.Tables(0).Rows(0)!efi
                        eFTQ = general.Tables(0).Rows(0)!ftq
                        eDIS = general.Tables(0).Rows(0)!dis
                    End If



                    Dim produccion_solo As Double = produccion
                    Dim produccion_tc_solo As Double = produccion_tc
                    Dim calidad_solo As Double = calidad
                    Dim calidad_tc_solo As Double = calidad_tc


                    cadSQL = "Select tc, lote, parte, turno, id, dia, desde, tiempo, buenas, malas, disponible, buenas_vienen, malas_vienen, diferencia_vienen, plan FROM " & rutaBD & ".horaxhora WHERE dia = '" & Format(ultimo, "yyyy/MM/dd") & "' AND hora = " & Format(ultimo, "HH") & " AND estatus = 'A' AND equipo = " & captura!maquinaid & " ORDER BY id DESC"
                    Dim hxh = consultaSEL(cadSQL)
                    Dim laHora = Format(ultimo, "HH")
                    Dim cadHXH = ""
                    Dim vienenMalas = 0
                    Dim vienenBuenas = 0
                    Dim arrastre = 0
                    Dim totalVienen = 0
                    Dim estHXH = 0
                    Dim cadAdicCorte As String = ""
                    Dim crearNuevo = False
                    Dim tipoCorte As Integer = 0
                    Dim fechaFin = ultimo
                    Dim fechaIni = DateAdd(DateInterval.Second, 1, ultimo)
                    If hxh.Tables(0).Rows.Count = 0 Then
                        crearNuevo = True
                    ElseIf hxh.Tables(0).Rows(0)!lote <> captura!oee_lote_actual Or hxh.Tables(0).Rows(0)!parte <> captura!oee_parte_actual Or hxh.Tables(0).Rows(0)!tc <> TC Then

                        If hxh.Tables(0).Rows(0)!tc <> TC Then
                            tipoCorte = 5
                        ElseIf hxh.Tables(0).Rows(0)!parte <> captura!oee_parte_actual Then
                            tipoCorte = 3

                        ElseIf hxh.Tables(0).Rows(0)!lote <> captura!oee_lote_actual Then
                            tipoCorte = 4
                        End If

                        crearNuevo = True
                    Else
                        If hxh.Tables(0).Rows(0)!id <> captura!ultima_hora Then
                            'Se buscan los pendientes
                            If captura!ultima_hora > 0 Then
                                cadSQL = "Select dia, hora, hasta, malas, buenas, plan, malas_vienen, buenas_vienen, plan_van, diferencia_vienen FROM " & rutaBD & ".horaxhora WHERE id = " & captura!ultima_hora
                                general = consultaSEL(cadSQL)
                                If general.Tables(0).Rows.Count > 0 Then
                                    vienenMalas = general.Tables(0).Rows(0)!malas + general.Tables(0).Rows(0)!malas_vienen
                                    vienenBuenas = general.Tables(0).Rows(0)!buenas + general.Tables(0).Rows(0)!buenas_vienen
                                    totalVienen = (general.Tables(0).Rows(0)!buenas + general.Tables(0).Rows(0)!malas) - general.Tables(0).Rows(0)!plan + +general.Tables(0).Rows(0)!diferencia_vienen
                                    cadAdicCorte = ", buenas_vienen = " & vienenBuenas & ", malas_vienen = " & vienenMalas & ", diferencia_vienen = " & totalVienen
                                    cadHXH = cadHXH & ";UPDATE " & rutaBD & ".horaxhora SET estatus = 'Z' WHERE id = " & captura!ultima_hora
                                End If

                            End If
                            cadHXH = cadHXH & ";UPDATE " & rutaBD & ".cat_maquinas SET ultima_hora = " & hxh.Tables(0).Rows(0)!id & " WHERE id = " & captura!maquinaid
                        End If
                        Dim horaDesde = Convert.ToDateTime(Format(hxh.Tables(0).Rows(0)!dia, "yyyy/MM/dd") & " " & hxh.Tables(0).Rows(0)!desde)
                        If paroCorte < horaDesde And horaDesde <> ultimo And tiempoParoCorte > 0 Then
                            tiempoParoCorte = tiempoValido(horaDesde, ultimo, captura!proceso)
                        End If
                        Dim proyectado = 0
                        If hxh.Tables(0).Rows(0)!disponible > 0 Then
                            proyectado = Math.Round(hxh.Tables(0).Rows(0)!tiempo * hxh.Tables(0).Rows(0)!plan / hxh.Tables(0).Rows(0)!disponible, 0)
                        End If
                        arrastre = (hxh.Tables(0).Rows(0)!buenas + hxh.Tables(0).Rows(0)!malas) - proyectado
                        Dim tiempoECorte = tiempoValido(horaDesde, ultimo, captura!proceso)
                        cadHXH = cadHXH & ";UPDATE " & rutaBD & ".horaxhora SET buffer = buffer + " & buffer & ", buenas = buenas + " & produccion_solo & ", malas = malas + " & calidad_solo & ", buenas_tc = buenas_tc + " & produccion_tc_solo & ", malas_tc = malas_tc + " & calidad_tc_solo & ", paro = " & tiempoParoCorte & ", tiempo = " & tiempoECorte & ", tocada = 1, secuencia = " & miSecuencia & ", turno = " & config.Tables(0).Rows(0)!turno_oee & ", arrastre = " & arrastre & cadAdicCorte & ", tripulacion_inicial = " & captura!oee_tripulacion_actual & ", operador_inicial = " & operador & " WHERE id = " & hxh.Tables(0).Rows(0)!id
                    End If
                    Dim reCompaginar = False
                    Dim planSumado = van
                    Dim nuevoPlan = 0, nvoPlan_van = van

                    If crearNuevo Then
                        If captura!ultima_hora > 0 Then
                            cadSQL = "Select dia, hora, desde, disponible, plan_van, mantto, hasta, malas, buenas, plan, malas_vienen, buenas_vienen, diferencia_vienen FROM " & rutaBD & ".horaxhora WHERE id = " & captura!ultima_hora
                            general = consultaSEL(cadSQL)
                            If general.Tables(0).Rows.Count > 0 Then
                                vienenMalas = general.Tables(0).Rows(0)!malas + general.Tables(0).Rows(0)!malas_vienen
                                vienenBuenas = general.Tables(0).Rows(0)!buenas + general.Tables(0).Rows(0)!buenas_vienen
                                totalVienen = (general.Tables(0).Rows(0)!buenas + general.Tables(0).Rows(0)!malas) - general.Tables(0).Rows(0)!plan + +general.Tables(0).Rows(0)!diferencia_vienen
                                cadAdicCorte = ", buenas_vienen = " & vienenBuenas & ", malas_vienen = " & vienenMalas & ", diferencia_vienen = " & totalVienen
                                If general.Tables(0).Rows(0)!hasta <> Format(fechaFin, "HH:mm:ss") And tipoCorte > 0 Then
                                    calcularTiempos(Convert.ToDateTime(Format(general.Tables(0).Rows(0)!dia, "yyyy/MM/dd") & " " & general.Tables(0).Rows(0)!desde), ultimo, captura!proceso, captura!maquinaid)
                                    If objetivo - van > 0 Then
                                        reCompaginar = True

                                    End If
                                    If general.Tables(0).Rows(0)!disponible - general.Tables(0).Rows(0)!mantto > 0 Then
                                        nuevoPlan = (tDisponible - tMantto) * general.Tables(0).Rows(0)!plan / (general.Tables(0).Rows(0)!disponible - general.Tables(0).Rows(0)!mantto)
                                        nvoPlan_van = general.Tables(0).Rows(0)!plan_van - general.Tables(0).Rows(0)!plan + nuevoPlan
                                    End If
                                    cadHXH = cadHXH & ";UPDATE " & rutaBD & ".horaxhora SET estatus = 'Z', hasta = '" & Format(fechaFin, "HH:mm:ss") & "'" & IIf(tipoCorte > 0, ", ruptura = " & tipoCorte, "") & ", disponible = " & tDisponible & ", mantto = " & tMantto & ", plan = " & nuevoPlan & ", plan_van = " & nvoPlan_van & " WHERE id = " & captura!ultima_hora
                                Else
                                    cadHXH = cadHXH & ";UPDATE " & rutaBD & ".horaxhora SET estatus = 'Z', hasta = '" & Format(fechaFin, "HH:mm:ss") & "'" & IIf(tipoCorte > 0, ", ruptura = " & tipoCorte, "") & " WHERE id = " & captura!ultima_hora
                                End If
                            End If

                        End If

                        calcularTiempos(fechaIni, Convert.ToDateTime(Format(fechaIni, "yyyy/MM/dd HH") & ":59:59"), captura!proceso, captura!maquinaid)
                        Dim miPlan = 0
                        If (tDisponible - tMantto) > TC Then
                            miPlan = Math.Round((tDisponible - tMantto) / TC, 0)
                        End If
                        cadHXH = cadHXH & ";INSERT INTO " & rutaBD & ".horaxhora (equipo, dia, hora, desde, hasta, buffer, buenas, malas, buenas_tc, malas_tc, lote, parte, disponible, mantto, tocada, secuencia, turno, buenas_vienen, malas_vienen, diferencia_vienen, tripulacion_inicial, operador_inicial, plan, plan_van, tipo, TC, ruptura) VALUES (" & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd") & "', " & laHora & ", '" & Format(fechaIni, "HH:mm:ss") & "', '" & laHora & ":59:59', " & buffer & ", " & produccion_solo & ", " & calidad_solo & ", " & produccion_tc_solo & ", " & calidad_tc_solo & ", " & captura!oee_lote_actual & ", " & captura!oee_parte_actual & "," & tDisponible & ", " & tMantto & ", 1, " & miSecuencia & ", " & config.Tables(0).Rows(0)!turno_oee & ", " & vienenBuenas & ", " & vienenMalas & ", " & totalVienen & ", " & captura!oee_tripulacion_actual & ", " & operador & ", " & miPlan & ", " & miPlan + nvoPlan_van & ", 1, " & TC & ", " & tipoCorte & ");UPDATE " & rutaBD & ".horaxhora SET estatus = 'Z', tocada = 1 WHERE (dia = '" & Format(ultimo, "yyyy/MM/dd") & "' AND hora < " & Format(ultimo, " HH") & " OR dia < '" & Format(fecha, "yyyy/MM/dd") & "') AND equipo = " & captura!maquinaid & " AND lote = " & captura!oee_lote_actual & " AND estatus = 'A'"
                        'Se buscan los pendientes
                        planSumado = miPlan + nvoPlan_van
                    End If
                    regsAfectados = consultaACT(cadHXH)
                    If crearNuevo Then
                        cadSQL = "Select MAX(id) As id FROM " & rutaBD & ".horaxhora WHERE equipo = " & captura!maquinaid & " And estatus = 'A' "
                        general = consultaSEL(cadSQL)
                        If general.Tables(0).Rows.Count > 0 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET ultima_hora = " & general.Tables(0).Rows(0)!id & " WHERE id = " & captura!maquinaid)
                        End If
                    End If
                    If reCompaginar Then
                        fechaFin = DateAdd(DateInterval.Second, 1, Convert.ToDateTime(Format(fechaIni, "yyyy/MM/dd HH") & ":59:59"))
                        crearProgramacion(captura!maquinaid, captura!oee_lote_actual, captura!oee_parte_actual, fechaFin, captura!proceso, objetivo, planSumado)
                    End If

                    If captura!paro_emergencia = "N" And paroEmergencia Then
                        If paroActual > 0 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finalizo = 1, finalizo_accion = 'O' WHERE id = " & paroActual)
                            If regsAfectados > 0 Then
                                paroActual = 0
                                reporteActual = 0
                            End If
                        End If

                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, reporte, maquina, fecha, tiempo, inicio, inicia, desde, estado, estatus, area, tipo) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", '" & traduccion(187) & "', 8, 0, " & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd") & "', 0, 1, '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', 'C', 'A', 0, 0)")
                        paroCreado = regsAfectados > 0
                    ElseIf salioEmergencia And paroActual > 0 Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finalizo = 1, finalizo_accion = 'O' WHERE id = " & paroActual)
                        If regsAfectados > 0 Then
                            paroActual = 0
                            reporteActual = 0
                            claseParoActual = 0
                            finaliza_sensor = False
                            piezasWIP = 0
                        End If
                    End If


                    If equipoActual <> captura!maquinaid Or equipoActual = 0 Then
                        equipoActual = captura!maquinaid
                        '    primeraVez = True
                        'End If
                        'If primeraVez Then
                        If calidad > 0 And captura!clasificacion_sensor > 0 Then

                            'Se busca en la tabla que acumula
                            cadSQL = "SELECT id FROM " & rutaBD & ".detallerechazos WHERE equipo = " & captura!maquinaid & " AND parte = " & captura!oee_parte_actual & " AND fecha = '" & Format(fecha, "yyyy/MM/dd") & "' AND turno " & captura!oee_turno_actual & " AND area = " & captura!area_sensor & " AND tipo = " & captura!clasificacion_sensor & " AND (lote = 0 OR lote = " & captura!oee_lote_actual & ") AND origen = 0 AND estatus = 'A' ORDER BY id DESC LIMIT 1"
                            general = consultaSEL(cadSQL)
                            If general.Tables(0).Rows.Count > 0 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".detallerechazos SET cantidad = cantidad + " & calidad & ", cantidad_tc = cantidad_tc + " & calidad_tc & IIf(cantidadAntes, ", cantidad_antes = cantidad_antes + " & calidad, "") & " WHERE id = " & general.Tables(0).Rows(0)!id & ";UPDATE " & rutaBD & ".lecturas_cortes SET calidad_clasificada = calidad_clasificada + " & calidad & " WHERE id = " & idCorte)
                            Else
                                regsAfectados = consultaACT("INSERT " & rutaBD & ".detallerechazos (rechazo, tipo, area, equipo, fecha, turno, origen, corte, notas, parte, lote, cantidad, cantidad_tc, usuario, actualizacion, cantidad_antes) VALUES('" & traduccion(72) & captura!id & ", " & captura!clasificacion_sensor & ", " & captura!area_sensor & ", " & captura!maquinaid & ", '" & Format(fecha, "yyyy/MM/dd") & "', " & captura!oee_turno_actual & ", 1, " & idCorte & ", '" & traduccion(73).Replace("campo_0", captura!id) & "', " & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & calidad & ", " & calidad_tc & ", 1, NOW(), " & IIf(cantidadAntes, calidad, 0) & ");UPDATE " & rutaBD & ".lecturas_cortes SET calidad_clasificada = calidad_clasificada + " & calidad & " WHERE id = " & idCorte)
                            End If
                        End If
                        cadSQLParo = ""
                        cadAdic = ""
                        Dim rateActual As Double = 0
                        Dim ultimaProduccion = ultimo
                        Dim ultimaReparacion = ultimo

                        cadSQL = "SELECT a.equipo, wip_paro, wip_contador, wip_tiempo, parosmostrar, tiempo_reinicio, tiempo_corte, transcurrido, transcurrido_pasar, estatus, iniciar, iniciar_1, iniciar_2, iniciar_3, iniciar_4, iniciar_5, iniciar_6, iniciar_7, iniciar_8, detener, detener_piezas, detener_notas, detener_resultados, detener_estimado, detener_area, detener_tipo, detener_paro, reanudar, proximo_paro, ultima_produccion, ultima_reparacion, fecha_desde, desde_rate, produccion, produccion_tc, calidad, calidad_tc, buffer, rate_mal_desde, rate, rate_tendencia_baja, rate_tendencia_alta, rate_efecto, paros, ftq, efi, dis, oee, ftq_tendencia_baja, efi_tendencia_baja, dis_tendencia_baja, oee_tendencia_baja, paro_actual, IFNULL(SUM(piezashr), 0) AS piezashr, IFNULL(SUM(t_paros), 0) AS t_paros FROM " & rutaBD & ".relacion_maquinas_lecturas a LEFT JOIN (SELECT equipo, fecha, produccion AS piezashr, paro AS t_paros FROM " & rutaBD & ".piezasxminuto) AS b On a.equipo = b.equipo AND b.fecha >= a.desde_rate WHERE a.equipo = " & captura!maquinaid

                        Dim miResumen = consultaSEL(cadSQL)
                        If miResumen.Tables(0).Rows.Count = 0 Then
                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".relacion_maquinas_lecturas (equipo, ultima_produccion, ultima_buffer, fecha_desde, desde_rate, estado_desde) VALUES(" & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "')")
                            miResumen = consultaSEL(cadSQL)
                        ElseIf miResumen.Tables(0).Rows(0)!equipo.Equals(System.DBNull.Value) Then
                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".relacion_maquinas_lecturas (equipo, ultima_produccion, ultima_buffer, fecha_desde, desde_rate, estado_desde) VALUES(" & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "')")
                            miResumen = consultaSEL(cadSQL)
                        Else
                            If Not miResumen.Tables(0).Rows(0)!ultima_produccion.Equals(System.DBNull.Value) Then
                                ultimaProduccion = miResumen.Tables(0).Rows(0)!ultima_produccion
                            End If
                            If Not miResumen.Tables(0).Rows(0)!ultima_reparacion.Equals(System.DBNull.Value) Then
                                ultimaReparacion = miResumen.Tables(0).Rows(0)!ultima_reparacion
                            Else
                                ultimaReparacion = ultimaProduccion
                            End If
                        End If

                        Dim tiempoTranscurrido = 1
                        Dim corteAPP = ultimo

                        If Not miResumen.Tables(0).Rows(0)!fecha_desde.Equals(System.DBNull.Value) Then
                            corteAPP = miResumen.Tables(0).Rows(0)!fecha_desde
                            tiempoTranscurrido = tiempoValido(miResumen.Tables(0).Rows(0)!fecha_desde, ultimo, captura!proceso)
                        Else
                            cadAdic = cadAdic & ", fecha_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                        End If

                        Dim cortePT = ultimo

                        If Not miResumen.Tables(0).Rows(0)!desde_rate.Equals(System.DBNull.Value) Then
                            cortePT = miResumen.Tables(0).Rows(0)!desde_rate
                        Else
                            cadAdic = cadAdic & ", desde_rate = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                        End If

                        Dim tiempoTranscurridoRate = 1

                        Dim produccionRate = miResumen.Tables(0).Rows(0)!piezashr
                        Dim paroRate = miResumen.Tables(0).Rows(0)!t_paros
                        Dim minutosRate = 0
                        Dim horasAtras = -24
                        If captura!oee_historico_rate = 0 Then
                            If cortePT > corteAPP Then
                                corteAPP = cortePT
                                'No se busca en la BD
                            End If
                        ElseIf captura!oee_historico_rate = 1 Then
                            horasAtras = -1
                        ElseIf captura!oee_historico_rate = 2 Then
                            horasAtras = -2
                        ElseIf captura!oee_historico_rate = 3 Then
                            horasAtras = -6
                        ElseIf captura!oee_historico_rate = 4 Then
                            horasAtras = -12
                        End If

                        If captura!oee_historico_rate > 0 Then
                            corteAPP = cortePT
                            minutosRate = DateAndTime.DateDiff("m", cortePT, ultimo)
                            If corteAPP < DateAdd(DateInterval.Hour, horasAtras, ultimo) Then
                                corteAPP = DateAdd(DateInterval.Hour, horasAtras, ultimo)

                            End If
                        End If
                        If corteAPP <> cortePT Then
                            cadSQL = "SELECT IFNULL(SUM(produccion), 0) AS piezashr, IFNULL(SUM(paro), 0) AS t_paros FROM " & rutaBD & ".piezasxminuto WHERE fecha >= '" & Format(corteAPP, "yyyy/MM/dd HH:mm:ss") & "' AND equipo = " & captura!maquinaid

                            general = consultaSEL(cadSQL)
                            If general.Tables(0).Rows.Count > 0 Then
                                produccionRate = general.Tables(0).Rows(0)!piezashr
                                paroRate = general.Tables(0).Rows(0)!t_paros
                            End If
                        Else
                            minutosRate = DateAndTime.DateDiff(DateInterval.Minute, cortePT, ultimo)
                            If minutosRate = 0 Then minutosRate = 1
                        End If
                        tiempoTranscurridoRate = tiempoValido(corteAPP, ultimo, captura!proceso)

                        Dim disponibilidad = tiempoTranscurrido - miResumen.Tables(0).Rows(0)!parosmostrar
                        Dim disponibilidadRate = tiempoTranscurridoRate 'No se restan los paros del período - paroRate

                        produccion = miResumen.Tables(0).Rows(0)!produccion + produccion + TTotal
                        produccion_tc = miResumen.Tables(0).Rows(0)!produccion_tc + produccion_tc + TTotal

                        If disponibilidad <= 0 Then disponibilidad = 1
                        If disponibilidadRate <= 0 Then disponibilidadRate = 1

                        Dim uRate As Double = 0

                        rateActual = produccionRate / disponibilidadRate

                        If medTiempo = 2 Then
                            rateActual = rateActual * 3600
                        ElseIf medTiempo = 1 Then
                            rateActual = rateActual * 60
                            ' TC = rateEquipo / 60
                        ElseIf medTiempo = 0 Then

                        ElseIf medTiempo = 3 Then
                            rateActual = rateActual * 86400
                        End If

                        If miResumen.Tables(0).Rows(0)!produccion = produccion And paroActual > 0 Then
                            'rateActual = miResumen.Tables(0).Rows(0)!rate
                        End If
                        uRate = rateActual

                        Dim fueraHorario = "N"

                        fueraHorario = IIf(miResumen.Tables(0).Rows(0)!transcurrido = tiempoTranscurrido And miResumen.Tables(0).Rows(0)!transcurrido_pasar = 1, "S", "N")
                        'cadAdic = cadAdic & ", fuera_programa = '" & fueraHorario & "'"
                        If miResumen.Tables(0).Rows(0)!produccion < produccion Or fueraHorario = "S" Then
                            cadAdic = cadAdic & ", ultima_produccion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "',  alarmado_manual = 'N'"
                            ultimaProduccion = ultimo
                            ultimaReparacion = ultimo
                            conProduccion = miResumen.Tables(0).Rows(0)!produccion < produccion
                        End If
                        buffer = miResumen.Tables(0).Rows(0)!buffer + buffer
                        If miResumen.Tables(0).Rows(0)!buffer < buffer Then
                            cadAdic = cadAdic & ", ultima_buffer = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            ultimaProduccion = ultimo
                            ultimaReparacion = ultimo
                            conProduccion = True
                        End If

                        paroCreado = paroActual > 0
                        If (paroActual = 0 Or claseParoActual <> 0) And reporteActual = 0 Then
                            cadSQL = "SELECT a.id, a.fecha, a.area, b.nombre, b.agrupador_2 FROM " & rutaBD & ".reportes a LEFT JOIN " & rutaBD & ".cat_fallas b ON a.falla = b.id WHERE a.maquina = " & captura!maquinaid & " AND a.estatus <= 10 AND a.afecta_oee = 'S' ORDER BY a.id LIMIT 1"
                            general = consultaSEL(cadSQL)
                            If general.Tables(0).Rows.Count > 0 Then

                                'Se finaliza algun paro clase 1 o 2
                                If paroActual > 0 Then
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finalizo = 1, finalizo_accion = 'O' WHERE id = " & paroActual)
                                    If regsAfectados > 0 Then
                                        paroActual = 0
                                        reporteActual = 0
                                        claseParoActual = 0
                                        finaliza_sensor = False
                                        piezasWIP = 0
                                    End If
                                End If

                                reporteActual = general.Tables(0).Rows(0)!id
                                Dim agrupador = ValNull(general.Tables(0).Rows(0)!agrupador_2, "N")

                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, reporte, maquina, fecha, tiempo, inicio, inicia, desde, estado, estatus, area, tipo) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", 'ANDON " & Strings.Left(ValNull(general.Tables(0).Rows(0)!nombre, "A"), 44) & "', 3, " & reporteActual & ", " & captura!maquinaid & ", '" & Format(general.Tables(0).Rows(0)!fecha, "yyyy/MM/dd") & "', 0, 1, '" & Format(general.Tables(0).Rows(0)!fecha, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(general.Tables(0).Rows(0)!fecha, "yyyy/MM/dd HH:mm:ss") & "', 'C', 'A', " & general.Tables(0).Rows(0)!area & ", " & agrupador & ");UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', parada_desde  = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                paroCreado = regsAfectados > 0
                            End If
                        End If

                        If captura!oee_umbral_produccion > 0 And captura!oee_umbral_produccion <= DateAndTime.DateDiff(DateInterval.Second, ultimaReparacion, ultimo) And Not paroCreado Then

                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, maquina, fecha, tiempo, inicio, inicia, desde, estado, estatus) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", '" & traduccion(74) & "', 1, " & captura!maquinaid & ", '" & Format(fecha, "yyyy/MM/dd") & "', 0, 1, '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', 'C', 'A');UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', parada_desde  = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                        End If

                        If captura!oee_umbral_produccion > 0 And conProduccion And claseParoActual = 1 Then

                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finalizo = 1, hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & paroActual & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', parada_desde = NULL, estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                        End If

                        If finaliza_sensor And paroActual > 0 And conProduccion Then
                            If (piezasWIP - produccion) <= 0 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET wip_van = " & produccion & ", estado = 'F', finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finalizo = 1 WHERE id = " & paroActual & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET wip_paro = 'N', wip_contador = 0, wip_tiempo = 0, estatus = 'F', parada_desde = NULL, estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                            Else
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET wip_van = " & produccion & " WHERE id = " & paroActual)
                            End If

                        End If

                        cadSQL = "SELECT id FROM " & rutaBD & ".detalleparos WHERE maquina = " & captura!maquinaid & " AND TIME_TO_SEC(TIMEDIFF(desde, NOW())) BETWEEN 0 AND 1800 AND estatus = 'A' AND estado = 'L' ORDER BY desde LIMIT 1"
                        general = consultaSEL(cadSQL)
                        If general.Tables(0).Rows.Count > 0 Then
                            cadAdic = cadAdic & ", proximo_paro = " & general.Tables(0).Rows(0)!id
                        Else
                            cadAdic = cadAdic & ", proximo_paro = 0 "
                        End If



                        Dim rateEfecto = "N"

                        If captura!oee_parte_actual = 0 Then
                            rateEfecto = "N"
                        ElseIf rateAlto > 0 And uRate >= rateEquipo * (rateAlto / 100) Then
                            rateEfecto = "A"
                        ElseIf rateBajo > 0 And uRate <= rateEquipo * (rateBajo / 100) Then
                            rateEfecto = "B"
                        End If

                        Dim cadFecha = ", rate_mal_desde = NULL"
                        If rateEfecto <> "N" Then
                            If rateEfecto = "N" Then
                                conRateBajo = True
                            End If
                            If miResumen.Tables(0).Rows(0)!rate_efecto <> rateEfecto Or miResumen.Tables(0).Rows(0)!rate_mal_desde.Equals(System.DBNull.Value) Then
                                cadFecha = ", rate_mal_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            Else
                                cadFecha = ""
                            End If
                        End If

                        If hayTipo3 Then conRateBajo = Not funcionando


                        Dim cadFechaBaja = ", rate_tendencia_baja = NULL"
                        If rateEfecto = "B" Then
                            If uRate <= miResumen.Tables(0).Rows(0)!rate Then
                                If miResumen.Tables(0).Rows(0)!rate_tendencia_baja.Equals(System.DBNull.Value) Then
                                    cadFechaBaja = ", rate_tendencia_baja = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                                Else
                                    cadFechaBaja = ""
                                End If
                            End If
                        End If

                        Dim cadFechaAlta = ", rate_tendencia_alta = NULL"
                        If rateEfecto = "A" Then
                            If uRate >= miResumen.Tables(0).Rows(0)!rate Then
                                If miResumen.Tables(0).Rows(0)!rate_tendencia_alta.Equals(System.DBNull.Value) Then
                                    cadFechaAlta = ", rate_tendencia_alta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                                Else
                                    cadFechaAlta = ""
                                End If
                            End If
                        End If

                        'Calculos de los indicadores

                        If tiempoTranscurrido = 0 Then
                            DIS = 100
                        Else
                            DIS = (disponibilidad / tiempoTranscurrido * 100)
                        End If
                        If disponibilidad = 0 And produccion_tc = 0 Then
                            EFI = 0
                        ElseIf disponibilidad = 0 Then
                            EFI = 100
                        Else
                            EFI = produccion_tc / disponibilidad * 100
                        End If

                        If EFI > 100 Then
                            EFI = 100
                        End If
                        calidad = calidad + miResumen.Tables(0).Rows(0)!calidad
                        calidad_tc = calidad_tc + miResumen.Tables(0).Rows(0)!calidad_tc
                        buffer = buffer + miResumen.Tables(0).Rows(0)!buffer
                        If produccion_tc > 0 Then
                            FTQ = (produccion_tc - miResumen.Tables(0).Rows(0)!calidad_tc) / produccion_tc * 100
                        End If

                        If FTQ < 0 Then
                            FTQ = 0
                        End If
                        OEE = EFI * FTQ * DIS / 10000
                        OEE = Math.Round(OEE, 3)
                        EFI = Math.Round(EFI, 3)
                        DIS = Math.Round(DIS, 3)
                        FTQ = Math.Round(FTQ, 3)


                        Dim cadFechaFTQ = ", ftq_tendencia_baja = NULL"
                        If FTQ <= miResumen.Tables(0).Rows(0)!ftq And miResumen.Tables(0).Rows(0)!ftq <> 100 Then
                            If miResumen.Tables(0).Rows(0)!ftq_tendencia_baja.Equals(System.DBNull.Value) Then
                                cadFechaFTQ = ", ftq_tendencia_baja = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            Else
                                cadFechaFTQ = ""
                            End If
                        End If

                        Dim cadFechaDIS = ", dis_tendencia_baja = NULL"
                        If DIS <= miResumen.Tables(0).Rows(0)!dis And miResumen.Tables(0).Rows(0)!dis <> 100 Then
                            If miResumen.Tables(0).Rows(0)!dis_tendencia_baja.Equals(System.DBNull.Value) Then
                                cadFechaDIS = ", dis_tendencia_baja = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            Else
                                cadFechaDIS = ""
                            End If
                        End If



                        Dim cadFechaEFI = ", efi_tendencia_baja = NULL"
                        If EFI <= miResumen.Tables(0).Rows(0)!efi And miResumen.Tables(0).Rows(0)!efi <> 0 Then
                            If miResumen.Tables(0).Rows(0)!efi_tendencia_baja.Equals(System.DBNull.Value) Then
                                cadFechaEFI = ", efi_tendencia_baja = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            Else
                                cadFechaEFI = ""
                            End If
                        End If



                        Dim cadFechaOEE = ", oee_tendencia_baja = NULL"
                        If OEE <= miResumen.Tables(0).Rows(0)!oee And miResumen.Tables(0).Rows(0)!oee <> 0 Then
                            If miResumen.Tables(0).Rows(0)!oee_tendencia_baja.Equals(System.DBNull.Value) Then
                                cadFechaOEE = ", oee_tendencia_baja = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'"
                            Else
                                cadFechaOEE = ""
                            End If
                        End If

                        Dim cadOEE = ", oee_imagen = 0"
                        If OEE < miResumen.Tables(0).Rows(0)!oee And miResumen.Tables(0).Rows(0)!oee <> 0 Then
                            cadOEE = ", oee_imagen = 2, oee_efecto = 'B'"
                        ElseIf OEE > miResumen.Tables(0).Rows(0)!oee And miResumen.Tables(0).Rows(0)!oee <> 0 Then
                            cadOEE = ", oee_imagen = 1, oee_efecto = 'A'"
                        End If

                        If EFI < miResumen.Tables(0).Rows(0)!efi And miResumen.Tables(0).Rows(0)!efi <> 0 Then
                            cadOEE = cadOEE & ", efi_imagen = 2, efi_efecto = 'B'"
                        ElseIf EFI > miResumen.Tables(0).Rows(0)!efi And miResumen.Tables(0).Rows(0)!efi <> 0 Then
                            cadOEE = cadOEE & ", efi_imagen = 1, efi_efecto = 'A'"
                        Else
                            cadOEE = cadOEE & ", efi_imagen = 0"
                        End If

                        If DIS < miResumen.Tables(0).Rows(0)!dis And miResumen.Tables(0).Rows(0)!dis <> 0 Then
                            cadOEE = cadOEE & ", dis_imagen = 2, dis_efecto = 'B'"
                        ElseIf DIS > miResumen.Tables(0).Rows(0)!dis And miResumen.Tables(0).Rows(0)!dis <> 0 Then
                            cadOEE = cadOEE & ", dis_imagen = 1, dis_efecto = 'A'"
                        Else
                            cadOEE = cadOEE & ", dis_imagen = 0"
                        End If

                        If FTQ < miResumen.Tables(0).Rows(0)!ftq And miResumen.Tables(0).Rows(0)!ftq <> 0 Then
                            cadOEE = cadOEE & ", ftq_imagen = 2, ftq_efecto = 'B'"
                        ElseIf FTQ > miResumen.Tables(0).Rows(0)!ftq And miResumen.Tables(0).Rows(0)!ftq <> 0 Then
                            cadOEE = cadOEE & ", ftq_imagen = 1, ftq_efecto = 'A'"
                        Else
                            cadOEE = cadOEE & ", ftq_imagen = 0"
                        End If

                        'Antes

                        Dim cadAdicDetener = ""
                        If miResumen.Tables(0).Rows(0)!detener > 0 And paroActual = 0 Then
                            If ValNull(miResumen.Tables(0).Rows(0)!detener_estimado, "N") = 0 Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, maquina, fecha, tiempo, inicio, inicia, desde, estado, estatus, notas, finaliza_sensor, tipo, area) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", '" & ValNull(miResumen.Tables(0).Rows(0)!detener_paro, "A") & "', 2, " & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd") & "', 0, " & miResumen.Tables(0).Rows(0)!detener & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', 'C', 'A', '" & ValNull(miResumen.Tables(0).Rows(0)!detener_notas, "A") & "', '" & ValNull(miResumen.Tables(0).Rows(0)!detener_piezas, "A") & "', " & ValNull(miResumen.Tables(0).Rows(0)!detener_tipo, "N") & ", " & ValNull(miResumen.Tables(0).Rows(0)!detener_area, "N") & ");UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                cadAdicDetener = ", reanudar = 0, detener = 0, detener_piezas = 'N', detener_notas = NULL, detener_estimado = 0, detener_area = 0, detener_tipo = 0, detener_paro = NULL "
                            Else

                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, maquina, fecha, tiempo, inicio, inicia, desde, estado, estatus, notas, finaliza_sensor, tipo, area) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", '" & ValNull(miResumen.Tables(0).Rows(0)!detener_paro, "A") & "', 2, " & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd") & "', " & ValNull(miResumen.Tables(0).Rows(0)!detener_estimado, "N") & ", " & miResumen.Tables(0).Rows(0)!detener & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', 'C', 'A', '" & ValNull(miResumen.Tables(0).Rows(0)!detener_notas, "A") & "', '" & ValNull(miResumen.Tables(0).Rows(0)!detener_piezas, "A") & "', " & ValNull(miResumen.Tables(0).Rows(0)!detener_tipo, "N") & ", " & ValNull(miResumen.Tables(0).Rows(0)!detener_area, "N") & ");UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)
                                cadAdicDetener = ", reanudar = 0, detener = 0, detener_piezas = 'N', detener_notas = NULL, detener_estimado = 0, detener_area = 0, detener_tipo = 0, detener_paro = NULL "
                            End If
                        ElseIf miResumen.Tables(0).Rows(0)!reanudar > 0 And paroActual > 0 Then
                            tiempoParoCorte = tiempoValido(paroCorte, ultimo, captura!proceso)
                            tiempoParo = tiempoParo + tiempoParoCorte

                            Dim cadDatos = "UPDATE " & rutaBD & ".detalleparos SET estado = 'F', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', tipo = " & ValNull(miResumen.Tables(0).Rows(0)!detener_tipo, "N") & ", notas = '" & ValNull(miResumen.Tables(0).Rows(0)!detener_notas, "A") & "', area = " & ValNull(miResumen.Tables(0).Rows(0)!detener_area, "N") & ", paro = '" & ValNull(miResumen.Tables(0).Rows(0)!detener_paro, "A") & "', resultados = '" & ValNull(miResumen.Tables(0).Rows(0)!detener_resultados, "A") & "', finalizo_accion = 'M' WHERE id = " & miResumen.Tables(0).Rows(0)!paro_actual
                            If claseParoActual <> 2 And claseParoActual <> 1 Then
                                cadDatos = "UPDATE " & rutaBD & ".detalleparos SET estado = 'F', tiempo = " & tiempoValido(paroDesde, ultimo, captura!proceso) & ", finaliza = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', finalizo_accion = 'M' WHERE id = " & miResumen.Tables(0).Rows(0)!paro_actual
                            End If

                            regsAfectados = consultaACT(cadDatos & ";UPDATE " & rutaBD & ".detalleparos SET hasta = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & miResumen.Tables(0).Rows(0)!paro_actual & " AND clase <> 0;UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'F', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', alarmado_manual = 'N', ultima_reparacion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)

                            paroActual = 0
                            reporteActual = 0
                            claseParoActual = 0
                            ultimaReparacion = ultimo
                            cadAdicDetener = ", reanudar = 0, detener = 0, detener_piezas = 'N', detener_notas = NULL, detener_estimado = 0, detener_area = 0, detener_tipo = 0, detener_paro = NULL "
                        End If



                        If captura!paro_wip = "S" And miResumen.Tables(0).Rows(0)!wip_paro = "S" Then
                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, maquina, fecha, tiempo, inicio, inicia, desde, hasta, estado, estatus, notas, finaliza_sensor, tipo, area, wip_piezas) VALUES (" & captura!oee_parte_actual & ", " & captura!oee_lote_actual & ", " & config.Tables(0).Rows(0)!turno_oee & ", '" & traduccion(75) & "', 0, " & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd") & "', " & ValNull(miResumen.Tables(0).Rows(0)!wip_tiempo, "N") & ", 1, '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', '" & Format(DateAndTime.DateAdd(DateInterval.Second, miResumen.Tables(0).Rows(0)!wip_tiempo, ultimo), "yyyy/MM/dd HH:mm:ss") & "', 'L', 'A', '" & traduccion(76) & "', '" & IIf(miResumen.Tables(0).Rows(0)!wip_contador > 0, "S", "N") & "', " & tipo_change & ", " & area_change & ", " & miResumen.Tables(0).Rows(0)!wip_contador & ");UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET wip_paro = 'Y', estatus = 'D', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE equipo = " & captura!maquinaid)

                        End If

                        '

                        cadAdic = cadAdic & cadFecha & cadFechaBaja & cadFechaAlta & cadFechaFTQ & cadFechaDIS & cadFechaEFI & cadFechaOEE & cadOEE & cadAdicDetener

                        Dim otraCad As String = ""
                        Dim min = Format(ultimo, "mm")
                        If (min = 0 And miResumen.Tables(0).Rows(0)!tiempo_reinicio <> min) Then
                            otraCad = ", tiempo_reinicio = " & min

                            If Format(ultimo, "HHmm") = "0000" And reinicio = 5 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_5 = 'S', factor_iniciar = '" & Format(ultimo, "ddMMyy") & "' WHERE equipo = " & captura!maquinaid & " AND (factor_iniciar <> '" & Format(ultimo, "ddMMyy") & "' OR ISNULL(factor_iniciar))")
                            End If
                            If DateAndTime.Weekday(ultimo) = 2 And reinicio = 6 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_6 = 'S', factor_iniciar = '" & Format(ultimo, "ddMMyy") & "' WHERE equipo = " & captura!maquinaid & " AND (factor_iniciar <> '" & Format(ultimo, "ddMMyy") & "' OR ISNULL(factor_iniciar))")
                            End If
                            If Format(ultimo, "dd") = "01" And reinicio = 7 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_7 = 'S', factor_iniciar = '" & Format(ultimo, "ddMMyy") & "' WHERE equipo = " & captura!maquinaid & " AND (factor_iniciar <> '" & Format(ultimo, "ddMMyy") & "' OR ISNULL(factor_iniciar))")
                            End If
                            If Format(ultimo, "ddMM") = "0101" And reinicio = 8 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_8 = 'S', factor_iniciar = '" & Format(ultimo, "ddMMyy") & "' WHERE equipo = " & captura!maquinaid & " AND (factor_iniciar <> '" & Format(ultimo, "ddMMyy") & "' OR ISNULL(factor_iniciar))")
                            End If
                            If Format(ultimo, "mm") = "00" And reinicio = 3 Then
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_3 = 'S', factor_iniciar = '" & Format(ultimo, "HHmm") & "' WHERE equipo = " & captura!maquinaid & " AND (factor_iniciar <> '" & Format(ultimo, "HHmm") & "' OR ISNULL(factor_iniciar))")
                            End If

                        End If

                        If hayTipo3 Then
                            uRate = -100
                            rateActual = -100
                            rateEquipo = -100
                            objetivo = 0
                        End If

                        cadSQL = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET oee_minutos_rate = " & minutosRate & ", ultima_lectura = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', parte = " & captura!oee_parte_actual & ", tripulacion = " & captura!oee_tripulacion_actual & ", turno = " & captura!oee_turno_actual & ", produccion = " & produccion & ", produccion_tc = " & produccion_tc & ", calidad = " & calidad & ", calidad_tc = " & calidad_tc & ", buffer = buffer + " & buffer & ", norden = '" & captura!norden & "', nparte = '" & captura!nparte & "', nequipo = '" & captura!nequipo & "', ntripulacion = '" & captura!ntripulacion & "', referencia = '" & captura!referencia & "', nturno = '" & captura!nturno & "', rate_teorico = " & rateEquipo & ", rate_min = " & rateBajo & ", rate_max = " & rateAlto & ", objetivo = " & objetivo & ", van = " & van & ", rate = " & rateActual & cadAdic & ", rate_efecto = '" & rateEfecto & "', ultimo_rate = " & uRate & ", ratemed = '" & rateUnidad & "', esperadodis = " & eDIS & ", esperadoftq = " & eFTQ & ", esperadoefi = " & eEFI & ", esperadooee = " & eOEE & ", hoyos = '" & IIf(incluyeHoyos, "S", "N") & "', efi = " & EFI & ", dis = " & DIS & ", ftq = " & FTQ & ", oee = " & OEE & otraCad & ", paro_actual = " & paroActual & ", parosmostrar = parosmostrar + " & tiempoParo & cadPlaneado & ", corte = " & idCorte & " WHERE equipo = " & captura!maquinaid


                        If miResumen.Tables(0).Rows(0)!iniciar = "S" Or (miResumen.Tables(0).Rows(0)!iniciar_1 = "S" And reinicio = 1) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_2 = "S" And reinicio = 2) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_3 = "S" And reinicio = 3) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_4 = "S" And reinicio = 4) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_5 = "S" And reinicio = 5) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_6 = "S" And reinicio = 6) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_7 = "S" And reinicio = 7) _
                    Or (miResumen.Tables(0).Rows(0)!iniciar_8 = "S" And reinicio = 8) Then
                            cadSQL = cadSQL & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET produccion = 0, calidad = 0, produccion_tc = 0, calidad_tc = 0, buffer = 0, parosmostrar = 0, rate_mal_desde = NULL, rate_tendencia_baja = NULL, rate_tendencia_alta = NULL, ftq_tendencia_baja = NULL, oee_tendencia_baja = NULL, dis_tendencia_baja = NULL, efi_tendencia_baja = NULL, parada_desde = NULL, paro_actual = 0, ultima_reparacion = NULL, ultima_produccion = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', alarmado_manual = 'N', ultima_buffer = NULL, iniciar = 'N', iniciar_1 = 'N', iniciar_2 = 'N', iniciar_3 = 'N', iniciar_4 = 'N', iniciar_5 = 'N', iniciar_6 = 'N', iniciar_7 = 'N', iniciar_8 = 'N', fecha_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', planeado = 'N', estado_desde = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'" & IIf(ValNull(captura!oee_historico_rate_reiniciar, "A") = "1", ", desde_rate = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "'", "") & " WHERE equipo = " & captura!maquinaid
                        End If

                        regsAfectados = consultaACT(cadSQL & ";" & "UPDATE " & rutaBD & ".relacion_procesos_sensores SET ultima_lectura = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & captura!id & cadSQLParo)
                        Dim corte

                        Dim minutoEste = Format(ultimo, "mm")
                        Dim buscarTipo = False

                        cadSQL = "SELECT id FROM " & rutaBD & ".lecturas_resumen WHERE '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' >= desde AND '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' <= hasta  AND equipo = " & captura!maquinaid
                        corte = consultaSEL(cadSQL)

                        If corte.Tables(0).Rows.Count = 0 Then
                            'Se crean los 4 registros
                            buscarTipo = True
                            Dim minuto = 0
                            esaqui = True
                            For i = 1 To 4
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".lecturas_resumen (equipo, desde, hasta, hora) VALUES (" & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd HH:") & Format(minuto, "00") & ":00', '" & Format(ultimo, "yyyy/MM/dd HH:") & Format(minuto + 14, "00") & ":59', " & Format(ultimo, "HH") & ")")
                                minuto = minuto + 15


                            Next
                            Dim ultimo2 = ultimo
                            ultimo2 = Convert.ToDateTime(Format(ultimo2, "yyyy/MM/dd HH") & ":00:00")
                            For i = 1 To 88
                                ultimo2 = DateAndTime.DateAdd(DateInterval.Minute, -15, ultimo2)
                                Dim cadSQLINI = "SELECT id FROM " & rutaBD & ".lecturas_resumen WHERE desde = '" & Format(ultimo2, "yyyy/MM/dd HH:mm") & ":00' AND equipo = " & captura!maquinaid
                                corte = consultaSEL(cadSQLINI)
                                If corte.Tables(0).Rows.Count = 0 Then
                                    Dim fechaHasta = DateAndTime.DateAdd(DateInterval.Second, 899, ultimo2)
                                    regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".lecturas_resumen (equipo, desde, hasta, hora) VALUES (" & captura!maquinaid & ", '" & Format(ultimo2, "yyyy/MM/dd HH:mm") & ":00', '" & Format(fechaHasta, "yyyy/MM/dd HH:mm:ss") & "', " & Format(fechaHasta, "HH") & ")")

                                End If
                            Next
                            'Se recalcula el pasado
                            corte = consultaSEL(cadSQL)

                        End If

                        If captura!replanear = "S" Then
                            crearProgramacion(captura!maquinaid, captura!oee_lote_actual, captura!oee_parte_actual, captura!replanear_desde, captura!proceso, objetivo, van)
                        End If
                        If captura!compaginar = "S" Then
                            compaginar(captura!maquinaid, captura!oee_lote_actual, Strings.Left(captura!compaginar_desde, 10), Strings.Right(captura!compaginar_desde, 2))
                        End If



                        If (minutoEste Mod 5 = 0 And miResumen.Tables(0).Rows(0)!tiempo_corte <> minutoEste) Or buscarTipo Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas Set tiempo_corte = " & minutoEste & " WHERE equipo = " & captura!maquinaid)
                            Dim cadenaTipo = "Select * FROM " & rutaBD & ".lecturas_resumen WHERE equipo = " & captura!maquinaid & " ORDER BY desde DESC LIMIT 8"
                            Dim calcularTipo = consultaSEL(cadenaTipo)
                            If calcularTipo.Tables(0).Rows.Count > 0 Then
                                For Each tipoColor In calcularTipo.Tables(0).Rows
                                    If tipoColor!produccion > 0 And tipoColor!produccion >= tipoColor!bajorate And tipoColor!produccion >= tipoColor!paro Then
                                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas_resumen Set tipo = 1 WHERE id = " & tipoColor!id)
                                    ElseIf tipoColor!bajorate > 0 And tipoColor!bajorate >= tipoColor!paro And tipoColor!bajorate >= tipoColor!produccion And tipoColor!bajorate >= tipoColor!sinplan Then
                                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas_resumen Set tipo = 2 WHERE id = " & tipoColor!id)
                                    ElseIf tipoColor!paro > 0 And tipoColor!paro >= tipoColor!bajorate And tipoColor!paro >= tipoColor!produccion And tipoColor!paro >= tipoColor!sinplan Then
                                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas_resumen Set tipo = 3 WHERE id = " & tipoColor!id)
                                    End If
                                Next
                            End If
                        End If

                        'Antes

                        cadAdic = ""
                        If paroActual > 0 Then
                            cadAdic = "paro = paro + 1"
                        ElseIf conRateBajo Then
                            cadAdic = "bajorate = bajorate + 1"
                        ElseIf conProduccion Then
                            cadAdic = "produccion = produccion + 1"
                        ElseIf miResumen.Tables(0).Rows(0)!transcurrido = tiempoTranscurrido And miResumen.Tables(0).Rows(0)!transcurrido_pasar = 1 Then
                            cadAdic = "sinplan = sinplan + 1"
                        Else
                            cadAdic = "produccion = produccion + 1"
                        End If
                        If miResumen.Tables(0).Rows(0)!transcurrido <> tiempoTranscurrido Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas Set transcurrido = " & tiempoTranscurrido & ", transcurrido_pasar = 0 WHERE equipo = " & captura!maquinaid)
                        Else
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas Set transcurrido_pasar = 1 WHERE equipo = " & captura!maquinaid)
                        End If


                        Dim cadSeg = ""
                        Dim cadSeg2 = ""
                        cadSeg = "DELETE FROM " & rutaBD & ".piezasxminuto WHERE equipo = " & captura!maquinaid & " And fecha <= DATE_ADD('" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', INTERVAL -1 DAY);"
                        If produccion_seg > 0 Or tiempoParo > 0 Then
                            'Antes

                            Dim hhdmm = Format(ultimo, "HHmm")
                            general = consultaSEL("SELECT hhmm FROM " & rutaBD & ".piezasxminuto WHERE equipo = " & captura!maquinaid & " AND hhmm = '" & hhdmm & "'")
                            cadSeg2 = "INSERT INTO " & rutaBD & ".piezasxminuto VALUES(" & captura!maquinaid & ", '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "', " & produccion_seg & ", " & tiempoParo & ", '" & hhdmm & "');"


                            If general.Tables(0).Rows.Count > 0 Then
                                cadSeg2 = "UPDATE " & rutaBD & ".piezasxminuto SET produccion = produccion + " & produccion_seg & ", paro = paro + " & tiempoParo & " WHERE equipo = " & captura!maquinaid & " AND hhmm = '" & hhdmm & "';"
                            End If
                        End If



                        'Despues


                        If tiempoParo > 0 And paroActual > 0 Then
                            cadSeg = cadSeg & "UPDATE " & rutaBD & ".detalleparos SET corte = '" & Format(ultimo, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & paroActual & ";"
                        End If
                        regsAfectados = consultaACT(cadSeg & cadSeg2 & "UPDATE " & rutaBD & ".lecturas_resumen SET " & cadAdic & "  WHERE id = " & corte.Tables(0).Rows(0)!id)
                    ElseIf calidad > 0 Or buffer > 0 Then
                        cadSQL = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET calidad = calidad + " & calidad & ", calidad_tc = calidad_tc + " & calidad_tc & ", buffer = buffer + " & buffer & " WHERE equipo = " & captura!maquinaid
                        regsAfectados = consultaACT(cadSQL)
                    End If
                    TTotal = 0

                Next
                regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas SET estatus = 2 WHERE estatus = 1")
            End If

            regsAfectados = consultaACT("UPDATE " & rutaBD & ".configuracion Set lectura_pendiente = 0")
        End If
        revisandoSensores = False


    End Sub

    Private Sub cambioTurno_Tick(sender As Object, e As EventArgs) Handles cambioTurno.Tick
        If Not estadoPrograma Then Exit Sub
        If modulos(5) = 0 Then Exit Sub
        Dim regsAfectados As Long = 0
        Dim cadSQL As String = ""
        Dim ultimoHora As String = Format(Now, "HH:mm:ss")
        Dim nombreTurno As String = ""

        cadSQL = "SELECT turno_oee, oee_por_turno, turno_secuencia FROM " & rutaBD & ".configuracion"
        Dim config As DataSet = consultaSEL(cadSQL)
        If config.Tables(0).Rows.Count > 0 Then
            miTurno = -1
            cadSQL = "SELECT id, inicia, termina, nombre, usuario FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A' AND (inicia <= '" & ultimoHora & "' OR termina >= '" & ultimoHora & "') AND id <> " & config.Tables(0).Rows(0)!turno_oee & " ORDER BY inicia, termina"
            Dim horarios As DataSet = consultaSEL(cadSQL)
            If horarios.Tables(0).Rows.Count > 0 Then
                For Each horario In horarios.Tables(0).Rows
                    If ultimoHora >= horario!inicia.ToString And ultimoHora <= horario!termina.ToString Then
                        miTurno = horario!id
                        nombreTurno = ValNull(horario!nombre, "A")
                        operador = ValNull(horario!usuario, "A")
                        Exit For
                    End If
                Next
                If miTurno = -1 Then
                    For Each horario In horarios.Tables(0).Rows
                        If (ultimoHora >= horario!inicia.ToString Or ultimoHora <= horario!termina.ToString) And horario!termina.ToString < horario!inicia.ToString Then
                            miTurno = horario!id
                            nombreTurno = ValNull(horario!nombre, "A")
                            operador = ValNull(horario!usuario, "A")
                            Exit For
                        End If
                    Next

                End If
                If miTurno = -1 Then
                    miTurno = config.Tables(0).Rows(0)!turno_oee
                Else
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_turno_actual = " & miTurno & " WHERE oee = 'S';UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_4 = 'S', turno = " & miTurno & ", nturno = '" & nombreTurno & "';UPDATE " & rutaBD & ".configuracion SET turno_oee = " & miTurno & ", turno_secuencia = turno_secuencia + 1")
                    agregarSolo("Cambio de turno: " & nombreTurno)
                End If
            End If
            'If miTurno > 0 Then
            '    'Se calculan los indicadores por turno antes del cambio
            '    'If ValNull(config.Tables(0).Rows(0)!oee_por_turno, "A") = "S" Then
            '    Try
            '        'Enviar el mensaje de corte por turno (últimas 8 horas)
            '        'Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "oee_por_turno;" & config.Tables(0).Rows(0)!turno_secuencia & Chr(34), AppWinStyle.MinimizedNoFocus)

            '        Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "hxh_por_turno;" & config.Tables(0).Rows(0)!turno_secuencia & Chr(34), AppWinStyle.MinimizedNoFocus)

            '    Catch ex As Exception
            '    End Try
            '    'End If
            '    regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET oee_turno_actual = " & miTurno & " WHERE oee = 'S';UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET iniciar_4 = 'S', turno = " & miTurno & ", nturno = '" & nombreTurno & "';UPDATE " & rutaBD & ".configuracion SET turno_oee = " & miTurno & ", turno_secuencia = turno_secuencia + 1")
            '    agregarSolo("Cambio de turno: " & nombreTurno)
            'End If
            'cadSQL = "SELECT * FROM " & rutaBD & ".control WHERE fecha = '" & Format(Now, "yyyyMMddHH") & "' AND tipo = 1"
            'Dim readerDS As DataSet = consultaSEL(cadSQL)
            'If readerDS.Tables(0).Rows.Count = 0 Then
            '    Try
            '        'Enviar el mensaje de corte por turno (últimas 8 horas)
            '        Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "hxh_por_hora;" & Format(Now, "HH") & Chr(34), AppWinStyle.MinimizedNoFocus)
            '        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".control (fecha, tipo) VALUES ('" & Format(Now, "yyyyMMddHH") & "', 1)")

            '    Catch ex As Exception
            '    End Try
            'End If

            'cadSQL = "SELECT * FROM " & rutaBD & ".control WHERE fecha = '" & Format(Now, "yyyyMMdd") & "' AND tipo = 2"
            'readerDS = consultaSEL(cadSQL)
            'If readerDS.Tables(0).Rows.Count = 0 Then
            '    Try
            '        'Enviar el mensaje de corte por turno (últimas 8 horas)
            '        Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "hxh_por_dia;" & Format(Now, "yyyyMMddHH") & Chr(34), AppWinStyle.MinimizedNoFocus)
            '        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".control (fecha, tipo) VALUES ('" & Format(Now, "yyyyMMdd") & "', 2)")

            '    Catch ex As Exception

            '    End Try
            'End If



        End If
    End Sub

    Sub crearMensajes()
        Dim idProceso = Process.GetCurrentProcess.Id
        Dim mensajesDS As DataSet
        Dim registroDS As DataSet
        Dim eMensaje = ""
        Dim laLinea As String = ""
        Dim laMaquina As String = ""
        Dim laArea As String = ""
        Dim laFalla As String = ""
        Dim fecha
        Dim tiempo As String = ""
        Dim cadSQL As String = ""
        Dim nroReporte As Integer = 0
        Dim eTitulo = ""

        'Escalada 4
        Dim miError As String = ""
        Dim optimizar As Boolean = False
        Dim mantenerPrioridad As Boolean = False
        Dim regsAfectados = 0

        Dim maximo_largo_mmcall As Integer = 40



        regsAfectados = consultaACT("UPDATE " & rutaBD & ".mensajes SET estatus = '" & idProceso & "' WHERE estatus = 'A' AND alerta >= 0")
        cadSQL = "SELECT a.id, a.canal, b.evento, a.prioridad FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_alertas b on a.alerta = b.id WHERE a.estatus = '" & idProceso & "' ORDER BY a.prioridad DESC, a.id"
        'Se preselecciona la voz
        mensajesDS = consultaSEL(cadSQL)
        Dim generarMensaje As Boolean
        If mensajesDS.Tables(0).Rows.Count > 0 Then
            For Each elmensaje In mensajesDS.Tables(0).Rows
                generarMensaje = False
                eMensaje = ""
                laLinea = ""
                laMaquina = ""
                laArea = ""
                laFalla = ""
                tiempo = ""
                generarMensaje = True
                nroReporte = 0

                If elmensaje!evento < 200 Then
                    cadSQL = "SELECT a.*, 0 AS rate, 0 AS oee, e.nombre as nlinea, f.nombre as nmaquina, g.nombre as narea, h.nombre as nfalla, c.id AS idalerta, c.acumular, c.mensaje_mmcall, c.mensaje, c.titulo, c.resolucion_mensaje, c.cancelacion_mensaje, d.fecha, d.inicio_atencion, d.inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".reportes d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".cat_lineas e ON d.linea = e.id LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.maquina = f.id LEFT JOIN " & rutaBD & ".cat_areas g ON d.area = g.id LEFT JOIN " & rutaBD & ".cat_fallas h ON d.falla = h.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento < 204 Then
                    cadSQL = "SELECT a.*, IF(d.rate_teorico > 0, d.rate / d.rate_teorico * 100, 0) AS rate, d.oee, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.acumular, c.mensaje, c.titulo, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.rate_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.ultima_produccion AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 204 Then
                    cadSQL = "SELECT a.*, ftq AS rate, d.oee, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.acumular, c.mensaje, c.titulo, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.ftq_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.ultima_produccion AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 205 Then
                    cadSQL = "SELECT a.*, dis AS rate, d.oee, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.acumular, c.mensaje, c.titulo, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.dis_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.ultima_produccion AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 206 Then
                    cadSQL = "SELECT a.*, efi AS rate, d.oee, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.acumular, c.mensaje, c.titulo, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.efi_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.ultima_produccion AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 207 Then
                    cadSQL = "SELECT a.*, oee AS rate, d.oee, e.nombre as nlinea, f.nombre as nmaquina, '' as narea, '' as nfalla, c.id AS idalerta, c.acumular, c.mensaje, c.titulo, c.mensaje_mmcall, c.resolucion_mensaje, c.cancelacion_mensaje, d.oee_tendencia_baja AS fecha, d.rate_tendencia_alta AS inicio_atencion, d.ultima_produccion AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas d ON a.proceso = d.equipo LEFT JOIN " & rutaBD & ".cat_maquinas f ON d.equipo = f.id LEFT JOIN " & rutaBD & ".cat_lineas e ON f.linea = e.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 301 Then
                    cadSQL = "SELECT a.*, e1.referencia, e1.nombre AS producto, b1.numero AS nlote, IFNULL((SELECT MIN(orden) FROM " & rutaBD & ".prioridades WHERE parte = b1.parte AND fecha >= NOW() AND estatus = 'A'), 100) AS prioridad, d.ruta_secuencia, d.ruta_secuencia_antes, IFNULL(c1.nombre, '" & traduccion(145) & "') AS ruta_antes, IFNULL(d1.nombre, '" & traduccion(145) & "') AS ruta_despues, c.id AS idalerta, c.acumular, c.mensaje_mmcall, c.mensaje, c.titulo, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS fecha, NOW() AS inicio_atencion, NOW() AS inicio_reporte, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".lotes_historia d ON a.proceso = d.id INNER JOIN " & rutaBD & ".lotes b1 ON d.lote = b1.id LEFT JOIN " & rutaBD & ".det_rutas c1 ON d.ruta_detalle_anterior = c1.id LEFT JOIN " & rutaBD & ".det_rutas d1 ON d.ruta_detalle = d1.id LEFT JOIN " & rutaBD & ".cat_partes e1 ON b1.parte = e1.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 302 Or elmensaje!evento = 303 Or elmensaje!evento = 305 Or elmensaje!evento = 306 Then
                    cadSQL = "SELECT a.*, d.hasta, d.numero AS nlote, d.fecha, TIME_TO_SEC(TIMEDIFF(d.hasta, NOW())) AS previo, d.ruta_secuencia, c1.referencia, c1.nombre AS producto, IFNULL(b1.nombre, '" & traduccion(145) & "') AS ruta_actual, IFNULL(e1.nombre, '" & traduccion(145) & "') as equipo, IFNULL(d1.nombre, '" & traduccion(145) & "') as nproceso, c.id AS idalerta, c.acumular, c.mensaje_mmcall, c.mensaje, c.titulo, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS inicio_atencion, NOW() AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".lotes d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".det_rutas b1 ON d.ruta_detalle = b1.id LEFT JOIN " & rutaBD & ".cat_partes c1 ON d.parte = c1.id LEFT JOIN " & rutaBD & ".cat_procesos d1 ON d.proceso = d1.id LEFT JOIN " & rutaBD & ".cat_maquinas e1 ON d.equipo= e1.id WHERE a.id = " & elmensaje!id
                ElseIf elmensaje!evento = 304 Or elmensaje!evento = 307 Then
                    cadSQL = "SELECT a.*, 0 AS previo, d.carga, d.alarma, d.alarma_rep, d.fecha, d.permitir_reprogramacion, d.equipo, d.fecha, IFNULL(b1.nombre, '" & traduccion(145) & "') as nequipo, IFNULL(c1.nombre, '" & traduccion(145) & "') as nproceso, IFNULL((SELECT SUM(cantidad) FROM " & rutaBD & ".programacion WHERE carga = d.id AND estatus = 'A'), 0) AS piezas, (SELECT COUNT(*) FROM " & rutaBD & ".lotes WHERE carga = d.id) AS avance, c.id AS idalerta, c.acumular, c.mensaje_mmcall, c.mensaje, c.titulo, c.resolucion_mensaje, c.cancelacion_mensaje, i.inicio AS inicio_atencion, NOW() AS inicio_reporte, d.estatus, i.repeticiones, i.fase, i.escalamientos1, i.escalamientos2, i.escalamientos3, i.escalamientos4, i.escalamientos5 FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_distribucion b ON a.lista = b.id AND b.estatus = 'A' INNER JOIN " & rutaBD & ".cat_alertas c ON a.alerta = c.id INNER JOIN " & rutaBD & ".alarmas i ON a.alarma = i.id LEFT JOIN " & rutaBD & ".cargas d ON a.proceso = d.id LEFT JOIN " & rutaBD & ".cat_maquinas b1 ON d.equipo = b1.id AND b1.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_procesos c1 ON b1.proceso = c1.id AND c1.estatus = 'A' WHERE a.id = " & elmensaje!id
                End If

                registroDS = consultaSEL(cadSQL)
                If registroDS.Tables(0).Rows.Count > 0 Then
                    nroReporte = registroDS.Tables(0).Rows(0)!proceso
                    If elmensaje!evento < 300 Then
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

                    If elmensaje!evento = 101 Or elmensaje!evento = 201 Or elmensaje!evento = 301 Or (elmensaje!evento >= 204 And elmensaje!evento <= 207) Then
                        fecha = registroDS.Tables(0).Rows(0)!fecha
                    ElseIf elmensaje!evento = 102 Or elmensaje!evento = 202 Or elmensaje!evento > 300 Then
                        fecha = registroDS.Tables(0).Rows(0)!inicio_atencion
                    ElseIf elmensaje!evento = 103 Or elmensaje!evento = 203 Then
                        fecha = registroDS.Tables(0).Rows(0)!inicio_reporte
                    End If
                    tiempo = calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, fecha, Now))
                    Dim mIFormato = "dd-MMM-yyyy HH:mm:ss"

                    If registroDS.Tables(0).Rows(0)!tipo = 8 Then
                        eMensaje = ValNull(registroDS.Tables(0).Rows(0)!resolucion_mensaje, "A")
                    ElseIf registroDS.Tables(0).Rows(0)!tipo = 7 Then
                        eMensaje = ValNull(registroDS.Tables(0).Rows(0)!cancelacion_mensaje, "A")
                    Else
                        If elmensaje!canal = 3 Then
                            mIFormato = "dd/MM HH:mm"
                            eMensaje = ValNull(registroDS.Tables(0).Rows(0)!mensaje_mmcall, "A")
                        Else
                            eMensaje = ValNull(registroDS.Tables(0).Rows(0)!mensaje, "A")
                        End If

                    End If

                    If eMensaje.Length > 0 Then
                        eMensaje = Replace(eMensaje, "[0]", nroReporte)
                        eMensaje = Replace(eMensaje, "[1]", laLinea)
                        eMensaje = Replace(eMensaje, "[2]", laMaquina)
                        eMensaje = Replace(eMensaje, "[3]", laArea)
                        eMensaje = Replace(eMensaje, "[4]", laFalla)
                        eMensaje = Replace(eMensaje, "[5]", Format(fecha, mIFormato))
                        eMensaje = Replace(eMensaje, "[11]", tiempo)
                        If elmensaje!evento < 300 Then
                            eMensaje = Replace(eMensaje, "[12]", Format(registroDS.Tables(0).Rows(0)!rate, "0.0"))
                            eMensaje = Replace(eMensaje, "[13]", Format(registroDS.Tables(0).Rows(0)!oee, "0.0"))

                        End If
                        If ValNull(registroDS.Tables(0).Rows(0)!fase - 10, "N") <= 0 Then
                            If ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N") > 0 Then
                                eMensaje = Replace(eMensaje, "[20]", traduccion(94) & ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N"))
                            Else
                                eMensaje = Replace(eMensaje, "[20]", "")
                            End If
                            eMensaje = Replace(eMensaje, "[30]", "")
                        ElseIf ValNull(registroDS.Tables(0).Rows(0)!fase - 10, "N") > 0 Then
                            Dim escala = ValNull(registroDS.Tables(0).Rows(0)!fase, "N") - 10
                            eMensaje = Replace(eMensaje, "[30]", traduccion(95) & If(escala > 0, escala, 0))
                            Dim repeticiones = 0
                            If escala = 1 Then
                                repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos1, "N")
                            ElseIf escala = 2 Then
                                repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos2, "N")
                            ElseIf escala = 3 Then
                                repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos3, "N")
                            ElseIf escala = 4 Then
                                repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos4, "N")
                            ElseIf escala = 5 Then
                                repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos5, "N")
                            End If
                            If repeticiones > 0 Then
                                eMensaje = Replace(eMensaje, "[20]", "Rep " & repeticiones)
                            Else
                                eMensaje = Replace(eMensaje, "[20]", "")
                            End If
                        Else
                            eMensaje = Replace(eMensaje, "[30]", "")
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
                            eMensaje = Replace(eMensaje, "[50]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                            eMensaje = Replace(eMensaje, "[51]", Format(registroDS.Tables(0).Rows(0)!hasta, mIFormato))
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
                            eMensaje = Replace(eMensaje, "[62]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                            eMensaje = Replace(eMensaje, "[63]", Format(registroDS.Tables(0).Rows(0)!hasta, mIFormato))
                            eMensaje = Replace(eMensaje, "[64]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!hasta, DateAndTime.Now)))
                            eMensaje = Replace(eMensaje, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                        ElseIf elmensaje!evento = 304 Or elmensaje!evento = 307 Then
                            eMensaje = Replace(eMensaje, "[80]", ValNull(registroDS.Tables(0).Rows(0)!carga, "A"))
                            eMensaje = Replace(eMensaje, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                            eMensaje = Replace(eMensaje, "[61]", ValNull(registroDS.Tables(0).Rows(0)!nequipo, "A"))
                            eMensaje = Replace(eMensaje, "[81]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                            eMensaje = Replace(eMensaje, "[82]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!fecha, DateAndTime.Now)))
                            eMensaje = Replace(eMensaje, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                            eMensaje = Replace(eMensaje, "[84]", ValNull(registroDS.Tables(0).Rows(0)!texto, "A"))
                        End If
                    Else
                        If elmensaje!evento = 101 Then
                            eMensaje = traduccion(80) & " " & nroReporte & traduccion(77)
                        ElseIf elmensaje!evento = 102 Then
                            eMensaje = traduccion(80) & " " & nroReporte & traduccion(78)
                        ElseIf elmensaje!evento = 103 Then
                            eMensaje = traduccion(80) & " " & nroReporte & traduccion(79)
                        ElseIf elmensaje!evento = 201 Then
                            eMensaje = traduccion(81) & laMaquina
                        ElseIf elmensaje!evento = 202 Then
                            eMensaje = traduccion(82) & laMaquina
                        ElseIf elmensaje!evento = 203 Then
                            eMensaje = laMaquina & " " & traduccion(74)
                        ElseIf elmensaje!evento = 204 Then
                            eMensaje = traduccion(83) & laMaquina
                        ElseIf elmensaje!evento = 205 Then
                            eMensaje = traduccion(84) & laMaquina
                        ElseIf elmensaje!evento = 206 Then
                            eMensaje = traduccion(85) & laMaquina
                        ElseIf elmensaje!evento = 207 Then
                            eMensaje = traduccion(86) & laMaquina
                        ElseIf elmensaje!evento = 301 Then
                            eMensaje = traduccion(87)
                        ElseIf elmensaje!evento = 302 Then
                            eMensaje = traduccion(88)
                        ElseIf elmensaje!evento = 303 Then
                            eMensaje = traduccion(89)
                        ElseIf elmensaje!evento = 304 Then
                            eMensaje = traduccion(90)
                        ElseIf elmensaje!evento = 305 Then
                            eMensaje = traduccion(91)
                        ElseIf elmensaje!evento = 306 Then
                            eMensaje = traduccion(92)
                        ElseIf elmensaje!evento = 307 Then
                            eMensaje = traduccion(93)
                        End If
                    End If

                    If elmensaje!canal = 2 Then
                        If eTitulo.Length = 0 Then
                            eTitulo = ValNull(registroDS.Tables(0).Rows(0)!titulo, "A")
                        End If
                        If eTitulo.Length > 0 Then
                            eTitulo = Replace(eTitulo, "[0]", nroReporte)
                            eTitulo = Replace(eTitulo, "[1]", laLinea)
                            eTitulo = Replace(eTitulo, "[2]", laMaquina)
                            eTitulo = Replace(eTitulo, "[3]", laArea)
                            eTitulo = Replace(eTitulo, "[4]", laFalla)
                            eTitulo = Replace(eTitulo, "[5]", Format(fecha, "ddd, dd-MMM-yyyy HH:mm"))
                            eTitulo = Replace(eTitulo, "[11]", tiempo)
                            If elmensaje!evento <= 300 Then
                                eTitulo = Replace(eTitulo, "[12]", Format(registroDS.Tables(0).Rows(0)!rate, "0") & "%")
                                eTitulo = Replace(eTitulo, "[13]", Format(registroDS.Tables(0).Rows(0)!oee, "0") & "%")
                            End If

                            If ValNull(registroDS.Tables(0).Rows(0)!fase - 10, "N") <= 0 Then
                                If ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N") > 0 Then
                                    eTitulo = Replace(eTitulo, "[20]", traduccion(94) & ValNull(registroDS.Tables(0).Rows(0)!repeticiones, "N"))
                                Else
                                    eTitulo = Replace(eTitulo, "[20]", "")
                                End If
                                eTitulo = Replace(eTitulo, "[30]", "")
                            ElseIf ValNull(registroDS.Tables(0).Rows(0)!fase - 10, "N") > 0 Then
                                Dim escala = ValNull(registroDS.Tables(0).Rows(0)!fase, "N") - 10
                                eTitulo = Replace(eTitulo, "[30]", traduccion(95) & If(escala > 0, escala, 0))
                                Dim repeticiones = 0
                                If escala = 1 Then
                                    repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos1, "N")
                                ElseIf escala = 2 Then
                                    repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos2, "N")
                                ElseIf escala = 3 Then
                                    repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos3, "N")
                                ElseIf escala = 4 Then
                                    repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos4, "N")
                                ElseIf escala = 5 Then
                                    repeticiones = ValNull(registroDS.Tables(0).Rows(0)!escalamientos5, "N")
                                End If
                                If repeticiones > 0 Then
                                    eTitulo = Replace(eTitulo, "[20]", "Rep " & repeticiones)
                                Else
                                    eTitulo = Replace(eTitulo, "[20]", "")
                                End If
                            Else
                                eTitulo = Replace(eTitulo, "[30]", "")
                            End If

                            eTitulo = Replace(eTitulo, "[90]", "")
                            eTitulo = Replace(eTitulo, System.Environment.NewLine, " ")

                            If elmensaje!evento = 301 Then
                                eTitulo = Replace(eTitulo, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                eTitulo = Replace(eTitulo, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                eTitulo = Replace(eTitulo, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                eTitulo = Replace(eTitulo, "[70]", ValNull(registroDS.Tables(0).Rows(0)!ruta_antes, "A"))
                                eTitulo = Replace(eTitulo, "[71]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia_antes, "A"))
                                eTitulo = Replace(eTitulo, "[72]", ValNull(registroDS.Tables(0).Rows(0)!ruta_despues, "A"))
                                eTitulo = Replace(eTitulo, "[73]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))

                            ElseIf elmensaje!evento = 302 Or elmensaje!evento = 305 Then
                                eTitulo = Replace(eTitulo, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                eTitulo = Replace(eTitulo, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                eTitulo = Replace(eTitulo, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                eTitulo = Replace(eTitulo, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                eTitulo = Replace(eTitulo, "[44]", ValNull(registroDS.Tables(0).Rows(0)!ruta_actual, "A"))
                                eTitulo = Replace(eTitulo, "[45]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))
                                eTitulo = Replace(eTitulo, "[50]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                                eTitulo = Replace(eTitulo, "[51]", Format(registroDS.Tables(0).Rows(0)!hasta, mIFormato))
                                eTitulo = Replace(eTitulo, "[52]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!hasta, DateAndTime.Now)))
                                eTitulo = Replace(eTitulo, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                            ElseIf elmensaje!evento = 303 Or elmensaje!evento = 306 Then
                                eTitulo = Replace(eTitulo, "[41]", ValNull(registroDS.Tables(0).Rows(0)!nlote, "A"))
                                eTitulo = Replace(eTitulo, "[42]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                                eTitulo = Replace(eTitulo, "[43]", ValNull(registroDS.Tables(0).Rows(0)!producto, "A"))
                                eTitulo = Replace(eTitulo, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                eTitulo = Replace(eTitulo, "[44]", ValNull(registroDS.Tables(0).Rows(0)!ruta_actual, "A"))
                                eTitulo = Replace(eTitulo, "[45]", ValNull(registroDS.Tables(0).Rows(0)!ruta_secuencia, "A"))
                                eTitulo = Replace(eTitulo, "[61]", ValNull(registroDS.Tables(0).Rows(0)!equipo, "A"))
                                eTitulo = Replace(eTitulo, "[62]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                                eTitulo = Replace(eTitulo, "[63]", Format(registroDS.Tables(0).Rows(0)!hasta, mIFormato))
                                eTitulo = Replace(eTitulo, "[64]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!hasta, DateAndTime.Now)))
                                eTitulo = Replace(eTitulo, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                            ElseIf elmensaje!evento = 304 Or elmensaje!evento = 307 Then
                                eTitulo = Replace(eTitulo, "[80]", ValNull(registroDS.Tables(0).Rows(0)!carga, "A"))
                                eTitulo = Replace(eTitulo, "[40]", ValNull(registroDS.Tables(0).Rows(0)!nproceso, "A"))
                                eTitulo = Replace(eTitulo, "[61]", ValNull(registroDS.Tables(0).Rows(0)!nequipo, "A"))
                                eTitulo = Replace(eTitulo, "[81]", Format(registroDS.Tables(0).Rows(0)!fecha, mIFormato))
                                eTitulo = Replace(eTitulo, "[82]", calcularTiempoCad(DateAndTime.DateDiff(DateInterval.Second, registroDS.Tables(0).Rows(0)!fecha, DateAndTime.Now)))
                                eTitulo = Replace(eTitulo, "[83]", calcularTiempoCad(registroDS.Tables(0).Rows(0)!previo))
                                eTitulo = Replace(eTitulo, "[84]", ValNull(registroDS.Tables(0).Rows(0)!texto, "A"))

                            End If
                        End If

                    ElseIf elmensaje!canal = 3 Then
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
                    End If
                    eMensaje = Strings.Left(eMensaje, 400)
                    eTitulo = Strings.Left(eTitulo, 100)
                End If
                Dim cadAdic = ""
                If eMensaje.Length > 0 Then
                    cadAdic = "INSERT INTO " & rutaBD & ".mensajes_procesados (texto, canal, titulo, prioridad, fecha, mensaje) VALUES ('" & eMensaje & "', " & elmensaje!canaL & ", '" & eTitulo & "', " & elmensaje!prioridad & ", NOW(), " & elmensaje!id & ");"
                End If
                regsAfectados = consultaACT(cadAdic & "UPDATE " & rutaBD & ".mensajes SET estatus = 'E', enviada = NOW() WHERE estatus = '" & idProceso & "'")
            Next
        End If
    End Sub

    Sub crearMensajesDOS()
        Dim idProceso = Process.GetCurrentProcess.Id
        Dim mensajesDS As DataSet
        Dim registroDS As DataSet
        Dim eMensaje = ""
        Dim cadSQL As String = ""
        Dim eTitulo = ""

        'Escalada 4
        Dim regsAfectados = 0

        Dim maximo_largo_mmcall As Integer = 40

        regsAfectados = consultaACT("UPDATE " & rutaBD & ".mensajes SET estatus = '" & idProceso & "' WHERE estatus = 'A' AND alerta < 0")
        cadSQL = "SELECT id, canal, 0 AS evento, 0 AS prioridad FROM " & rutaBD & ".mensajes WHERE estatus = '" & idProceso & "' ORDER BY id"
        'Se preselecciona la voz
        mensajesDS = consultaSEL(cadSQL)
        If mensajesDS.Tables(0).Rows.Count > 0 Then
            Dim cadAdic = ""
            For Each elmensaje In mensajesDS.Tables(0).Rows
                eMensaje = ""
                cadSQL = "SELECT a.*, c.nombre AS nequipo, e.nombre AS ntipo, d.mensaje, d.titulo, d.mensaje_mmcall, d.mensaje2, d.titulo2, d.mensaje_mmcall2, b.nombre, b.referencia, b.tiempo, d.hora FROM " & rutaBD & ".mensajes a INNER JOIN " & rutaBD & ".cat_checklists b ON a.proceso = b.id INNER JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id INNER JOIN " & rutaBD & ".cat_generales e ON a.tipo = e.id INNER JOIN " & rutaBD & ".plan_checklists d ON a.prioridad = d.id WHERE a.id = " & elmensaje!id
                registroDS = consultaSEL(cadSQL)
                If registroDS.Tables(0).Rows.Count > 0 Then
                    Dim elTiempo = ""
                    If registroDS.Tables(0).Rows(0)!tiempo > 0 Then
                        elTiempo = Format(registroDS.Tables(0).Rows(0)!tiempo / 60, "0")
                    Else
                        elTiempo = traduccion(184)
                    End If

                    If registroDS.Tables(0).Rows(0)!canal = 3 Then
                        eMensaje = IIf(registroDS.Tables(0).Rows(0)!alerta = 0, ValNull(registroDS.Tables(0).Rows(0)!mensaje_mmcall, "A"), ValNull(registroDS.Tables(0).Rows(0)!mensaje_mmcall2, "A"))
                    Else
                        eMensaje = IIf(registroDS.Tables(0).Rows(0)!alerta = 0, ValNull(registroDS.Tables(0).Rows(0)!mensaje, "A"), ValNull(registroDS.Tables(0).Rows(0)!mensaje2, "A"))
                    End If
                    If eMensaje.Length = 0 Then
                        eMensaje = "Se va a generar el checklist " & registroDS.Tables(0).Rows(0)!nombre & " a las " & registroDS.Tables(0).Rows(0)!hora.ToString
                    Else
                        eMensaje = Replace(eMensaje, "[0]", registroDS.Tables(0).Rows(0)!hora.ToString)
                        eMensaje = Replace(eMensaje, "[1]", ValNull(registroDS.Tables(0).Rows(0)!nombre, "A"))
                        eMensaje = Replace(eMensaje, "[2]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                        eMensaje = Replace(eMensaje, "[3]", ValNull(registroDS.Tables(0).Rows(0)!nequipo, "A"))
                        eMensaje = Replace(eMensaje, "[4]", ValNull(registroDS.Tables(0).Rows(0)!ntipo, "A"))
                        eMensaje = Replace(eMensaje, "[5]", elTiempo)
                    End If
                    If registroDS.Tables(0).Rows(0)!canal = 2 Then
                        eTitulo = IIf(registroDS.Tables(0).Rows(0)!alerta = 0, ValNull(registroDS.Tables(0).Rows(0)!titulo, "A"), ValNull(registroDS.Tables(0).Rows(0)!titulo2, "A"))
                        If eTitulo.Length = 0 Then
                            eTitulo = "Se va a generar el checklist " & registroDS.Tables(0).Rows(0)!nombre & " a la hora: " & registroDS.Tables(0).Rows(0)!hora.ToString
                        Else
                            eTitulo = Replace(eTitulo, "[0]", registroDS.Tables(0).Rows(0)!hora.ToString)
                            eTitulo = Replace(eTitulo, "[1]", ValNull(registroDS.Tables(0).Rows(0)!nombre, "A"))
                            eTitulo = Replace(eTitulo, "[2]", ValNull(registroDS.Tables(0).Rows(0)!referencia, "A"))
                            eTitulo = Replace(eTitulo, "[3]", ValNull(registroDS.Tables(0).Rows(0)!nequipo, "A"))
                            eTitulo = Replace(eTitulo, "[4]", ValNull(registroDS.Tables(0).Rows(0)!ntipo, "A"))
                            eTitulo = Replace(eTitulo, "[5]", elTiempo)
                        End If
                    End If
                    If registroDS.Tables(0).Rows(0)!canal = 3 Then
                        Dim antes As String = "ÃÀÁÄÂÈÉËÊÌÍÏÎÒÓÖÔÙÚÜÛãàáäâèéëêìíïîòóöôùúüûÑñÇç"
                        Dim ahora As String = "AAAAAEEEEIIIIOOOOUUUUaaaaaeeeeiiiioooouuuunncc"
                        For i = 0 To antes.Length - 1
                            eMensaje = Replace(eMensaje, antes(i), ahora(i))
                        Next
                        eMensaje = Replace(eMensaje, ";", " ")
                        eMensaje = Replace(eMensaje, "\", "-")
                        eMensaje = Replace(eMensaje, "/", "-")
                        eMensaje = Replace(eMensaje, System.Environment.NewLine, " ")
                        eMensaje = Strings.Left(eMensaje, 40)

                        'Se cambian los caracteres especiales
                    End If
                    eMensaje = Strings.Left(eMensaje, 400)
                    eTitulo = Strings.Left(eTitulo, 100)
                End If
                eMensaje.Trim()
                If eMensaje.Length > 0 Then
                    cadAdic = cadAdic & "INSERT INTO " & rutaBD & ".mensajes_procesados (texto, canal, titulo, prioridad, fecha, mensaje) VALUES ('" & eMensaje & "', " & elmensaje!canal & ", '" & eTitulo & "', 0, NOW(), " & elmensaje!id & ");"
                End If
            Next
            regsAfectados = consultaACT(cadAdic & "UPDATE " & rutaBD & ".mensajes SET estatus = 'E', enviada = NOW() WHERE estatus = '" & idProceso & "'")
        End If
    End Sub


    Sub paseaStock()
        If modulos(4) = 0 Or Not estadoPrograma Then Exit Sub

        Dim aCortar = False
        Dim TTotal As Long = 0
        Dim cadSQL As String = "SELECT IFNULL((SELECT MIN(orden) FROM " & rutaBD & ".prioridades WHERE parte = a.parte AND fecha >= NOW() AND estatus = 'A'), 100) AS prioridad, a.id, a.estado, a.proceso, a.ruta_detalle, a.fecha, a.calcular_hasta, b.tiempo_stock, c.ultimo_parte, c.paro_wip, a.parte, a.equipo FROM " & rutaBD & ".lotes a LEFT JOIN " & rutaBD & ".det_rutas b ON a.ruta_detalle = b.id LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id WHERE (a.estado = 0 or a.calcular_hasta <> 'N') AND a.estatus = 'A' ORDER BY prioridad, a.fecha ASC"
        Dim reader As DataSet = consultaSEL(cadSQL)
        Dim pases = 0
        Dim realculoStock = 0
        Dim realculoEquipo = 0

        Dim regsAfectados = 0
        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + Microsoft.VisualBasic.Strings.Left(errorBD, 250), 9, 0)
        Else
            If reader.Tables(0).Rows.Count > 0 Then
                For Each lotes In reader.Tables(0).Rows
                    If lotes!id = 1459 Then
                        Dim uno = 1
                    End If
                    If lotes!calcular_hasta = "N" Or lotes!calcular_hasta = "1" Or lotes!calcular_hasta = "3" Then
                        cadSQL = "SELECT capacidad_stock - (SELECT COUNT(*) FROM " & rutaBD & ".lotes WHERE estado = 20 And proceso = " & lotes!proceso & " AND estatus = 'A') AS capstock FROM " & rutaBD & ".cat_procesos WHERE id = " & lotes!proceso
                        Dim procesos As DataSet = consultaSEL(cadSQL)
                        If procesos.Tables(0).Rows.Count > 0 Then
                            If procesos.Tables(0).Rows(0)!capstock > 0 And (lotes!calcular_hasta <> "3" Or lotes!fecha <= Now()) Then
                                Dim loteID = lotes!id
                                Dim procesoID = lotes!proceso
                                Dim completado = False
                                Dim fechaEstimada = Now()
                                Dim FHasta = calcularFechaEstimada(lotes!fecha, lotes!tiempo_stock, lotes!proceso)
                                If lotes!estado = 0 Then
                                    pases = pases + 1
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET estado = 20, fecha = NOW(), calcular_hasta = 'N', hasta = '" & Format(FHasta, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & lotes!id & ";UPDATE " & rutaBD & ".lotes_historia SET fecha_stock = NOW(), tiempo_espera = TIME_TO_SEC(TIMEDIFF(NOW(), fecha_entrada)) WHERE lote = " & lotes!id & " AND proceso = " & lotes!proceso)
                                Else
                                    realculoStock = realculoStock + 1
                                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET hasta = '" & Format(FHasta, "yyyy/MM/dd HH:mm:ss") & "', calcular_hasta = 'N' WHERE id = " & lotes!id)
                                End If

                            End If
                        End If
                    ElseIf lotes!calcular_hasta = "2" And lotes!estado = 50 Then
                        cadSQL = "SELECT a.tiempo_proceso, a.tiempo_setup, tiempo_setup_idem, piezas_finalizar_paro FROM " & rutaBD & ".det_rutas a LEFT JOIN " & rutaBD & ".cat_procesos b ON a.proceso = b.id WHERE a.id = " & lotes!ruta_detalle
                        Dim procesos As DataSet = consultaSEL(cadSQL)
                        If procesos.Tables(0).Rows.Count > 0 Then
                            pases = pases + 1
                            Dim loteID = lotes!id
                            Dim procesoID = lotes!proceso
                            Dim completado = False
                            Dim fechaEstimada = Now()
                            Dim tiempo_sumar = ValNull(procesos.Tables(0).Rows(0)!tiempo_proceso, "N")
                            If ValNull(lotes!ultimo_parte, "N") <> lotes!parte Then
                                tiempo_sumar = tiempo_sumar + procesos.Tables(0).Rows(0)!tiempo_setup
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET ultimo_parte = " & lotes!parte & " WHERE id = " & lotes!equipo & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET desde_rate = NOW() WHERE equipo = " & lotes!equipo)
                            ElseIf ValNull(lotes!ultimo_parte, "N") = lotes!parte Then
                                tiempo_sumar = tiempo_sumar + procesos.Tables(0).Rows(0)!tiempo_setup_idem
                            End If
                            If modulos(5) = 1 Then
                                Dim validarParoWIP = False
                                Dim tiempoParo = 0
                                If lotes!paro_wip = "S" And ValNull(lotes!ultimo_parte, "N") = lotes!parte And procesos.Tables(0).Rows(0)!tiempo_setup_idem > 0 Then
                                    tiempoParo = procesos.Tables(0).Rows(0)!tiempo_setup_idem
                                    validarParoWIP = True
                                ElseIf lotes!paro_wip = "S" And ValNull(lotes!ultimo_parte, "N") <> lotes!parte And procesos.Tables(0).Rows(0)!tiempo_setup > 0 Then
                                    tiempoParo = procesos.Tables(0).Rows(0)!tiempo_setup
                                    validarParoWIP = True
                                End If
                                If validarParoWIP Then
                                    'Se evalua si el equipo está vacío
                                    cadSQL = "SELECT COUNT(*) as tlotes FROM " & rutaBD & ".lotes WHERE estatus = 'A' AND estado = 50 AND equipo = " & lotes!equipo
                                    Dim bEquipo As DataSet = consultaSEL(cadSQL)
                                    If bEquipo.Tables(0).Rows.Count > 0 Then
                                        If bEquipo.Tables(0).Rows(0)!tlotes = 1 Then
                                            'Se solicita un paro no programado por change over
                                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET wip_paro = 'S', wip_contador = " & procesos.Tables(0).Rows(0)!piezas_finalizar_paro & ", wip_tiempo = " & tiempoParo & " WHERE equipo = " & lotes!equipo)
                                        End If
                                    End If
                                End If
                            End If
                            realculoEquipo = realculoEquipo + 1
                            Dim FHasta = calcularFechaEstimada(lotes!fecha, tiempo_sumar, lotes!proceso)
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET hasta = '" & Format(FHasta, "yyyy/MM/dd HH:mm:ss") & "', calcular_hasta = 'N' WHERE id = " & lotes!id)
                        End If

                    End If

                Next
            End If
            If pases > 0 Then agregarSolo(traduccion(96) & pases & traduccion(97))
            If realculoStock > 0 Then agregarSolo(traduccion(98) & realculoStock & traduccion(99))
            If realculoEquipo > 0 Then agregarSolo(traduccion(100) & realculoEquipo & traduccion(99))
            If pases > 0 Then agregarLOG(traduccion(96) & pases & traduccion(101), 1, 1)
            'Se actualiza el catálogo de máquinas para indicarle si aún hay productos en la máquina

            cadSQL = "SELECT a.id, b.paro_actual, a.paro_wip FROM " & rutaBD & ".cat_maquinas a LEFT JOIN " & rutaBD & ".relacion_maquinas_lecturas b ON a.id = b.equipo WHERE a.oee = 'S' AND a.id NOT IN (SELECT equipo FROM " & rutaBD & ".lotes WHERE estatus = 'A' AND estado = 50) AND a.oee_parte_actual <> 0"
            reader = consultaSEL(cadSQL)
            If reader.Tables(0).Rows.Count > 0 Then
                For Each lotes In reader.Tables(0).Rows
                    If lotes!paro_wip = "S" Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas a SET oee_parte_actual = 0, oee_lote_actual = 0 WHERE id = " & lotes!id & ";UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET produccion = 0, calidad = 0, produccion_tc = 0, calidad_tc = 0, buffer = 0, parosmostrar = 0, rate_mal_desde = NULL, rate_tendencia_baja = NULL, rate_tendencia_alta = NULL, ftq_tendencia_baja = NULL, oee_tendencia_baja = NULL, dis_tendencia_baja = NULL, efi_tendencia_baja = NULL, parada_desde = NULL, paro_actual = 0, ultima_reparacion = NULL, ultima_produccion = NOW(), alarmado_manual = 'N', ultima_buffer = NULL, iniciar = 'N', iniciar_1 = 'N', iniciar_2 = 'N', iniciar_3 = 'N', iniciar_4 = 'N', iniciar_5 = 'N', iniciar_6 = 'N', iniciar_7 = 'N', iniciar_8 = 'N', fecha_desde = NOW(), planeado = 'N', estado_desde = NOW(), desde_rate = NULL WHERE equipo = " & lotes!id)

                        If lotes!paro_actual > 0 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".detalleparos SET estado = 'F', finaliza = NOW(), hasta = NOW(), tiempo = TIME_TO_SEC(TIMEDIFF(NOW(), inicia)), finalizo = 1, finalizo_accion = 'O' WHERE id = " & lotes!paro_actual)
                        End If
                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".detalleparos (parte, lote, turno, paro, clase, maquina, fecha, tiempo, inicio, inicia, desde, Estado, estatus, notas) VALUES (0, 0, " & miTurno & ", '" & traduccion(102) & "', 1, " & lotes!id & ", NOW(), 0, 1, NOW(), NOW(), 'C', 'A', '" & traduccion(103) & "');UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET estatus = 'D', parada_desde  = NOW(), estado_desde = NOW() WHERE equipo = " & lotes!id)
                    End If
                Next
            End If
        End If
    End Sub



    Sub asignarCarga()
        If modulos(4) = 0 Or Not estadoPrograma Then Exit Sub


        Dim cadSQL As String = "Select dias_programacion, holgura_reprogramar FROM " & rutaBD & ".configuracion"
        Dim readerDS As DataSet = consultaSEL(cadSQL)
        Dim dias_programacion = 30
        Dim regsAfectados = 0
        Dim holgura_reprogramar = 0
        If readerDS.Tables(0).Rows.Count > 0 Then
            Dim configuracion As DataRow = readerDS.Tables(0).Rows(0)
            dias_programacion = ValNull(configuracion!dias_programacion, "N")
            holgura_reprogramar = ValNull(configuracion!holgura_reprogramar, "N")
        Else
            Exit Sub
        End If
        Dim pases = 0
        Dim desasignar = 0
        regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET carga = 0 WHERE estado <> 99 and estatus = 'A'")
        cadSQL = "SELECT a.id, a.equipo, b.parte, b.cantidad, c.proceso FROM " & rutaBD & ".cargas a INNER JOIN " & rutaBD & ".programacion b ON a.id = b.carga AND b.estatus = 'A' LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' WHERE a.estatus = 'A' AND completada <> 'Y' ORDER BY a.fecha ASC"
        Dim reader As DataSet = consultaSEL(cadSQL)
        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + Microsoft.VisualBasic.Strings.Left(errorBD, 250), 9, 0)
        Else

            If reader.Tables(0).Rows.Count > 0 Then
                For Each cargas In reader.Tables(0).Rows
                    cadSQL = "SELECT a.id, IFNULL((SELECT MIN(orden) FROM " & rutaBD & ".prioridades WHERE parte = a.parte AND fecha >= NOW() AND estatus = 'A'), 100) AS prioridad FROM " & rutaBD & ".lotes a WHERE a.estatus = 'A' AND a.carga = 0 AND a.estado <= 50 AND a.proceso = " & cargas!proceso & " AND a.parte = " & cargas!parte & " ORDER BY inspecciones, prioridad, fecha  ASC"
                    Dim reader2 As DataSet = consultaSEL(cadSQL)
                    If reader2.Tables(0).Rows.Count > 0 Then
                        Dim faltan = ValNull(cargas!cantidad, "N")
                        For Each lotes In reader2.Tables(0).Rows
                            If faltan > 0 Then
                                faltan = faltan - 1
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET carga = " & cargas!id & " WHERE id = " & lotes!id)
                                pases = pases + 1
                            Else
                                Exit For
                            End If
                        Next
                    End If
                Next
            End If
            cadSQL = "SELECT a.id, a.equipo, a.estatus, c.proceso, a.fecha, a.completada, IFNULL(z.piezas_2, 0) AS piezas, IFNULL(Y.avance_2, 0) AS avance, IFNULL(X.enequipo_2, 0) AS enequipo FROM " & rutaBD & ".cargas a LEFT JOIN " & rutaBD & ".cat_maquinas c ON a.equipo = c.id AND c.estatus = 'A' LEFT JOIN (SELECT carga, equipo, COUNT(*) AS enequipo_2 FROM " & rutaBD & ".lotes WHERE estado = 50  AND estatus = 'A' GROUP BY carga, equipo) AS X ON X.carga = a.id AND X.equipo = a.equipo LEFT JOIN (SELECT carga, COUNT(*) avance_2 FROM " & rutaBD & ".lotes WHERE estado <= 50 GROUP BY carga) AS Y ON Y.carga = a.id LEFT JOIN (SELECT carga, SUM(cantidad) AS piezas_2 FROM " & rutaBD & ".programacion WHERE estatus = 'A' GROUP BY carga) AS z ON z.carga = a.id WHERE a.estatus = 'A' AND a.completada <> 'Y' ORDER BY a.fecha ASC"
            reader = consultaSEL(cadSQL)
            If reader.Tables(0).Rows.Count > 0 Then
                For Each cargas In reader.Tables(0).Rows
                    If cargas!piezas <= cargas!avance Then
                        If cargas!completada = "N" Then
                            Dim diferencia = DateAndTime.DateDiff(DateInterval.Second, cargas!fecha, DateAndTime.Now)
                            If diferencia < 0 Or diferencia > holgura_reprogramar Then
                                'Se reprograma las programaciones subsiguientes
                                cadSQL = "SELECT a.id, a.fecha, b.proceso FROM " & rutaBD & ".cargas a LEFT JOIN " & rutaBD & ".cat_maquinas b ON a.equipo = b.id WHERE a.permitir_reprogramacion = 'S' AND a.id <> " & cargas!id & " AND a.estatus = 'A' AND a.fecha >= '" & Format(cargas!fecha, "yyyy/MM/dd HH:mm:ss") & "' AND a.equipo = " & cargas!equipo
                                Dim cargasAdic As DataSet = consultaSEL(cadSQL)
                                If cargasAdic.Tables(0).Rows.Count > 0 Then
                                    For Each reprograma In cargasAdic.Tables(0).Rows
                                        Dim FHasta = calcularFechaEstimada(reprograma!fecha, diferencia, reprograma!proceso)
                                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET fecha_anterior = fecha, fecha = '" & Format(FHasta, "yyyy/MM/dd HH:mm:ss") & "', reprogramaciones = reprogramaciones + 1 WHERE id = " & reprograma!id)
                                    Next
                                End If
                                regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET completada = 'S' WHERE id = " & cargas!id)
                            End If
                        ElseIf cargas!enequipo >= cargas!piezas Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET completada = 'Y' WHERE id = " & cargas!id)
                        ElseIf cargas!enequipo > 0 Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET completada = 'U' WHERE id = " & cargas!id)
                        ElseIf cargas!completada <> "S" Then
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET completada = 'S' WHERE id = " & cargas!id)
                        End If
                    ElseIf cargas!completada <> "N" Then
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cargas SET completada = 'N' WHERE id = " & cargas!id)
                    End If
                Next
            End If
            If pases > 0 Then
                agregarLOG(traduccion(104).Replace("campo_0", pases), 1, 1)
            End If
        End If



    End Sub

    Sub calcularEstimado()
        If modulos(4) = 0 Or Not estadoPrograma Then Exit Sub

        Dim aCortar = False
        Dim TTotal As Long = 0

        Dim cadSQL = "SELECT id, ruta, inicia FROM " & rutaBD & ".lotes WHERE ISNULL(estimada)"
        Dim reader As DataSet = consultaSEL(cadSQL)
        Dim regsAfectados = 0
        Dim pases1 = 0
        Dim pases2 = 0

        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + Microsoft.VisualBasic.Strings.Left(errorBD, 250), 9, 0)
        Else
            If reader.Tables(0).Rows.Count > 0 Then
                For Each lotes In reader.Tables(0).Rows
                    cadSQL = "SELECT proceso, tiempo_stock, tiempo_proceso, tiempo_setup FROM " & rutaBD & ".det_rutas WHERE estatus = 'A' AND ruta = " & lotes!ruta & " ORDER BY secuencia"
                    Dim procesos As DataSet = consultaSEL(cadSQL)
                    Dim fechaInicial = lotes!inicia
                    Dim fechaHasta = fechaInicial
                    If procesos.Tables(0).Rows.Count > 0 Then
                        Dim tiempoEstimadoTotal = 0
                        For Each operaciones In procesos.Tables(0).Rows
                            'fechaHasta = calcularFechaEstimada(fechaInicial, operaciones!tiempo_stock, operaciones!proceso)
                            fechaHasta = calcularFechaEstimada(fechaHasta, operaciones!tiempo_proceso + operaciones!tiempo_setup, operaciones!proceso)
                            regsAfectados = consultaACT("UPDATE " & rutaBD & ".ruta_congelada SET inicia = " & DateAndTime.DateDiff(DateInterval.Second, lotes!inicia, fechaHasta) & ", estimada = '" & Format(fechaInicial, "yyyy/MM/dd HH:mm:ss") & "', calcular_hasta = 'N' WHERE id = " & lotes!id)
                            fechaInicial = fechaHasta
                        Next
                        pases1 = pases1 + 1
                        regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes SET tiempo_estimado = " & DateAndTime.DateDiff(DateInterval.Second, lotes!inicia, fechaHasta) & ", estimada = '" & Format(fechaInicial, "yyyy/MM/dd HH:mm:ss") & "', calcular_hasta = 'N' WHERE id = " & lotes!id)
                    End If
                Next
            End If
        End If

        cadSQL = "SELECT a.id, a.proceso, b.tiempo_stock, b.tiempo_proceso, b.tiempo_setup, a.fecha_entrada FROM " & rutaBD & ".lotes_historia a LEFT JOIN " & rutaBD & ".det_rutas b ON a.ruta_detalle = b.id WHERE ISNULL(a.fecha_estimada)"
        reader = consultaSEL(cadSQL)
        regsAfectados = 0
        If errorBD.Length > 0 Then
            agregarLOG(traduccion(21) + Microsoft.VisualBasic.Strings.Left(errorBD, 250), 9, 0)
        Else
            Dim pases = 0
            If reader.Tables(0).Rows.Count > 0 Then
                For Each lotes In reader.Tables(0).Rows
                    Dim fechaHasta = calcularFechaEstimada(lotes!fecha_entrada, lotes!tiempo_stock + lotes!tiempo_proceso + lotes!tiempo_setup, lotes!proceso)
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".lotes_historia SET tiempo_estimado = " & lotes!tiempo_stock + lotes!tiempo_proceso + lotes!tiempo_setup & ",fecha_estimada = '" & Format(fechaHasta, "yyyy/MM/dd HH:mm:ss") & "' WHERE id = " & lotes!id)
                    pases2 = pases2 + 1
                Next
            End If
        End If
        If pases1 > 0 Then agregarSolo(traduccion(105) & pases1 & traduccion(99))
        If pases2 > 0 Then agregarSolo(traduccion(105) & pases2 & traduccion(106))
    End Sub


    Private Sub cincoBotones_Tick(sender As Object, e As EventArgs)
        If cincoBotProcesando Or Not estadoPrograma Then Exit Sub
        cincoBotProcesando = True

        Dim cadSQL = "SELECT id, start_time, requester, requester_key FROM " & rutaMMCALL & ".records WHERE ISNULL(end_time)"
        Dim cadAgregar = ""
        Dim registros As DataSet = consultaSEL(cadSQL)
        Dim pendiente As DataSet
        Dim general As DataSet
        Dim idMaquina As Long = 0
        Dim idLinea As Long = 0
        Dim idArea As Long = 0
        Dim idUSuario As Long = 0
        Dim regsAfectados = 0

        Dim ultimoHora = Format(Now, "HH:mm:ss")
        Dim restar = 0
        Dim reportesCreados = 0



        If registros.Tables(0).Rows.Count > 0 Then
            For Each registro In registros.Tables(0).Rows

                cadSQL = "SELECT record, estatus FROM " & rutaBD & ".temporal_mmcall WHERE record = " & registro!id
                pendiente = consultaSEL(cadSQL)
                If pendiente.Tables(0).Rows.Count = 0 Then

                    Dim fechaReporte = registro!start_time
                    cadSQL = "SELECT * FROM " & rutaBD & ".cat_turnos WHERE id = " & be_turno_actual
                    Dim horarios As DataSet = consultaSEL(cadSQL)
                    If horarios.Tables(0).Rows.Count > 0 Then
                        If ValNull(horarios.Tables(0).Rows(0)!cambiodia, "A") = "S" Then

                            If ValNull(horarios.Tables(0).Rows(0)!mover, "N") = 1 And Format(fechaReporte, "HH") >= "00" And Format(fechaReporte, "HH:mm:ss") < horarios.Tables(0).Rows(0)!termina.ToString Then
                                fechaReporte = DateAdd(DateInterval.Day, -1, fechaReporte)
                            ElseIf ValNull(horarios.Tables(0).Rows(0)!mover, "N") = 2 And Format(fechaReporte, "HH:mm:ss") >= horarios.Tables(0).Rows(0)!inicia.ToString And Format(fechaReporte, "HHmmss") <= "235959" Then
                                fechaReporte = DateAdd(DateInterval.Day, 1, fechaReporte)
                            End If
                        End If
                        idUSuario = ValNull(horarios.Tables(0).Rows(0)!usuario, "N")
                    End If
                    idMaquina = 0
                    idLinea = 0
                    idArea = 0
                    Dim tipo_andon = 0
                    'Se busca máquina/línea
                    cadSQL = "SELECT id, nombre, id_mmcall, linea, usuario, tipo_andon FROM " & rutaBD & ".cat_maquinas WHERE estatus = 'A' AND NOT ISNULL(id_mmcall) ORDER BY id_mmcall DESC, id"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        Dim mHallada = False
                        For Each maquinas In general.Tables(0).Rows
                            If ValNull(maquinas!id_mmcall, "A") <> "" Then
                                Dim arreCanales = maquinas!id_mmcall.Split(New Char() {";"c})
                                For i = LBound(arreCanales) To UBound(arreCanales)
                                    Dim reqMaquina = Strings.Trim(arreCanales(i)).ToUpper
                                    Dim reqMMCall = Strings.Trim(registro!requester).ToUpper

                                    reqMaquina = Strings.Replace(reqMaquina, vbCrLf, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbCr, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbLf, "")
                                    If reqMaquina = reqMMCall Then

                                        idMaquina = maquinas!id
                                        idLinea = maquinas!linea
                                        If ValNull(maquinas!usuario, "N") > 0 Then
                                            idUSuario = ValNull(maquinas!usuario, "N")
                                        End If
                                        tipo_andon = ValNull(maquinas!tipo_andon, "N")
                                        mHallada = True
                                        Exit For
                                    End If
                                Next
                                If mHallada Then Exit For
                            End If
                        Next

                    End If

                    'Se busca área
                    cadSQL = "SELECT id, id_mmcall, botoneras FROM " & rutaBD & ".cat_areas WHERE estatus = 'A' AND NOT ISNULL(id_mmcall) ORDER BY botoneras DESC, id_mmcall DESC, id"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        Dim aHallada = False
                        Dim bHallada = False
                        For Each maquinas In general.Tables(0).Rows
                            If ValNull(maquinas!botoneras, "A") <> "" Then
                                Dim arreCanales = maquinas!botoneras.Split(New Char() {";"c})
                                For i = LBound(arreCanales) To UBound(arreCanales)
                                    Dim reqMaquina = Strings.Trim(arreCanales(i)).ToUpper
                                    Dim reqMMCall = Strings.Trim(registro!requester).ToUpper

                                    reqMaquina = Strings.Replace(reqMaquina, vbCrLf, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbCr, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbLf, "")
                                    If reqMaquina = reqMMCall Then
                                        bHallada = True
                                        Exit For
                                    End If
                                Next
                            Else
                                bHallada = True
                            End If
                            If ValNull(maquinas!id_mmcall, "A") <> "" And bHallada Then
                                Dim arreCanales = maquinas!id_mmcall.Split(New Char() {";"c})

                                For i = LBound(arreCanales) To UBound(arreCanales)
                                    Dim reqBotonera = Strings.Trim(arreCanales(i)).ToUpper
                                    Dim reqMaquina = Strings.Trim(arreCanales(i)).ToUpper
                                    Dim reqMMCall = Strings.Trim(registro!requester_key).ToUpper
                                    Dim reqMMCall2 = Strings.Trim(registro!requester).ToUpper

                                    reqMaquina = Strings.Replace(reqMaquina, vbCrLf, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbCr, "")
                                    reqMaquina = Strings.Replace(reqMaquina, vbLf, "")
                                    If tipo_andon = 2 Then '5 botones
                                        If Strings.InStr(reqMaquina, "(") > 0 Then
                                            'Extraer botonera y botón
                                            reqBotonera = Strings.Left(reqMaquina, Strings.InStr(reqMaquina, "(") - 1)
                                            reqMaquina = Strings.Mid(reqMaquina, Strings.InStr(reqMaquina, "(") + 1, 1)
                                            If Val(reqMaquina) = Val(reqMMCall) And reqBotonera = reqMMCall2 Then
                                                idArea = maquinas!id
                                                aHallada = True
                                                Exit For
                                            End If
                                        ElseIf reqMaquina = reqMMCall Then
                                            idArea = maquinas!id
                                            aHallada = True
                                            Exit For
                                        End If
                                    ElseIf tipo_andon = 1 Then
                                        If reqMaquina = reqMMCall2 Then
                                            aHallada = True
                                            idArea = maquinas!id
                                            Exit For
                                        End If
                                    End If
                                Next
                                If aHallada Then Exit For

                            End If
                        Next
                    End If
                    'Se agrega el registro en la tabla de reportes
                    reportesCreados = reportesCreados + 1
                    cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".reportes (linea, maquina, area, falla, falla_ajustada, solicitante, turno, fecha_reporte, fecha, mmcall, origen) VALUES(" & idLinea & ", " & idMaquina & ", " & idArea & ", 0, 0, " & IIf(idUSuario = 0, 1, idUSuario) & ", " & be_turno_actual & ", '" & Format(fechaReporte, "yyyy/MM/dd") & "', '" & Format(fechaReporte, "yyyy/MM/dd HH:mm:ss") & "', NOW(), " & registro!id & ")"
                    cadAgregar = cadAgregar & ";INSERT INTO " & rutaBD & ".temporal_mmcall (record, boton1) VALUES(" & registro!id & ", '" & Format(registro!start_time, "yyyy/MM/dd HH:mm:ss") & "');"
                End If
            Next
            If cadAgregar.Length > 0 Then
                cadAgregar = Microsoft.VisualBasic.Strings.Mid(cadAgregar, 1, Len(cadAgregar) - 1)
                regsAfectados = consultaACT(cadAgregar)
                If reportesCreados > 0 Then
                    agregarLOG(traduccion(172).Replace("campo_0", reportesCreados), 0, 0)
                End If
            End If
        End If



        'Se revisa de reversa (cierre de tickets)
        cadSQL = "SELECT * FROM " & rutaBD & ".temporal_mmcall WHERE estatus = 0"
        registros = consultaSEL(cadSQL)
        If registros.Tables(0).Rows.Count > 0 Then
            reportesCreados = 0
            cadAgregar = ""
            For Each registro In registros.Tables(0).Rows
                cadSQL = "SELECT end_time FROM " & rutaMMCALL & ".records WHERE id = " & registro!record & " AND NOT ISNULL(end_time)"
                pendiente = consultaSEL(cadSQL)
                cadSQL = "SELECT a.area, b.tecnico, b.falla, b.cerrar_boton FROM " & rutaBD & ".reportes a INNER JOIN " & rutaBD & ".cat_areas b ON a.area = b.id WHERE a.origen = " & registro!record
                general = consultaSEL(cadSQL)
                Dim accion = "N"
                Dim tecnico = 0
                Dim falla = 0
                Dim fANDON = Format(DateAndTime.Now, "yyyy/MM/dd HH:mm:ss")
                If general.Tables(0).Rows.Count > 0 Then
                    accion = ValNull(general.Tables(0).Rows(0)!cerrar_boton, "A")
                    tecnico = ValNull(general.Tables(0).Rows(0)!tecnico, "N")
                    falla = ValNull(general.Tables(0).Rows(0)!falla, "N")
                End If
                If pendiente.Tables(0).Rows.Count > 0 Then
                    Dim fMMcall = Format(pendiente.Tables(0).Rows(0)!end_time, "yyyy/MM/dd HH:mm:ss")
                    reportesCreados = reportesCreados + 1
                    'Se valida si se cierra el reporte o no
                    If accion = "N" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 10, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)) WHERE origen = " & registro!record
                    ElseIf accion = "R" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 100, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), cierre_atencion = '" & fMMcall & "', tiemporeparacion = 0, inicio_reporte = NOW(), detalle = '" & traduccion(107) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), falla = " & falla & ", falla_ajustada = " & falla & ", tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE origen = " & registro!record
                    ElseIf accion = "C" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 1000, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), cierre_atencion = '" & fMMcall & "', tiemporeparacion = 0, tiemporeporte = 0, inicio_reporte = '" & fMMcall & "', cierre_reporte = '" & fMMcall & "', detalle = '" & traduccion(108) & "', comentarios = '" & traduccion(108) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), falla = " & falla & ", falla_ajustada = " & falla & ", tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE origen = " & registro!record
                    End If
                    cadAgregar = cadAgregar & ";UPDATE " & rutaBD & ".temporal_mmcall SET estatus = 1 WHERE record = " & registro!record & ";"
                Else
                    cadSQL = "SELECT end_time FROM " & rutaMMCALL & ".records WHERE id = " & registro!record
                    pendiente = consultaSEL(cadSQL)
                    If pendiente.Tables(0).Rows.Count = 0 Then
                        'Se eliminó el registro
                        reportesCreados = reportesCreados + 1
                        If accion = "N" Then
                            cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 10, inicio_atencion = '" & fANDON & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fANDON & "', fecha)) WHERE origen = " & registro!record
                        ElseIf accion = "R" Then
                            cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 100, inicio_atencion = '" & fANDON & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fANDON & "', fecha)), cierre_atencion = '" & fANDON & "', tiemporeparacion = 0, inicio_reporte = NOW(), detalle = '" & traduccion(107) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fANDON & "', fecha)), falla = " & falla & ", falla_ajustada = " & falla & ", tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE origen = " & registro!record
                        ElseIf accion = "C" Then
                            cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 1000, inicio_atencion = '" & fANDON & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fANDON & "', fecha)), cierre_atencion = '" & fANDON & "', tiemporeparacion = 0, tiemporeporte = 0, inicio_reporte = '" & fANDON & "', cierre_reporte = '" & fANDON & "', detalle = '" & traduccion(108) & "', comentarios = '" & traduccion(108) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fANDON & "', fecha)), falla = " & falla & ", falla_ajustada = " & falla & ", tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE origen = " & registro!record
                        End If
                        'Se agrega el registro en la tabla de reportes
                        cadAgregar = cadAgregar & ";UPDATE " & rutaBD & ".temporal_mmcall SET estatus = 1 WHERE record = " & registro!record & ";"
                    End If
                End If
            Next
            If cadAgregar.Length > 0 Then
                cadAgregar = Microsoft.VisualBasic.Strings.Mid(cadAgregar, 1, Len(cadAgregar) - 1)
                regsAfectados = consultaACT(cadAgregar)
                If reportesCreados > 0 Then
                    agregarLOG(traduccion(173).Replace("campo_0", reportesCreados), 0, 0)
                End If
            End If
        End If

        cincoBotProcesando = False
    End Sub

    Private Sub tmpPrueba_Tick(sender As Object, e As EventArgs)

        If modulos(5) = 0 Or Not estadoPrograma Then Exit Sub

        Dim regsAfectados = consultaACT("UPDATE " & rutaBD & ".lecturas SET estatus = 9 WHERE estatus = 0 AND (LENGTH(sensor) = 10 AND LEFT(sensor, 2) = 99)")

        Dim cadSQL = "SELECT sensor, valor FROM " & rutaBD & ".lecturas WHERE estatus = 9"

        Dim capturas = consultaSEL(cadSQL)
        If capturas.Tables(0).Rows.Count > 0 Then
            For Each captura In capturas.Tables(0).Rows
                'Se pregunta por el diálogo
                Dim equipo = Val(Strings.Right(captura!sensor, 8))
                Dim dialogo As String = ValNull(captura!valor, "A")
                If dialogo.Length > 0 Then
                    cadSQL = ""
                    Dim dialogos = captura!valor.Split(New Char() {"~"c})
                    If dialogos.Length = 3 Then
                        If dialogos(0) = "99" Then
                            If dialogos(1) = "000010" Then
                                cadSQL = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET velocidad = '" & dialogos(2) & "' WHERE equipo = " & equipo & ";"
                            ElseIf dialogos(1) = "000020" Then
                                If dialogos(2) = "0" Then
                                    cadSQL = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET detener = 1, detener_estimado = 1800, detener_tipo = 1, detener_notas = '" & traduccion(109) & "', detener_area = 1, detener_piezas = 'S', detener_paro = '" & traduccion(109) & "' WHERE equipo = " & equipo & ";"
                                Else
                                    cadSQL = "UPDATE " & rutaBD & ".relacion_maquinas_lecturas SET reanudar = 1 WHERE equipo = " & equipo & ";"
                                End If
                            ElseIf dialogos(1) = "000030" Then
                                If dialogos(2) = "0" Then
                                    cadSQL = "INSERT INTO " & rutaMMCALL & ".tasks (location_id, task, message, recipients, status, created) VALUES (1, 'page', 'Axis 21: Emergencia, extintor extraido', 105, 0, NOW());"
                                End If

                            End If
                            regsAfectados = consultaACT(cadSQL & "UPDATE " & rutaBD & ".lecturas Set estatus = 2 WHERE estatus = 9")
                        End If
                    End If

                End If

            Next
        End If
    End Sub

    Private Sub checklist_Tick(sender As Object, e As EventArgs)
        If botCheckList Or Not estadoPrograma Then Exit Sub
        botCheckList = True


        Dim eMensaje = ""
        'Escalada 4
        Dim regsAfectados = 0

        Dim maximo_largo_mmcall As Integer = 40

        Dim cadSQL As String = "SELECT * FROM " & rutaBD & ".plan_checklists WHERE anticipacion = 'S' AND estatus = 'A' AND (ISNULL(anticipado) OR anticipado <> CURDATE()) AND DATE_ADD(hora, INTERVAL tiempo * -1 SECOND) <= CURTIME()"
        Dim planes = consultaSEL(cadSQL)

        Dim idAlerta = 0
        Dim regsAfectadosT As Boolean = False

        'If planes.Tables(0).Rows.Count > 0 Then
        '    For Each plan In planes.Tables(0).Rows
        '        Dim enviarDia As Boolean = False
        '        Dim diaSemana = DateAndTime.Weekday(Now)
        '        If plan!frecuencia = "T" Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "LV" And diaSemana >= 2 And diaSemana <= 6 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "L" And diaSemana = 2 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "M" And diaSemana = 3 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "MI" And diaSemana = 4 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "J" And diaSemana = 5 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "V" And diaSemana = 6 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "S" And diaSemana = 7 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "D" And diaSemana = 1 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "1M" And Val(Today.Day) = 1 Then
        '            enviarDia = True
        '        ElseIf plan!frecuencia = "UM" And Val(Today.Day) = Date.DaysInMonth(Today.Year, Today.Month) Then
        '            enviarDia = True
        '        End If

        '        If enviarDia Then
        '            Dim filtroAdic As String = ""

        '            If plan!tipo <> 0 Then
        '                filtroAdic = " AND a.tipo = " & plan!tipo
        '            End If
        '            If plan!checklists <> "S" Then
        '                cadSQL = "SELECT a.*, b.recipiente FROM " & rutaBD & ".det_plan_checklists a INNER JOIN " & rutaBD & ".cat_checklists b ON a.checklist = b.id AND b.estatus = 'A'" & filtroAdic & " WHERE a.plan = " & plan!id
        '            Else
        '                cadSQL = "SELECT a.id AS checklist, a.recipiente FROM " & rutaBD & ".cat_checklists a WHERE a.estatus = 'A' " & filtroAdic
        '            End If
        '            Dim checklists = consultaSEL(cadSQL)


        '            If checklists.Tables(0).Rows.Count > 0 Then
        '                Dim cadAgregar As String = ""
        '                For Each cl In checklists.Tables(0).Rows
        '                    If ValNull(plan!llamada, "A") = "S" Then
        '                        cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 0, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 0);"
        '                    End If
        '                    If ValNull(plan!sms, "A") = "S" Then
        '                        cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 1, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 0);"
        '                    End If
        '                    If ValNull(plan!correo, "A") = "S" Then
        '                        cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 2, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 0);"
        '                    End If
        '                    If ValNull(plan!mmcall, "A") = "S" Then
        '                        cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 3, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 0);"
        '                    End If
        '                    If plan!asignacion > 1 Then
        '                        If plan!asignadores = "N" And plan!asignacion = 4 Then
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.id = b.registro AND b.tipo = 4 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
        '                        ElseIf plan!asignadores = "N" And plan!asignacion = 3 Then
        '                            'Puestos de trabajo
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.cargo = b.registro AND b.tipo = 3 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
        '                        ElseIf plan!asignadores = "N" And plan!asignacion = 2 Then
        '                            'Departamentos
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.departamento = b.registro AND b.tipo = 2 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
        '                        ElseIf plan!asignadores = "S" And plan!asignacion = 4 Then
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a WHERE a.estatus = 'A'"
        '                        ElseIf plan!asignadores = "S" And plan!asignacion = 3 Then
        '                            'Puestos de trabajo
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".cat_generales b ON a.cargo = b.id AND b.estatus = 'A' WHERE a.estatus = 'A' "
        '                        ElseIf plan!asignadores = "S" And plan!asignacion = 2 Then
        '                            'Departamentos
        '                            cadSQL = "SELECT a.correo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".cat_generales b ON a.departamento = b.id AND b.estatus = 'A' WHERE a.estatus = 'A'"
        '                        End If
        '                        'Se buscan todos los departamentos
        '                        Dim correos = consultaSEL(cadSQL)
        '                        Dim cadCorreos As String = ""

        '                        If correos.Tables(0).Rows.Count > 0 Then
        '                            For Each correo In correos.Tables(0).Rows
        '                                If ValNull(correo!correo, "A") <> "" Then
        '                                    cadCorreos = cadCorreos & correo!correo & ";"
        '                                End If
        '                            Next
        '                        End If
        '                        If cadCorreos.Length > 0 Then

        '                            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-2000, 2, " & cl!checklist & ", " & plan!id & ", 0, 0);")
        '                            Dim reader As DataSet = consultaSEL("SELECT LAST_INSERT_ID();")
        '                            If reader.Tables(0).Rows.Count > 0 Then
        '                                If ValNull(reader.Tables(0).Rows(0).Item(0), "N") > 0 Then
        '                                    regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".checklist_correos (mensaje, correos) VALUES (" & reader.Tables(0).Rows(0).Item(0) & ", '" & cadCorreos & "')")
        '                                    regsAfectadosT = True
        '                                End If
        '                            End If
        '                        End If
        '                    End If
        '                Next
        '                If cadAgregar.Length > 0 Then
        '                    cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".plan_checklists SET anticipado = NOW() WHERE id = " & plan!id
        '                    regsAfectados = consultaACT(cadAgregar)
        '                    regsAfectadosT = True
        '                End If
        '            End If
        '        End If

        '    Next
        'End If

        'Checklist listo

        cadSQL = "SELECT * FROM " & rutaBD & ".plan_checklists WHERE estatus = 'A' AND (ISNULL(ejecutado) OR ejecutado <> CURDATE())"
        planes = consultaSEL(cadSQL)

        If planes.Tables(0).Rows.Count > 0 Then
            Dim eFecha As String = Format(DateAndTime.Now(), "yyyy/MM/dd HH:mm:ss")
            For Each plan In planes.Tables(0).Rows
                Dim enviarDia As Boolean = False
                Dim diaSemana = DateAndTime.Weekday(Now)
                If plan!frecuencia = "T" Then
                    enviarDia = True
                ElseIf plan!frecuencia = "LV" And diaSemana >= 2 And diaSemana <= 6 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "L" And diaSemana = 2 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "M" And diaSemana = 3 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "MI" And diaSemana = 4 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "J" And diaSemana = 5 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "V" And diaSemana = 6 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "S" And diaSemana = 7 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "D" And diaSemana = 1 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "1M" And Val(Today.Day) = 1 Then
                    enviarDia = True
                ElseIf plan!frecuencia = "UM" And Val(Today.Day) = Date.DaysInMonth(Today.Year, Today.Month) Then
                    enviarDia = True
                End If

                If enviarDia Then
                    enviarDia = Strings.Left(plan!hora.ToString(), 2) = Strings.Left(Format(Now(), "HH"), 2)
                End If

                If enviarDia Then
                    Dim filtroAdic As String = ""

                    If plan!tipo <> 0 Then
                        filtroAdic = " AND a.tipo = " & plan!tipo
                    End If
                    If plan!checklists <> "S" Then
                        cadSQL = "SELECT a.*, b.recipiente, b.tiempo FROM " & rutaBD & ".det_plan_checklists a INNER JOIN " & rutaBD & ".cat_checklists b ON a.checklist = b.id AND b.estatus = 'A'" & filtroAdic & " WHERE a.plan = " & plan!id
                    Else
                        cadSQL = "SELECT a.id AS checklist, a.recipiente, a.tiempo FROM " & rutaBD & ".cat_checklists a WHERE a.estatus = 'A' " & filtroAdic
                    End If
                    Dim checklists = consultaSEL(cadSQL)


                    If checklists.Tables(0).Rows.Count > 0 Then
                        Dim cadAgregar As String = ""
                        For Each cl In checklists.Tables(0).Rows
                            Dim cadCorreos As String = ""

                            'Se llena una sola vez el recorset del checklist
                            cadSQL = "SELECT * FROM " & rutaBD & ".det_checklist WHERE checklist = " & cl!checklist & " ORDER BY orden"
                            Dim variables As DataSet = consultaSEL(cadSQL)
                            Dim cadDetalle As String = ""
                            If variables.Tables(0).Rows.Count > 0 Then
                                cadDetalle = "INSERT INTO " & rutaBD & ".checkeje_det (checklist, variable, orden) VALUES "
                                For Each variable In variables.Tables(0).Rows
                                    cadDetalle = cadDetalle & "(#Checklist#, " & variable!variable & ", " & variable!orden & "),"
                                Next
                                cadDetalle = Strings.Left(cadDetalle, cadDetalle.Length - 1)
                            End If
                            Dim vencimiento = ""

                            If cl!tiempo > 0 Then
                                vencimiento = Format(calcularFechaEstimada(Now, cl!tiempo, -1), "yyyy/MM/dd HH:mm:ss")
                            End If
                            If plan!ejecucion = "S" Then
                                If ValNull(plan!llamada2, "A") = "S" Then
                                    cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 0, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 1)"
                                End If
                                If ValNull(plan!sms2, "A") = "S" Then
                                    cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 1, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 1)"
                                End If
                                If ValNull(plan!correo2, "A") = "S" Then
                                    cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 2, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 1)"
                                End If
                                If ValNull(plan!mmcall2, "A") = "S" Then
                                    cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-1000, 3, " & cl!checklist & ", " & plan!id & ", " & cl!recipiente & ", 1)"
                                End If
                            End If

                            If plan!asignacion > 1 Then
                                If plan!asignadores = "N" And plan!asignacion = 4 Then
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.id = b.registro AND b.tipo = 4 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
                                ElseIf plan!asignadores = "N" And plan!asignacion = 3 Then
                                    'Puestos de trabajo
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.cargo = b.registro AND b.tipo = 3 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
                                ElseIf plan!asignadores = "N" And plan!asignacion = 2 Then
                                    'Departamentos
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".relacion_plan_checklists b ON a.departamento = b.registro AND b.tipo = 2 AND b.plan = " & plan!id & " WHERE a.estatus = 'A'"
                                ElseIf plan!asignadores = "S" And plan!asignacion = 4 Then
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a WHERE a.estatus = 'A'"
                                ElseIf plan!asignadores = "S" And plan!asignacion = 3 Then
                                    'Puestos de trabajo
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".cat_generales b ON a.cargo = b.id AND b.estatus = 'A' WHERE a.estatus = 'A' "
                                ElseIf plan!asignadores = "S" And plan!asignacion = 2 Then
                                    'Departamentos
                                    cadSQL = "SELECT a.correo, a.id, a.departamento, a.cargo FROM " & rutaBD & ".cat_usuarios a INNER JOIN " & rutaBD & ".cat_generales b ON a.departamento = b.id AND b.estatus = 'A' WHERE a.estatus = 'A'"
                                End If
                                'Se buscan todos los departamentos
                                Dim correos = consultaSEL(cadSQL)


                                If correos.Tables(0).Rows.Count > 0 Then
                                    For Each correo In correos.Tables(0).Rows
                                        If ValNull(correo!correo, "A") <> "" Then
                                            cadCorreos = cadCorreos & correo!correo & ";"
                                        End If
                                        'Se crean los checklist por usuario
                                        Dim nuevoID = 1
                                        'Se crean los checklist
                                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".checkeje_cab (checklist, plan, vence, frecuencia, asignado_a, asignado_original, departamento, cargo, tipo_asignacion, creado, modificado, creacion, modificacion) VALUES (" & cl!checklist & ", " & plan!id & ", " & IIf(vencimiento = "", "null", "'" & vencimiento & "'") & ", '" & plan!frecuencia & "', " & correo!id & ", " & correo!id & ", " & correo!departamento & ", " & correo!cargo & ", " & plan!asignacion & ", 1, 1, '" & eFecha & "', '" & eFecha & "');")
                                        Dim bNumero As DataSet = consultaSEL("SELECT MAX(id) FROM " & rutaBD & ".checkeje_cab")
                                        If bNumero.Tables(0).Rows.Count > 0 Then
                                            nuevoID = ValNull(bNumero.Tables(0).Rows(0).Item(0), "N")
                                            regsAfectados = consultaACT(cadDetalle.Replace("#Checklist#", nuevoID))
                                        End If
                                    Next
                                End If
                            Else
                                If plan!asignadores = "N" And plan!asignacion = 0 Then
                                    cadSQL = "SELECT b.id AS departamento, 0 AS cargo FROM " & rutaBD & ".relacion_plan_checklists a INNER JOIN " & rutaBD & ".cat_generales b ON a.registro = b.id AND b.estatus = 'A' WHERE a.tipo = 0 AND a.plan = " & plan!id
                                ElseIf plan!asignadores = "S" And plan!asignacion = 0 Then
                                    'Puestos de trabajo
                                    cadSQL = "SELECT a.id AS departamento, 0 AS cargo FROM " & rutaBD & ".cat_generales WHERE tabla = 70 AND estatus = 'A'"
                                ElseIf plan!asignadores = "N" And plan!asignacion = 1 Then
                                    cadSQL = "SELECT b.id AS cargo, 0 AS departamento FROM " & rutaBD & ".relacion_plan_checklists a INNER JOIN " & rutaBD & ".cat_generales b ON a.registro = b.id AND b.estatus = 'A' WHERE a.tipo = 1 AND a.plan = " & plan!id
                                ElseIf plan!asignadores = "S" And plan!asignacion = 1 Then
                                    'Puestos de trabajo
                                    cadSQL = "SELECT a.id AS cargo, 0 AS departamento FROM " & rutaBD & ".cat_generales WHERE tabla = 130 AND estatus = 'A'"
                                End If

                                Dim correos = consultaSEL(cadSQL)
                                If correos.Tables(0).Rows.Count > 0 Then
                                    For Each correo In correos.Tables(0).Rows
                                        'Se crean los checklist por registro
                                        Dim nuevoID = 1
                                        'Se crean los checklist
                                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".checkeje_cab (checklist, plan, vence, frecuencia, departamento, cargo, tipo_asignacion, creado, modificado, creacion, modificacion) VALUES (" & cl!checklist & ", " & plan!id & ", " & IIf(vencimiento = "", "null", vencimiento) & ", '" & plan!frecuencia & "', " & correo!departamento & ", " & correo!cargo & ", " & plan!asignacion & ", 1, 1, '" & eFecha & "', '" & eFecha & "');")
                                        Dim bNumero As DataSet = consultaSEL("SELECT MAX(id) FROM " & rutaBD & ".checkeje_cab")
                                        If bNumero.Tables(0).Rows.Count > 0 Then
                                            nuevoID = ValNull(bNumero.Tables(0).Rows(0).Item(0), "N")
                                            regsAfectados = consultaACT(cadDetalle.Replace("#Checklist#", nuevoID))
                                        End If
                                    Next
                                End If
                            End If

                            If cadCorreos.Length > 0 And plan!ejecucion = "S" Then
                                regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".mensajes (alerta, canal, proceso, prioridad, lista, alarma) VALUES (-2000, 2, " & cl!checklist & ", " & plan!id & ", 0, 1);")
                                Dim reader As DataSet = consultaSEL("SELECT MAX(id) FROM " & rutaBD & ".mensajes")
                                If reader.Tables(0).Rows.Count > 0 Then
                                    If ValNull(reader.Tables(0).Rows(0).Item(0), "N") > 0 Then
                                        regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".checklist_correos (mensaje, correos) VALUES (" & reader.Tables(0).Rows(0).Item(0) & ", '" & cadCorreos & "')")
                                        regsAfectadosT = True
                                    End If
                                End If
                            End If
                        Next
                        If cadAgregar.Length > 0 Then
                            cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".plan_checklists SET ejecutado = NOW() WHERE id = " & plan!id
                            regsAfectados = consultaACT(cadAgregar)
                            regsAfectadosT = True
                        End If
                    End If
                End If

            Next
        End If
        If regsAfectadosT Then
            crearMensajesDOS()
        End If

    End Sub

    Private Sub TextEdit1_EditValueChanged(sender As Object, e As EventArgs) Handles TextEdit1.EditValueChanged

    End Sub

    Sub actualizarBD()
        'Se valida que exista la BD
        Dim cadSQL = "USE " & rutaBD
        Dim reader As DataSet = consultaSEL(cadSQL)
        Dim regsAfectados = 0
        Dim huboProceso = False
        If reader Is Nothing Then
            cadSQL = "Select * FROM siggma.configuracion" & rutaBD
            reader = consultaSEL(cadSQL)
            If reader Is Nothing Then
                huboProceso = True
                'Se crea la bases de datos
                regsAfectados = consultaACT("CREATE DATABASE /*!32312 If Not EXISTS*/`sigma` /*!40100 Default CHARACTER Set latin1 */;

USE `sigma`;

/*Table structure for table `actualizaciones` */

DROP TABLE IF EXISTS `actualizaciones`;

CREATE TABLE `actualizaciones` (
  `plantas` datetime DEFAULT NULL COMMENT 'Se actualizaron las plantas?',
  `lineas` datetime DEFAULT NULL COMMENT 'Se actualizaron las líneas?',
  `maquinas` datetime DEFAULT NULL COMMENT 'Se actualizaron las máquinas?',
  `procesos` datetime DEFAULT NULL COMMENT 'Se actualizaron las procesos?',
  `rutas` datetime DEFAULT NULL COMMENT 'Se actualizaron las rutas?',
  `det_rutas` datetime DEFAULT NULL COMMENT 'Se actualizaron las rutas detalle?',
  `det_procesos` datetime DEFAULT NULL COMMENT 'Se actualizaron los procesos detalle?',
  `partes` datetime DEFAULT NULL COMMENT 'Se actualizaron las partes?',
  `recipientes` datetime DEFAULT NULL COMMENT 'Se actualizaron las recipientes?',
  `alertas` datetime DEFAULT NULL COMMENT 'Se actualizaron las alertas?',
  `situaciones` datetime DEFAULT NULL COMMENT 'Se actualizaron las situaciones?',
  `horario` datetime DEFAULT NULL COMMENT 'Se actualizaron las horarios?',
  `planes` datetime DEFAULT NULL COMMENT 'Se actualizaron las planes?',
  `prioridades` datetime DEFAULT NULL COMMENT 'Se actualizaron las prioridades?',
  `areas` datetime DEFAULT NULL COMMENT 'Se actualizaron las areas?',
  `fallas` datetime DEFAULT NULL COMMENT 'Se actualizaron las fallas?',
  `generales` datetime DEFAULT NULL COMMENT 'Se actualizaron las tablas generales?',
  `distribucion` datetime DEFAULT NULL COMMENT 'Se actualizaron las distribuciones?',
  `correos` datetime DEFAULT NULL COMMENT 'Se actualizaron los correos?',
  `turnos` datetime DEFAULT NULL COMMENT 'Se actualizaron los turnos?',
  `usuarios` datetime DEFAULT NULL COMMENT 'Se actualizaron los usuarios?',
  `traducciones` datetime DEFAULT NULL COMMENT 'Se actualizaron las traducciones?',
  `politicas` datetime DEFAULT NULL COMMENT 'Se actualizaron las politicas?',
  `rates` datetime DEFAULT NULL COMMENT 'Se actualizaron las rates?',
  `estimados` datetime DEFAULT NULL COMMENT 'Se actualizaron las estimados?',
  `objetivos` datetime DEFAULT NULL COMMENT 'Se actualizaron las objetivos?',
  `sensores` datetime DEFAULT NULL COMMENT 'Se actualizaron las sensores?',
  `paros` datetime DEFAULT NULL COMMENT 'Se actualizaron los paros?',
  `rechazos` datetime DEFAULT NULL,
  `variables` datetime DEFAULT NULL,
  `checklists` datetime DEFAULT NULL,
  `valores` datetime DEFAULT NULL,
  `planes_checklists` datetime DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Control de actualización';

/*Data for the table `actualizaciones` */

insert  into `actualizaciones`(`plantas`,`lineas`,`maquinas`,`procesos`,`rutas`,`det_rutas`,`det_procesos`,`partes`,`recipientes`,`alertas`,`situaciones`,`horario`,`planes`,`prioridades`,`areas`,`fallas`,`generales`,`distribucion`,`correos`,`turnos`,`usuarios`,`traducciones`,`politicas`,`rates`,`estimados`,`objetivos`,`sensores`,`paros`,`rechazos`,`variables`,`checklists`,`valores`,`planes_checklists`) values ('2020-01-01 10:00:00','2020-07-15 09:24:41','2020-08-01 02:05:42','0000-00-00 00:00:00','0000-00-00 00:00:00','0000-00-00 00:00:00',NULL,'2020-07-31 23:46:39',NULL,'2020-07-30 16:23:40',NULL,NULL,'0000-00-00 00:00:00',NULL,'2020-07-14 13:22:15','2020-07-21 09:11:37','2020-07-16 17:59:51','2020-07-14 13:24:12','2020-03-27 10:25:09',NULL,'2020-07-15 09:10:47',NULL,NULL,'2020-07-17 14:36:36','2020-07-17 15:04:00','2020-07-31 14:37:47',NULL,'2020-07-22 10:57:29','2020-07-26 23:33:47',NULL,NULL,NULL,NULL);

/*Table structure for table `alarmas` */

DROP TABLE IF EXISTS `alarmas`;

CREATE TABLE `alarmas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `alerta` bigint(20) DEFAULT '0' COMMENT 'ID de la alerta',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `proceso_detalle` bigint(20) DEFAULT '0' COMMENT 'ID del detalle del proceso',
  `prioridad` varchar(1) DEFAULT NULL COMMENT 'Prioridad de la alerta',
  `inicio` datetime DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de inicio de la alarma',
  `fin` datetime DEFAULT NULL COMMENT 'Fecha de fin de la alarma',
  `tiempo` bigint(8) DEFAULT '0' COMMENT 'Tiempo que duro la alarma activa (en segundos)',
  `activada` datetime DEFAULT NULL COMMENT 'Fecha de activación',
  `repetida` datetime DEFAULT NULL COMMENT 'Fecha de repetición',
  `escalada1` datetime DEFAULT NULL COMMENT 'Fecha de escalamiento (1)',
  `escalada2` datetime DEFAULT NULL COMMENT 'Fecha de escalamiento (2)',
  `escalada3` datetime DEFAULT NULL COMMENT 'Fecha de escalamiento (3)',
  `escalada4` datetime DEFAULT NULL COMMENT 'Fecha de escalamiento (4)',
  `escalada5` datetime DEFAULT NULL COMMENT 'Fecha de escalamiento (5)',
  `fase` int(2) DEFAULT '0' COMMENT 'Fase en que está la alarma',
  `estatus` int(1) DEFAULT '0' COMMENT 'Estatus de la alarma',
  `repeticiones` int(4) DEFAULT '0' COMMENT 'Total repeticiones',
  `escalamientos1` int(4) DEFAULT '0' COMMENT 'Total escalamientos (1)',
  `escalamientos2` int(4) DEFAULT '0' COMMENT 'Total escalamientos (2)',
  `escalamientos3` int(4) DEFAULT '0' COMMENT 'Total escalamientos (3)',
  `escalamientos4` int(4) DEFAULT '0' COMMENT 'Total escalamientos (4)',
  `escalamientos5` int(4) DEFAULT '0' COMMENT 'Total escalamientos (5)',
  `informado` char(1) DEFAULT 'N' COMMENT 'Al terminarse esta informada?',
  `acumulada` char(1) DEFAULT 'N' COMMENT 'Es una alarma acumulada?',
  `termino` bigint(20) DEFAULT '0' COMMENT 'Usuario que terminó la alerta',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`alerta`,`inicio`),
  KEY `NewIndex2` (`inicio`),
  KEY `NewIndex3` (`tiempo`)
) ENGINE=MyISAM AUTO_INCREMENT=332000 DEFAULT CHARSET=latin1 COMMENT='Detalle de alarmas';

/*Data for the table `alarmas` */

/*Table structure for table `calidad_historia` */

DROP TABLE IF EXISTS `calidad_historia`;

CREATE TABLE `calidad_historia` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `lote` bigint(20) DEFAULT NULL COMMENT 'ID del lote',
  `tipo` int(2) DEFAULT NULL COMMENT 'Tipo de movimiento',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del Número de parte',
  `inspeccion_id` bigint(20) DEFAULT '0' COMMENT 'Número de inspección',
  `inspeccionado_por` bigint(20) DEFAULT '0' COMMENT 'Inspeccionado por (usuario)',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia de operación',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `inicia` datetime DEFAULT NULL COMMENT 'Fecha de inicio en el sistema',
  `finaliza` datetime DEFAULT NULL COMMENT 'Fecha de fin en el sistema',
  `tiempo` bigint(12) DEFAULT '0' COMMENT 'Tiempo total',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`lote`),
  KEY `NewIndex2` (`lote`,`parte`,`proceso`),
  KEY `NewIndex3` (`lote`,`parte`,`proceso`,`equipo`,`inicia`),
  KEY `NewIndex4` (`lote`,`finaliza`)
) ENGINE=MyISAM AUTO_INCREMENT=212 DEFAULT CHARSET=latin1 COMMENT='Histórico de inspecciones de calidad';

/*Data for the table `calidad_historia` */

/*Table structure for table `cargas` */

DROP TABLE IF EXISTS `cargas`;

CREATE TABLE `cargas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha y hora de programación',
  `fecha_original` datetime DEFAULT NULL COMMENT 'Fecha original',
  `fecha_anterior` datetime DEFAULT NULL COMMENT 'Fecha anterior',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `alarma` char(1) DEFAULT 'S' COMMENT 'Generar alarma',
  `alarma_rep` char(1) DEFAULT 'N' COMMENT 'Alarma reportada?',
  `permitir_reprogramacion` char(1) DEFAULT 'S' COMMENT 'Permitir reprogramación?',
  `completada` char(1) DEFAULT 'N' COMMENT 'Carga completa',
  `carga` varchar(20) DEFAULT NULL COMMENT 'Número de carga',
  `reprogramaciones` int(4) DEFAULT '0' COMMENT 'Veces que se ha reprogramado',
  `alarma_rep_p` char(1) NOT NULL DEFAULT 'N',
  `alarma_rep_paso` char(1) DEFAULT NULL,
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=66 DEFAULT CHARSET=latin1 COMMENT='Tabla de cargas';

/*Data for the table `cargas` */

/*Table structure for table `cat_alertas` */

DROP TABLE IF EXISTS `cat_alertas`;

CREATE TABLE `cat_alertas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `referencia` varchar(50) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(60) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `solapar` char(1) DEFAULT 'N' COMMENT 'Solapar alertas',
  `evento` bigint(20) DEFAULT '0' COMMENT 'ID del evento a monitorear',
  `tipo` int(3) DEFAULT '0' COMMENT 'Tipo: 0=Normal, 9=Escape',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `proceso` char(1) DEFAULT 'S' COMMENT 'Proceso asociado a la alerta',
  `linea` char(1) DEFAULT 'S' COMMENT 'Línea asignada a la alerta',
  `maquina` char(1) DEFAULT 'S' COMMENT 'Máquina asignada a la alerta',
  `area` char(1) DEFAULT 'S' COMMENT 'Área asignada a la alerta',
  `falla` char(1) DEFAULT 'S' COMMENT 'Falla asignada a la alerta',
  `prioridad` int(1) DEFAULT '0' COMMENT 'Prioridad (0: Normal, 1: Alta)',
  `transcurrido` bigint(12) DEFAULT '0' COMMENT 'Tiempo trnscurrido en segundos para la alarma',
  `acumular` char(1) DEFAULT 'N' COMMENT 'Acumular fallas antes de enviar',
  `acumular_veces` bigint(6) DEFAULT '0' COMMENT 'Número de veces a acumular',
  `acumular_tiempo` bigint(8) DEFAULT '0' COMMENT 'Tiempo de acumulación',
  `acumular_inicializar` char(1) DEFAULT 'N' COMMENT 'Inicializa el contador una vez alcanzada la frecuencia',
  `log` char(1) DEFAULT 'N' COMMENT 'Se generará LOG',
  `sms` char(1) DEFAULT 'N' COMMENT 'Se enviará SMS',
  `correo` char(1) DEFAULT 'N' COMMENT 'Se enviará correo',
  `llamada` char(1) DEFAULT 'N' COMMENT 'Se hará llamada',
  `mmcall` char(1) DEFAULT 'N' COMMENT 'Se enviará llamada a MMCall',
  `lista` bigint(20) DEFAULT '0' COMMENT 'Lista de  distribución',
  `veces` int(3) DEFAULT '0' COMMENT 'Veces a alarmar',
  `escalar1` char(1) DEFAULT 'N' COMMENT 'Escalar 1ro',
  `tiempo1` bigint(8) DEFAULT '0' COMMENT 'Tiempo de escalación (1)',
  `lista1` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución (1)',
  `log1` char(1) DEFAULT 'N' COMMENT 'Generar LOG (1)',
  `sms1` char(1) DEFAULT 'N' COMMENT 'Enviar SMS (1)',
  `correo1` char(1) DEFAULT 'N' COMMENT 'Enviar correo (1)',
  `llamada1` char(1) DEFAULT 'N' COMMENT 'Generar Llamada (1)',
  `mmcall1` char(1) DEFAULT 'N' COMMENT 'Área de MMCall (1)',
  `repetir1` char(1) DEFAULT 'N' COMMENT 'Repetir el escalamiento (1)',
  `veces1` int(3) DEFAULT '0' COMMENT 'Veces a escalar (1)',
  `escalar2` char(1) DEFAULT 'N' COMMENT 'Escalar 2do',
  `tiempo2` bigint(8) DEFAULT '0' COMMENT 'Tiempo de escalación (2)',
  `lista2` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución (2)',
  `log2` char(1) DEFAULT 'N' COMMENT 'Generar LOG (2)',
  `sms2` char(1) DEFAULT 'N' COMMENT 'Enviar SMS (2)',
  `correo2` char(1) DEFAULT 'N' COMMENT 'Enviar correo (2)',
  `llamada2` char(1) DEFAULT 'N' COMMENT 'Generar Llamada (2)',
  `mmcall2` char(1) DEFAULT 'N' COMMENT 'Área de MMCall (2)',
  `repetir2` char(1) DEFAULT 'N' COMMENT 'Repetir el escalamiento (2)',
  `veces2` int(3) DEFAULT '0' COMMENT 'Veces a escalar (2)',
  `escalar3` char(1) DEFAULT 'N' COMMENT 'Escalar 3ro',
  `tiempo3` bigint(8) DEFAULT '0' COMMENT 'Tiempo de escalación (3)',
  `lista3` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución (3)',
  `log3` char(1) DEFAULT 'N' COMMENT 'Generar LOG (3)',
  `sms3` char(1) DEFAULT 'N' COMMENT 'Enviar SMS (3)',
  `correo3` char(1) DEFAULT 'N' COMMENT 'Enviar correo (3)',
  `llamada3` char(1) DEFAULT 'N' COMMENT 'Generar Llamada (3)',
  `mmcall3` char(1) DEFAULT 'N' COMMENT 'Área de MMCall (3)',
  `repetir3` char(1) DEFAULT 'N' COMMENT 'Repetir el escalamiento (3)',
  `veces3` int(3) DEFAULT '0' COMMENT 'Veces a escalar (3)',
  `escalar4` char(1) DEFAULT 'N' COMMENT 'Escalar 4to',
  `tiempo4` bigint(8) DEFAULT '0' COMMENT 'Tiempo de escalación (4)',
  `lista4` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución (4)',
  `log4` char(1) DEFAULT 'N' COMMENT 'Generar LOG (4)',
  `sms4` char(1) DEFAULT 'N' COMMENT 'Enviar SMS (4)',
  `correo4` char(1) DEFAULT 'N' COMMENT 'Enviar correo (4)',
  `llamada4` char(1) DEFAULT 'N' COMMENT 'Generar Llamada (4)',
  `mmcall4` char(1) DEFAULT 'N' COMMENT 'Área de MMCall (4)',
  `repetir4` char(1) DEFAULT 'N' COMMENT 'Repetir el escalamiento (4)',
  `veces4` int(3) DEFAULT '0' COMMENT 'Veces a escalar (4)',
  `escalar5` char(1) DEFAULT 'N' COMMENT 'Escalar 5to',
  `tiempo5` bigint(8) DEFAULT '0' COMMENT 'Tiempo de escalación (5)',
  `lista5` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución (5)',
  `log5` char(1) DEFAULT 'N' COMMENT 'Generar LOG (5)',
  `sms5` char(1) DEFAULT 'N' COMMENT 'Enviar SMS (5)',
  `correo5` char(1) DEFAULT 'N' COMMENT 'Enviar correo (5)',
  `llamada5` char(1) DEFAULT 'N' COMMENT 'Generar Llamada (5)',
  `mmcall5` char(1) DEFAULT 'N' COMMENT 'Área de MMCall (5)',
  `repetir5` char(1) DEFAULT 'N' COMMENT 'Repetir el escalamiento (5)',
  `veces5` int(3) DEFAULT '0' COMMENT 'Veces a escalar (5)',
  `repetir` char(1) DEFAULT 'N' COMMENT 'Repetir llamada',
  `repetir_tiempo` bigint(8) DEFAULT '0' COMMENT 'Repetir llamada (segundos)',
  `repetir_log` char(1) DEFAULT 'N' COMMENT 'Generar log en la repetición',
  `repetir_sms` char(1) DEFAULT 'N' COMMENT 'Enviar SMS en la repetición',
  `repetir_correo` char(1) DEFAULT 'N' COMMENT 'Enviar correo en la repetición',
  `repetir_llamada` char(1) DEFAULT 'N' COMMENT 'Generar llamada en la repetición',
  `repetir_mmcall` char(1) DEFAULT 'N' COMMENT 'Área de MMCall en la repetición',
  `repetir_veces` int(4) DEFAULT '0' COMMENT 'Número de veces a repetir',
  `estadistica` char(1) DEFAULT 'N' COMMENT 'Generar estadística',
  `escape_veces` int(2) DEFAULT '3' COMMENT 'Número de veces que se repetirá una llamada',
  `escape_accion` char(1) DEFAULT 'N' COMMENT 'Acción de Escape',
  `escape_mensaje` varchar(200) DEFAULT NULL COMMENT 'Mensaje a enviar si se agotan las llamadas',
  `escape_lista` bigint(20) DEFAULT '0' COMMENT 'Lista de distribución',
  `informar_resolucion` char(1) DEFAULT 'N' COMMENT 'Informar resolución',
  `cancelacion_mensaje` varchar(200) DEFAULT NULL COMMENT 'Mensaje de cancelación',
  `resolucion_mensaje` varchar(200) DEFAULT NULL COMMENT 'Mensaje de resolución',
  `tiempo0` bigint(8) DEFAULT '0' COMMENT 'Tiempo previo para alertas informativas',
  `mensaje` varchar(300) DEFAULT NULL,
  `titulo` varchar(100) DEFAULT NULL,
  `mensaje_mmcall` char(80) DEFAULT NULL,
  `mensaje1_mantener` char(1) DEFAULT 'N',
  `mensaje1` varchar(300) DEFAULT NULL,
  `titulo1` varchar(100) DEFAULT NULL,
  `mensaje1_mmcall` varchar(80) DEFAULT NULL,
  `mensaje2_mantener` char(1) DEFAULT 'N',
  `mensaje2` varchar(300) DEFAULT NULL,
  `titulo2` varchar(100) DEFAULT NULL,
  `mensaje2_mmcall` varchar(50) DEFAULT NULL,
  `mensaje3_mantener` char(1) DEFAULT 'N',
  `mensaje3` varchar(300) DEFAULT NULL,
  `titulo3` varchar(100) DEFAULT NULL,
  `mensaje3_mmcall` varchar(80) DEFAULT NULL,
  `mensaje4_mantener` char(1) DEFAULT 'N',
  `mensaje4` varchar(200) DEFAULT NULL,
  `titulo4` varchar(100) DEFAULT NULL,
  `mensaje4_mmcall` varchar(80) DEFAULT NULL,
  `mensaje5_mantener` char(1) DEFAULT 'N',
  `mensaje5` varchar(200) DEFAULT NULL,
  `titulo5` varchar(100) DEFAULT NULL,
  `mensaje5_mmcall` varchar(80) DEFAULT NULL,
  `repetir_mantener` char(1) DEFAULT 'N',
  `repetir_mensaje` varchar(200) DEFAULT NULL,
  `repetir_titulo` varchar(100) DEFAULT NULL,
  `repetir_mensaje_mmcall` varchar(80) DEFAULT NULL,
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`evento`,`proceso`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=89 DEFAULT CHARSET=latin1 COMMENT='Catálogo de alertas';

/*Data for the table `cat_alertas` */

/*Table structure for table `cat_areas` */

DROP TABLE IF EXISTS `cat_areas`;

CREATE TABLE `cat_areas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `agrupador_1` bigint(20) DEFAULT '0' COMMENT 'Agrupador (1)',
  `agrupador_2` bigint(20) DEFAULT '0' COMMENT 'Agrupador (2)',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `audios_ruta` varchar(1000) DEFAULT NULL COMMENT 'Carpeta de audios',
  `audios_activar` char(1) DEFAULT 'N',
  `audios_prefijo` varchar(1000) DEFAULT NULL COMMENT 'Audio prefijo',
  `audios_general` char(1) DEFAULT 'S' COMMENT 'Grabar también en carpeta general',
  `id_mmcall` varchar(500) DEFAULT NULL COMMENT 'ID de MMCall',
  `recipiente` bigint(20) DEFAULT '0' COMMENT 'ID del recipiente',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `Index01` (`url_mmcall`)
) ENGINE=MyISAM AUTO_INCREMENT=10000 DEFAULT CHARSET=latin1 COMMENT='Catálogo de áreas';

/*Data for the table `cat_areas` */

/*Table structure for table `cat_checklists` */

DROP TABLE IF EXISTS `cat_checklists`;

CREATE TABLE `cat_checklists` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `referencia` varchar(30) DEFAULT NULL,
  `tipo` bigint(20) DEFAULT '0' COMMENT 'ID de tipo',
  `departamento` bigint(20) DEFAULT '0',
  `equipo` bigint(20) DEFAULT '0',
  `equipo_automatico` char(1) DEFAULT 'N',
  `imagen` varchar(255) DEFAULT NULL,
  `variables` char(1) DEFAULT 'N',
  `notas` varchar(300) DEFAULT NULL,
  `url_mmcall` varchar(1000) DEFAULT NULL,
  `tiempo` bigint(10) DEFAULT '0',
  `recipiente` bigint(20) DEFAULT '0',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=9 DEFAULT CHARSET=latin1;

/*Data for the table `cat_checklists` */

/*Table structure for table `cat_correos` */

DROP TABLE IF EXISTS `cat_correos`;

CREATE TABLE `cat_correos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `nombre` varchar(60) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `para` varchar(2000) DEFAULT NULL COMMENT 'Lista de distribución',
  `copia` varchar(2000) DEFAULT NULL COMMENT 'Lista de distribución (con copia)',
  `oculta` varchar(2000) DEFAULT NULL COMMENT 'Lista de distribución (con copia oculta)',
  `titulo` varchar(200) DEFAULT NULL COMMENT 'Título del correo',
  `cuerpo` varchar(1000) DEFAULT NULL COMMENT 'Cuerpo del correo',
  `extraccion` varchar(40) DEFAULT NULL COMMENT 'Extracción del reporte',
  `ultimo_envio` datetime DEFAULT NULL COMMENT 'Fecha del último envío',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=11 DEFAULT CHARSET=latin1 COMMENT='Catálogo de correos';

/*Data for the table `cat_correos` */

/*Table structure for table `cat_defectos` */

DROP TABLE IF EXISTS `cat_defectos`;

CREATE TABLE `cat_defectos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(20) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del defecto',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo de la descripción',
  `agrupador1` bigint(20) DEFAULT NULL COMMENT 'ID del agrupador (1)',
  `agrupador2` bigint(20) DEFAULT NULL COMMENT 'ID del agrupador (2)',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `minimo` decimal(17,7) DEFAULT NULL COMMENT 'Minima cantidad a reportar',
  `maximo` decimal(17,7) DEFAULT NULL COMMENT 'Maxima cantidad a reportar',
  `defecto` char(1) DEFAULT NULL COMMENT 'Registro por defecto',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Catálogo de defectos';

/*Data for the table `cat_defectos` */

/*Table structure for table `cat_distribucion` */

DROP TABLE IF EXISTS `cat_distribucion`;

CREATE TABLE `cat_distribucion` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(60) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `telefonos` varchar(2000) DEFAULT '' COMMENT 'Número de teléfono',
  `correos` varchar(2000) DEFAULT '' COMMENT 'Correo electrónico',
  `mmcall` varchar(2000) DEFAULT '' COMMENT 'Requesters de MMCall',
  `habilitado` char(1) DEFAULT 'S' COMMENT 'Habilitar como disponible',
  `hora_desde` time DEFAULT NULL COMMENT 'Hora de inicio',
  `hora_hasta` time DEFAULT NULL COMMENT 'Hora de fin',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=30 DEFAULT CHARSET=latin1 COMMENT='Catálogo de listas de distribución';

/*Data for the table `cat_distribucion` */

/*Table structure for table `cat_equipos` */

DROP TABLE IF EXISTS `cat_equipos`;

CREATE TABLE `cat_equipos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la linea',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `agrupador1` bigint(11) DEFAULT NULL COMMENT 'ID del agrupador (1)',
  `agrupador2` bigint(11) DEFAULT NULL COMMENT 'ID del agrupador (2)',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Catálogo de equipos';

/*Data for the table `cat_equipos` */

/*Table structure for table `cat_fallas` */

DROP TABLE IF EXISTS `cat_fallas`;

CREATE TABLE `cat_fallas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(200) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Imagen a mostrar',
  `linea` char(1) CHARACTER SET latin1 DEFAULT 'N' COMMENT 'ID de la línea',
  `maquina` char(1) CHARACTER SET latin1 DEFAULT 'N' COMMENT 'ID de la máquina',
  `area` char(1) CHARACTER SET latin1 DEFAULT 'N' COMMENT 'ID del área',
  `agrupador_1` bigint(20) DEFAULT '0' COMMENT 'Agrupador (1)',
  `afecta_oee` char(1) DEFAULT 'S',
  `agrupador_2` bigint(20) DEFAULT '0' COMMENT 'Agrupador (2)',
  `ultima_incidencia` datetime DEFAULT NULL COMMENT 'Última vez que pasó',
  `notas` varchar(300) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Notas varias',
  `url_mmcall` varchar(1000) CHARACTER SET latin1 DEFAULT NULL COMMENT 'URL de MMCall',
  `estatus` char(1) CHARACTER SET latin1 DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  `codigo` varchar(50) CHARACTER SET latin1 DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=1433 DEFAULT CHARSET=utf16 COMMENT='Catálogo de fallas';

/*Data for the table `cat_fallas` */

/*Table structure for table `cat_frases` */

DROP TABLE IF EXISTS `cat_frases`;

CREATE TABLE `cat_frases` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la línea',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `area` bigint(20) DEFAULT '0' COMMENT 'ID del área',
  `falla` bigint(20) DEFAULT '0' COMMENT 'ID de la falla',
  `mensaje` varchar(200) DEFAULT '0' COMMENT 'Agrupador (2)',
  `largo` int(2) DEFAULT NULL COMMENT 'Largo del mensaje',
  `esperado` bigint(12) DEFAULT NULL COMMENT 'Tiempo esperado para reparación',
  `url_mmcall` varchar(250) DEFAULT NULL COMMENT 'URL de MMCall',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;

/*Data for the table `cat_frases` */

/*Table structure for table `cat_generales` */

DROP TABLE IF EXISTS `cat_generales`;

CREATE TABLE `cat_generales` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `tabla` int(6) DEFAULT NULL COMMENT 'ID de la tabla',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=50 DEFAULT CHARSET=latin1 COMMENT='Tablas generales';

/*Data for the table `cat_generales` */

/*Table structure for table `cat_grupos` */

DROP TABLE IF EXISTS `cat_grupos`;

CREATE TABLE `cat_grupos` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(20) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `agrupador1` bigint(11) DEFAULT NULL COMMENT 'ID del agrupador (1)',
  `agrupador2` bigint(11) DEFAULT NULL COMMENT 'ID del agrupador (2)',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 COMMENT='Catálogo de grupos';

/*Data for the table `cat_grupos` */

/*Table structure for table `cat_idiomas` */

DROP TABLE IF EXISTS `cat_idiomas`;

CREATE TABLE `cat_idiomas` (
  `id` int(3) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `codigo` varchar(5) DEFAULT NULL COMMENT 'Código del idioma',
  `nombre` varchar(50) DEFAULT '0' COMMENT 'Nombre del idioma',
  `icono` varchar(10) DEFAULT NULL COMMENT 'icono',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

/*Data for the table `cat_idiomas` */

/*Table structure for table `cat_lineas` */

DROP TABLE IF EXISTS `cat_lineas`;

CREATE TABLE `cat_lineas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `agrupador_1` bigint(20) DEFAULT '0' COMMENT 'Agrupador (1)',
  `agrupador_2` bigint(20) DEFAULT '0' COMMENT 'Agrupador (2)',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `disponibilidad` int(1) DEFAULT '0',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=39 DEFAULT CHARSET=latin1 COMMENT='Catálogo de líneas';

/*Data for the table `cat_lineas` */

/*Table structure for table `cat_listas` */

DROP TABLE IF EXISTS `cat_listas`;

CREATE TABLE `cat_listas` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(20) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `clase` bigint(11) DEFAULT NULL COMMENT 'ID de la clase',
  `area` bigint(11) DEFAULT NULL COMMENT 'ID del área',
  `equipo` bigint(11) DEFAULT NULL COMMENT 'ID del equipo',
  `tiempo_llenado` bigint(6) DEFAULT NULL COMMENT 'Tiempo límite para el llenado (segundos)',
  `tiempo_alarma` char(1) DEFAULT NULL COMMENT 'Generar alarma al sobrepasar el límite por',
  `prioridad` int(2) DEFAULT NULL COMMENT 'Prioridad del registro',
  `horario` bigint(11) DEFAULT NULL COMMENT 'Calendario',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` timestamp NULL DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 COMMENT='Catálogo de listas de verificación';

/*Data for the table `cat_listas` */

/*Table structure for table `cat_maquinas` */

DROP TABLE IF EXISTS `cat_maquinas`;

CREATE TABLE `cat_maquinas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `disponibilidad` int(2) DEFAULT '0' COMMENT 'Disponibilidad de la empresa',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la línea (0=Suelta)',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso asociado',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `programar` char(1) DEFAULT 'N' COMMENT 'Permite la programación',
  `ultimo_parte` bigint(20) DEFAULT '0' COMMENT 'Última parte',
  `capacidad` bigint(12) DEFAULT '0' COMMENT 'Capacidad en segundos',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'Tipo de máquina',
  `agrupador_1` bigint(20) DEFAULT '0' COMMENT 'Agrupador (1)',
  `agrupador_2` bigint(20) DEFAULT '0' COMMENT 'Agrupador (2)',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `oee` char(1) DEFAULT 'N',
  `oee_turno_actual` bigint(20) DEFAULT '0',
  `oee_tripulacion_actual` bigint(20) DEFAULT '0',
  `oee_lote_actual` bigint(20) DEFAULT '0',
  `oee_parte_actual` bigint(20) DEFAULT '0',
  `oee_tipo_rate` int(2) DEFAULT '0' COMMENT 'OJO/Quitar',
  `oee_permitir_resto` char(1) DEFAULT 'N' COMMENT 'OJO/Quitar',
  `oee_objetivo_por_equipo` char(1) DEFAULT 'N' COMMENT 'OJO/Quitar',
  `oee_objetivo_piezas` bigint(20) DEFAULT '0' COMMENT 'OJO/Quitar',
  `oee_objetivo_piezas_base` bigint(20) DEFAULT '0' COMMENT 'OJO/Quitar',
  `oee_progreso_reiniciar` char(1) DEFAULT 'N' COMMENT 'OJO/Quitar',
  `oee_reiniciar_tipo` int(2) DEFAULT '0' COMMENT 'OJO/Quitar',
  `oee_reiniciar_hora` int(2) DEFAULT '0' COMMENT 'OJO/Quitar',
  `oee_umbral_produccion` bigint(6) DEFAULT '0' COMMENT 'Segundos a esperar para paro',
  `oee_umbral_alerta` bigint(6) DEFAULT '0' COMMENT 'Segundos a esperar para alerta',
  `oee_estado` char(1) DEFAULT 'N' COMMENT 'Estado de la máquina',
  `oee_estado_desde` datetime DEFAULT NULL COMMENT 'Fecha del último estado',
  `oee_estado_cambio` datetime DEFAULT NULL COMMENT 'Fecha de cambio',
  `confirmar_reparacion` char(1) DEFAULT 'N' COMMENT 'Confirmar reparacion',
  `oee_historico_rate` int(1) DEFAULT '0',
  `oee_historico_rate_reiniciar` int(1) DEFAULT '0',
  `boton_1` datetime DEFAULT NULL,
  `boton_2` datetime DEFAULT NULL,
  `id_mmcall` varchar(500) DEFAULT NULL COMMENT 'ID de MMCall',
  `usuario` bigint(20) DEFAULT '0' COMMENT 'ID del usuario',
  `tipo_andon` int(1) DEFAULT '0' COMMENT 'Tipo de botonera',
  `activar_buffer` char(1) DEFAULT 'N' COMMENT 'Permitir la activación de buffer',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `id_mapa` int(4) DEFAULT '0' COMMENT 'ID mapa 1',
  `id_mapa2` int(4) DEFAULT '0' COMMENT 'ID mapa 2',
  `id_mapa3` int(4) DEFAULT '0' COMMENT 'ID mapa 3',
  `paro_wip` char(1) DEFAULT 'N' COMMENT 'Aceptar paro desde WIP',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`linea`),
  KEY `Index01` (`id_mmcall`)
) ENGINE=MyISAM AUTO_INCREMENT=1006 DEFAULT CHARSET=latin1 COMMENT='Catálogo de máquinas';

/*Data for the table `cat_maquinas` */

/*Table structure for table `cat_medios` */

DROP TABLE IF EXISTS `cat_medios`;

CREATE TABLE `cat_medios` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `tipo` int(4) DEFAULT NULL COMMENT 'Tipo de comunicaión',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` timestamp NULL DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=6 DEFAULT CHARSET=latin1 COMMENT='Catálogo de medios de envío';

/*Data for the table `cat_medios` */

/*Table structure for table `cat_paros` */

DROP TABLE IF EXISTS `cat_paros`;

CREATE TABLE `cat_paros` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'Tipo de paro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `agrupador1` bigint(20) DEFAULT '0' COMMENT 'ID del agrupador (1)',
  `agrupador2` bigint(20) DEFAULT '0' COMMENT 'ID del agrupador (2)',
  `adelantar` char(1) DEFAULT 'N' COMMENT 'Se puede adelantar?',
  `cancelar` char(1) DEFAULT 'N' COMMENT 'Se puede cancelar?',
  `con_clave` char(1) DEFAULT 'N' COMMENT 'Se puede cabiar con clave',
  `una_vez` char(1) DEFAULT 'N' COMMENT 'Paro de una vez',
  `periodico` char(1) DEFAULT 'N' COMMENT 'Es un paro periódico',
  `semana` char(7) DEFAULT NULL COMMENT 'Día de semana',
  `habiles` char(1) DEFAULT NULL COMMENT 'Se aplica en día no hábiles',
  `desde` time DEFAULT NULL COMMENT 'Hora de inicio',
  `hasta` time DEFAULT NULL COMMENT 'Hora de fin',
  `inicia` date DEFAULT NULL COMMENT 'Fecha de inicio',
  `finaliza` date DEFAULT NULL COMMENT 'Fecha de finalización',
  `tiempo_seg` bigint(10) DEFAULT NULL COMMENT 'Tiempo del pago en segundos',
  `clendario` bigint(20) DEFAULT '0' COMMENT 'ID del calenadario',
  `carpeta` varchar(255) DEFAULT NULL COMMENT 'Carpeta de imagenes y videos',
  `defecto` char(1) DEFAULT NULL COMMENT 'Registro por defecto',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=latin1 COMMENT='Catálogo de paros';

/*Data for the table `cat_paros` */

/*Table structure for table `cat_partes` */

DROP TABLE IF EXISTS `cat_partes`;

CREATE TABLE `cat_partes` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `ruta` bigint(20) DEFAULT '0' COMMENT 'Ruta de fabricación asociada',
  `tipo` int(2) DEFAULT '0' COMMENT 'Tipo de item (0=Producción, 1=SMED/Herramentales, 2=Modelos)',
  `maquinas` char(1) DEFAULT 'N',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=1134 DEFAULT CHARSET=latin1 COMMENT='Catálogo de numeros de parte';

/*Data for the table `cat_partes` */

/*Table structure for table `cat_partes_smed` */

DROP TABLE IF EXISTS `cat_partes_smed`;

CREATE TABLE `cat_partes_smed` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `ruta` bigint(20) DEFAULT '0' COMMENT 'Ruta de fabricación asociada',
  `tipo` int(2) DEFAULT '0' COMMENT 'Tipo de item (0=Producción, 1=SMED/Herramentales, 2=Modelos)',
  `maquinas` char(1) DEFAULT 'N',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=1126 DEFAULT CHARSET=latin1;

/*Data for the table `cat_partes_smed` */

/*Table structure for table `cat_plantas` */

DROP TABLE IF EXISTS `cat_plantas`;

CREATE TABLE `cat_plantas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `idioma` int(4) DEFAULT '0',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=24 DEFAULT CHARSET=latin1 COMMENT='Catálogo de plantas';

/*Data for the table `cat_plantas` */

/*Table structure for table `cat_procesos` */

DROP TABLE IF EXISTS `cat_procesos`;

CREATE TABLE `cat_procesos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `tipo` char(1) DEFAULT NULL COMMENT 'Tipo de proceso: N = Normal, E = Espera',
  `capacidad_stock` bigint(6) DEFAULT '0' COMMENT 'Capacidad en Stock (lotes)',
  `capacidad_proceso` bigint(6) DEFAULT '0' COMMENT 'Capacidad en proceso (lotes)',
  `reduccion_setup` char(1) NOT NULL DEFAULT 'S',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`referencia`)
) ENGINE=MyISAM AUTO_INCREMENT=34 DEFAULT CHARSET=latin1 COMMENT='Catalogo de procesos';

/*Data for the table `cat_procesos` */

/*Table structure for table `cat_rutas` */

DROP TABLE IF EXISTS `cat_rutas`;

CREATE TABLE `cat_rutas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `alarma` char(1) DEFAULT NULL COMMENT 'Genera alarmas?',
  `inicia` int(6) DEFAULT NULL,
  `finaliza` int(6) DEFAULT NULL,
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=33 DEFAULT CHARSET=latin1 COMMENT='Catálogo de rutas de producción';

/*Data for the table `cat_rutas` */

/*Table structure for table `cat_situaciones` */

DROP TABLE IF EXISTS `cat_situaciones`;

CREATE TABLE `cat_situaciones` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `tipo` int(2) DEFAULT NULL COMMENT 'Tipo de situación (0=Calidad, 50= Scrap)',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=latin1 COMMENT='Catálogo de situaciones';

/*Data for the table `cat_situaciones` */

/*Table structure for table `cat_tripulacion` */

DROP TABLE IF EXISTS `cat_tripulacion`;

CREATE TABLE `cat_tripulacion` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `compania` bigint(20) DEFAULT '0' COMMENT 'Compañía asociada',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=6 DEFAULT CHARSET=latin1 COMMENT='Catálogo de tripulaciones';

/*Data for the table `cat_tripulacion` */

/*Table structure for table `cat_turnos` */

DROP TABLE IF EXISTS `cat_turnos`;

CREATE TABLE `cat_turnos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `nombre` varchar(100) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Nombre del registro',
  `referencia` varchar(30) COLLATE utf8_bin DEFAULT NULL COMMENT 'Referencia',
  `inicia` time DEFAULT NULL COMMENT 'Hora de inicio',
  `termina` time DEFAULT NULL COMMENT 'Hora de Fin',
  `cambiodia` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Pasa de un día a otro (S/N)',
  `especial` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Turno especial (S/N)',
  `tipo` int(1) DEFAULT '0' COMMENT 'Tipo de turno (0=Diurno, 1=Matutino, 2=Nocturno, 3=Mixto)',
  `mover` int(1) DEFAULT '0' COMMENT 'Recorrer fecha de reporte (1=Dia anterior, 2=Dia posterior, 0=No recorrer)',
  `secuencia` int(2) DEFAULT NULL COMMENT 'Secuencia del turno',
  `usuario` bigint(20) DEFAULT '0' COMMENT 'ID del usuario',
  `estatus` char(1) COLLATE utf8_bin DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` timestamp NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Usuario que creó el registro',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Usuario que modificó el registro',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=13 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de turnos';

/*Data for the table `cat_turnos` */

/*Table structure for table `cat_usuarios` */

DROP TABLE IF EXISTS `cat_usuarios`;

CREATE TABLE `cat_usuarios` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(50) DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `clave` varchar(255) DEFAULT NULL COMMENT 'Conraseña',
  `prefijo` varchar(50) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(100) DEFAULT NULL COMMENT 'Notas varias',
  `rol` char(1) DEFAULT NULL COMMENT 'Rol de usuario',
  `politica` int(2) DEFAULT NULL COMMENT 'Política de contraseña',
  `operacion` char(1) DEFAULT 'N' COMMENT 'Ver todas las operaciones (S/N)',
  `linea` char(1) DEFAULT 'N' COMMENT 'Ver todas las líneas (S/N)',
  `maquina` char(1) DEFAULT 'N' COMMENT 'Ver todas las máquinas (S/N)',
  `area` char(1) DEFAULT 'N' COMMENT 'Ver todas las áreas (S/N)',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `admin` char(1) DEFAULT 'N' COMMENT 'Es administrador',
  `calidad` char(1) DEFAULT 'N' COMMENT 'Puede hacer inspecciones de calidad',
  `reversos` char(1) DEFAULT 'N' COMMENT 'Puede hacer reversos',
  `programacion` char(1) DEFAULT 'N' COMMENT 'Ver programación(lectura)',
  `inventario` char(1) DEFAULT 'N' COMMENT 'Ver inventario',
  `tecnico` char(1) DEFAULT 'N' COMMENT 'Rol de técnico',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la plant',
  `apagar_alertas` char(1) DEFAULT 'N' COMMENT 'El usuario puede cancelar alertas',
  `ver_alertas` char(1) DEFAULT 'N' COMMENT 'El usuario puede ver las alertas',
  `cerrar_al_ejecutar` char(1) DEFAULT 'N' COMMENT 'Cerrar menú al ajecutar',
  `vista_resumida_fallas` char(1) DEFAULT NULL COMMENT 'Vista resumida de las fallas',
  `confirmar_reparacion` char(1) DEFAULT 'N' COMMENT 'Puede confirmar una reparación',
  `ultima_pantalla` int(2) DEFAULT '0' COMMENT 'Última pantalla usada',
  `inicializada` char(1) DEFAULT 'S' COMMENT 'Contraseña inicializada',
  `preferencias_andon` varchar(50) DEFAULT '00000000000000000000000000000000000000000000000000' COMMENT 'Preferecia de Andon',
  `departamento` bigint(20) DEFAULT '0' COMMENT 'ID del Departamento',
  `compania` bigint(20) DEFAULT '0' COMMENT 'ID de la compañía',
  `tema` int(3) DEFAULT '0' COMMENT 'ID del tema',
  `turno` bigint(20) DEFAULT '0' COMMENT 'Turno asociado',
  `claveant1` varchar(255) DEFAULT NULL COMMENT 'Últimas claves usadas',
  `claveant2` varchar(255) DEFAULT NULL COMMENT 'Últimas claves usadas',
  `claveant3` varchar(255) DEFAULT NULL COMMENT 'Últimas claves usadas',
  `claveant4` varchar(255) DEFAULT NULL COMMENT 'Últimas claves usadas',
  `claveant5` varchar(255) DEFAULT NULL COMMENT 'Últimas claves usadas',
  `ucambio` date DEFAULT NULL COMMENT 'Fecha del último cambio',
  `entrada` datetime DEFAULT NULL COMMENT 'Fecha de la última entrada',
  `salida` datetime DEFAULT NULL COMMENT 'Fecha de la última salida',
  `ultimo_reporte` int(1) DEFAULT '0' COMMENT 'Último reporte consultado',
  `idioma` int(4) DEFAULT '0',
  `planta_defecto` bigint(20) DEFAULT '0',
  `plantas` char(1) DEFAULT 'S',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=241 DEFAULT CHARSET=latin1 COMMENT='Catalogo de usuarios';

/*Data for the table `cat_usuarios` */

insert  into `cat_usuarios`(`id`,`secuencia`,`referencia`,`nombre`,`clave`,`prefijo`,`notas`,`rol`,`politica`,`operacion`,`linea`,`maquina`,`area`,`imagen`,`admin`,`calidad`,`reversos`,`programacion`,`inventario`,`tecnico`,`planta`,`apagar_alertas`,`ver_alertas`,`cerrar_al_ejecutar`,`vista_resumida_fallas`,`confirmar_reparacion`,`ultima_pantalla`,`inicializada`,`preferencias_andon`,`departamento`,`compania`,`tema`,`turno`,`claveant1`,`claveant2`,`claveant3`,`claveant4`,`claveant5`,`ucambio`,`entrada`,`salida`,`ultimo_reporte`,`idioma`,`planta_defecto`,`plantas`,`estatus`,`creacion`,`modificacion`,`creado`,`modificado`) values (1,1,'ADMIN','ADMINISTRADOR DEL SISTEMA','ž¾€\\—ÎeJÛy¥~Î-T^','','','A',1,'S','S','S','S','','S','S','S','N','N','N',1,'S','S','N','S','N',4,'N','10000110010011001000101011111011111101111511111111',8,0,1,0,'','','','','','2020-06-22','2020-08-01 00:42:15','2020-07-21 09:41:21',4,0,0,'S','A','2019-06-30 23:52:20','2020-01-02 16:18:22',1,1),(2,2,'TECNICO','TECNICO DE LA APLICACION','ž¾€\\—ÎeJÛy¥~Î-T^',NULL,NULL,'T',1,'N','S','S','S',NULL,'N','N','N','N','N','S',1,'N','N','N',NULL,'N',0,'N','11010000000000000000000000000000000000001100000000',0,0,1,0,'ÅJe!Y€åYæQ','190612',NULL,NULL,NULL,'2020-02-12','2020-07-22 09:48:31','2020-07-21 13:40:34',1,0,0,'S','A',NULL,NULL,0,0),(3,3,'GESTION','GESTOR DE LA APLICACION','ž¾€\\—ÎeJÛy¥~Î-T^',NULL,'','G',0,'S','S','S','S','','N','N','N','N','N','N',1,'N','N','S',NULL,'N',NULL,NULL,'11111111111111111111',8,0,1,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,0,0,'S','A','2019-07-08 09:37:25','2020-01-12 13:54:51',1,1),(4,4,'SUPERVISOR','SUPERVISOR DE LA APLICACION','ž¾€\\—ÎeJÛy¥~Î-T^',NULL,'','S',1,'S','S','S','S','','N','N','N','N','N','S',1,'N','N','N',NULL,'N',NULL,'N','000000',8,0,1,0,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,0,0,0,'S','A','2019-09-03 02:01:28','2020-02-13 16:55:17',1,1);

/*Table structure for table `cat_variables` */

DROP TABLE IF EXISTS `cat_variables`;

CREATE TABLE `cat_variables` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(20) DEFAULT NULL,
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'Agrupador (1)',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `maquinas` char(1) DEFAULT 'S',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `unidad` bigint(20) DEFAULT '0' COMMENT 'ID de MMCall',
  `recipiente` bigint(20) DEFAULT '0' COMMENT 'ID del recipiente',
  `tipo_valor` int(1) DEFAULT '0',
  `alarma_binaria` int(1) DEFAULT '0',
  `minimo` decimal(25,10) DEFAULT NULL,
  `maximo` decimal(25,10) DEFAULT NULL,
  `por_defecto` varchar(50) DEFAULT NULL,
  `requerida` char(1) DEFAULT 'N',
  `acumular` char(1) DEFAULT 'N',
  `sensor` int(10) DEFAULT '0',
  `sobrescribir` char(1) DEFAULT 'N',
  `reiniciar` char(1) DEFAULT 'N',
  `tope` decimal(25,10) DEFAULT '0.0000000000',
  `reiniciar_en` decimal(25,10) DEFAULT '0.0000000000',
  `van` decimal(25,10) DEFAULT '0.0000000000',
  `fecha_reiniico` datetime DEFAULT NULL,
  `fecha_tope` datetime DEFAULT NULL,
  `alarmado_desde` datetime DEFAULT NULL,
  `alarmado` char(1) DEFAULT 'N',
  `ultima_lectura_fecha` datetime DEFAULT NULL,
  `ultima_lectura_valor` decimal(25,10) DEFAULT NULL,
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `Index01` (`url_mmcall`)
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

/*Data for the table `cat_variables` */

/*Table structure for table `causa_raiz` */

DROP TABLE IF EXISTS `causa_raiz`;

CREATE TABLE `causa_raiz` (
  `reporte` bigint(20) NOT NULL DEFAULT '0' COMMENT 'ID del reporte',
  `p1` varchar(500) DEFAULT NULL,
  `p2` varchar(500) DEFAULT NULL,
  `p3` varchar(500) DEFAULT NULL,
  `p4` varchar(500) DEFAULT NULL,
  `p5` varchar(500) DEFAULT NULL,
  `plan` varchar(500) DEFAULT NULL,
  `fecha` varchar(100) DEFAULT NULL,
  `responsable` varchar(300) DEFAULT NULL,
  `departamento` varchar(300) DEFAULT NULL,
  `mano_de_obra` varchar(500) DEFAULT NULL,
  `maquina` varchar(500) DEFAULT NULL,
  `medio_ambiente` varchar(500) DEFAULT NULL,
  `metodo` varchar(500) DEFAULT NULL,
  `material` varchar(500) DEFAULT NULL,
  `comentarios` varchar(500) DEFAULT NULL,
  `creado` bigint(20) DEFAULT '0',
  `modificado` bigint(20) DEFAULT '0',
  `creacion` datetime DEFAULT NULL,
  `modificacion` datetime DEFAULT NULL,
  PRIMARY KEY (`reporte`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `causa_raiz` */

/*Table structure for table `checkeje_cab` */

DROP TABLE IF EXISTS `checkeje_cab`;

CREATE TABLE `checkeje_cab` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `checklist` bigint(20) DEFAULT '0' COMMENT 'ID del checklist',
  `plan` bigint(20) DEFAULT '0' COMMENT 'ID del plan',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `tipo` int(1) DEFAULT '0' COMMENT 'Tipo de creación (0=Automática; 1=Manual)',
  `alarmado` char(1) DEFAULT 'N' COMMENT 'Checklist alarmado',
  `alarmado_fecha` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó',
  `alarmado_variable` char(1) DEFAULT NULL COMMENT 'Checklist alarmado por variable',
  `alarmado_variable_fecha` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó por variable',
  `checklist.Enabled = False` bigint(20) DEFAULT '0',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `estado` int(2) DEFAULT '0' COMMENT 'Estado del checklist',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `iniciado` datetime DEFAULT NULL COMMENT 'Fecha de inicio',
  `finalizado` datetime DEFAULT NULL COMMENT 'Fecha de fin',
  `tiempo` bigint(6) DEFAULT '0' COMMENT 'Tiempo usado por el checklist',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=9 DEFAULT CHARSET=latin1;

/*Data for the table `checkeje_cab` */

/*Table structure for table `checkeje_det` */

DROP TABLE IF EXISTS `checkeje_det`;

CREATE TABLE `checkeje_det` (
  `checklist` bigint(20) DEFAULT '0' COMMENT 'ID del checklist',
  `variable` bigint(20) DEFAULT '0' COMMENT 'ID del variable',
  `orden` int(4) DEFAULT '0' COMMENT 'Orden en el checklist',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `valor_num` decimal(25,10) DEFAULT '0.0000000000',
  `valor_tabla` int(3) DEFAULT '0',
  `valor_directo` varchar(50) DEFAULT NULL,
  `alarmada` char(1) DEFAULT NULL,
  `fecha_alarma` datetime DEFAULT NULL,
  KEY `checklist` (`checklist`,`variable`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `checkeje_det` */

/*Table structure for table `configuracion` */

DROP TABLE IF EXISTS `configuracion`;

CREATE TABLE `configuracion` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` varchar(100) DEFAULT NULL COMMENT 'Nombre de la planta',
  `rfc` varchar(30) DEFAULT NULL COMMENT 'RFC',
  `licencia` varchar(50) DEFAULT NULL COMMENT 'Número de licencia',
  `tiempo` bigint(8) DEFAULT NULL COMMENT 'Tiempo de revisión',
  `correo_cuenta` varchar(100) DEFAULT NULL COMMENT 'Perfil de correo',
  `correo_puerto` varchar(100) DEFAULT NULL COMMENT 'Puerto',
  `correo_ssl` char(1) DEFAULT NULL COMMENT 'Seguridad SSL',
  `correo_clave` varchar(100) DEFAULT NULL COMMENT 'Contraseña',
  `correo_host` varchar(100) DEFAULT NULL COMMENT 'Host',
  `flag_agregar` char(1) DEFAULT NULL COMMENT 'Flag de que se agregó una falla',
  `ejecutando_desde` datetime DEFAULT NULL COMMENT 'Ejecutando desde',
  `ultima_falla` bigint(20) DEFAULT NULL COMMENT 'Último ID de falla revisado',
  `ultima_revision` datetime DEFAULT NULL COMMENT 'Fecha de la última revisión',
  `revisar_cada` bigint(8) DEFAULT '0' COMMENT 'Revisar cada n segundos',
  `utilizar_arduino` char(1) DEFAULT NULL COMMENT 'Usar arduino?',
  `puerto_comm1` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (1)',
  `puerto_comm1_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (1)',
  `puerto_comm2` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (2)',
  `puerto_comm2_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (2)',
  `puerto_comm3` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (3)',
  `puerto_comm3_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (3)',
  `puerto_comm4` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (4)',
  `puerto_comm4_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (4)',
  `puerto_comm5` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (5)',
  `puerto_comm5_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (5)',
  `puerto_comm6` varchar(10) DEFAULT NULL COMMENT 'Puerto comm (6)',
  `puerto_comm6_par` varchar(100) DEFAULT NULL COMMENT 'Parámetros Puerto comm (6)',
  `ruta_sms` varchar(500) DEFAULT NULL COMMENT 'Ruta para los SMS',
  `ruta_audios` varchar(500) DEFAULT NULL COMMENT 'Ruta para las llamadas',
  `optimizar_llamada` char(1) DEFAULT NULL COMMENT 'Optimiza las llamadas',
  `optimizar_sms` char(1) DEFAULT NULL COMMENT 'Optimiza los SMS',
  `optimizar_correo` char(1) DEFAULT NULL COMMENT 'Optimiza los correos',
  `optimizar_mmcall` char(1) DEFAULT NULL COMMENT 'Optimiza las llamadas a MMCall',
  `mantener_prioridad` char(1) DEFAULT NULL COMMENT 'Mantener prioridad en la optimización',
  `voz_predeterminada` varchar(255) DEFAULT NULL COMMENT 'Voz predeterminada',
  `escape_mmcall` char(1) DEFAULT NULL COMMENT 'Escape para MMCall',
  `escape_mmcall_mensaje` varchar(200) DEFAULT NULL COMMENT 'Mensaje a enviar MMCall',
  `escape_mmcall_lista` bigint(11) DEFAULT NULL COMMENT 'Lista de distribución (requesters ocupados)',
  `escape_mmcall_cancelar` char(1) DEFAULT NULL COMMENT 'Cancelar el llamado a MMCall',
  `escape_llamadas` int(1) DEFAULT NULL COMMENT 'Número de veces a llamar',
  `escape_accion` char(1) DEFAULT NULL COMMENT 'Acción a tomar',
  `escape_lista` bigint(11) DEFAULT NULL COMMENT 'Lista de distribución',
  `escape_mensaje` varchar(200) DEFAULT NULL COMMENT 'Mensaje a enviar',
  `escape_mensaje_propio` char(1) DEFAULT NULL COMMENT 'Enviar mensaje al propio móvil',
  `veces_reproducir` int(1) DEFAULT '1' COMMENT 'Número de veces que se repeduce un audio',
  `gestion_log` char(6) DEFAULT NULL COMMENT 'Año y mes de la última gestión',
  `gestion_meses` int(4) DEFAULT NULL COMMENT 'Número de meses en línea',
  `correo_titulo_falla` char(1) DEFAULT NULL COMMENT 'Mantener el título de la falla',
  `correo_titulo` varchar(100) DEFAULT NULL COMMENT 'Título opcional del correo',
  `correo_cuerpo` varchar(200) DEFAULT NULL COMMENT 'Cuerpo del correo',
  `correo_firma` varchar(100) DEFAULT NULL COMMENT 'Firma del correo',
  `timeout_llamadas` int(4) DEFAULT NULL COMMENT 'Time Out para llamadas',
  `timeout_sms` int(4) DEFAULT NULL COMMENT 'Time Out para SMS',
  `traducir` char(1) DEFAULT NULL COMMENT 'Traducir mensajes de voz',
  `tiempo_corte` bigint(8) DEFAULT '0' COMMENT 'Tiempo del corte en minutos',
  `ultimo_corte` datetime DEFAULT NULL COMMENT 'Fecha y hora del último corte',
  `bajo_hasta` int(3) DEFAULT NULL,
  `bajo_color` varchar(20) DEFAULT NULL,
  `bajo_etiqueta` varchar(30) DEFAULT NULL,
  `medio_hasta` int(3) DEFAULT NULL,
  `medio_color` varchar(20) DEFAULT NULL,
  `medio_etiqueta` varchar(30) DEFAULT NULL,
  `alto_color` varchar(20) DEFAULT NULL,
  `alto_etiqueta` varchar(30) DEFAULT NULL,
  `noatendio_color` varchar(20) DEFAULT NULL,
  `noatendio_etiqueta` varchar(30) DEFAULT NULL,
  `escaladas_color` varchar(20) DEFAULT NULL,
  `escaladas_etiqueta` varchar(30) DEFAULT NULL,
  `flag_monitor` char(1) DEFAULT 'N' COMMENT 'Flag para leer desde el monitor',
  `ruta_archivos_enviar` varchar(500) DEFAULT NULL COMMENT 'Ruta de los archivos a enviar por correo',
  `server_mmcall` varchar(100) DEFAULT NULL COMMENT 'Server para MMCall',
  `cad_consolidado` varchar(20) DEFAULT NULL COMMENT 'Cadena de la consolidado',
  `ruta_imagenes` varchar(500) DEFAULT NULL COMMENT 'Ruta de imágenes',
  `tiempo_imagen` int(4) DEFAULT NULL COMMENT 'Tiempo entre imagenes',
  `graficas_seleccion` varchar(100) DEFAULT NULL COMMENT 'Gráficas a reportar',
  `graficas_duracion` varchar(100) DEFAULT NULL,
  `timeout_fallas` int(10) DEFAULT '0' COMMENT 'Timeout para crear alerta',
  `avisar_segundos` bigint(4) DEFAULT NULL COMMENT 'Avisar con tantos segundos antes',
  `color_aviso` varchar(20) DEFAULT NULL COMMENT 'Color del aviso',
  `contar_post` char(1) DEFAULT NULL COMMENT 'Contar luego de vencer el tiempo',
  `color_post` varchar(20) DEFAULT NULL COMMENT 'Color del post',
  `escaner_prefijo` varchar(10) DEFAULT NULL COMMENT 'Prefijo del escaner',
  `escaner_sufijo` varchar(10) DEFAULT NULL COMMENT 'Sufijo del escaner',
  `tiempo_holgura` int(4) DEFAULT '0',
  `tiempo_entre_lecturas` int(4) DEFAULT NULL COMMENT 'Tiempo entre lecturas (seg)',
  `tiempo_escaner` int(4) DEFAULT '0' COMMENT 'Tiempo de espera entre milesegundos',
  `largo_escaner` int(2) DEFAULT '0' COMMENT 'Largo mínimo de la frase del escaner',
  `lote_inspeccion_clave` char(1) DEFAULT NULL COMMENT 'Requiere clave el envío de lotes a calidad',
  `reverso_permitir` char(1) DEFAULT NULL COMMENT 'Permitir reverso? (S/N/C)',
  `reverso_referencia` varchar(20) DEFAULT NULL COMMENT 'Referencia para reversar',
  `dias_programacion` int(4) DEFAULT '0' COMMENT 'Días atras para la programación',
  `holgura_reprogramar` int(6) DEFAULT '0' COMMENT 'Holgura en segundos para reprogramar',
  `tipo_flujo` char(1) DEFAULT NULL COMMENT 'Tipo de flujo',
  `estimado_productividad` int(3) DEFAULT NULL COMMENT 'Estimado productividad',
  `confirmar_mensaje_mantto` char(1) DEFAULT 'S' COMMENT 'Confirmar mensaje de configuración',
  `url_mmcall` varchar(1000) DEFAULT NULL COMMENT 'URL de MMCall',
  `accion_mmcall` char(1) DEFAULT 'N' COMMENT 'Acumula MMCall',
  `tiempo_reporte` bigint(12) DEFAULT NULL COMMENT 'Tiempo para cerrar el reporte',
  `be_alarmas_correos` char(1) DEFAULT 'N' COMMENT 'Enviar alarmas por correo',
  `be_alarmas_mmcall` char(1) DEFAULT 'N' COMMENT 'Enviar alarmas por MMCall',
  `be_alarmas_llamadas` char(1) DEFAULT 'N' COMMENT 'Enviar alarmas por llamada',
  `be_alarmas_sms` char(1) DEFAULT 'N' COMMENT 'Enviar alarmas por SMS',
  `be_envio_reportes` char(1) DEFAULT 'N' COMMENT 'Enviar reportes',
  `be_revision_correos` int(6) DEFAULT '0' COMMENT 'Envíos de correos (segundos)',
  `be_revision_mmcall` int(6) DEFAULT '0' COMMENT 'Envíos de MMCall (segundos)',
  `be_revision_arduino` int(6) DEFAULT '0' COMMENT 'Envíos de Arduino (segundos)',
  `be_log_lineas` int(4) DEFAULT '0' COMMENT 'Líneas a visulizar en en log',
  `be_log_activar` char(1) DEFAULT 'N' COMMENT 'Activar el log?',
  `maximo_largo_mmcall` int(2) DEFAULT '0' COMMENT 'Máximo de caracteres para mensajes a MMCall',
  `separador_mail` char(1) DEFAULT ';' COMMENT 'Separador de correos',
  `limitar_inicio` int(4) DEFAULT '0' COMMENT 'Segundos a limitar el inicio de sesión',
  `limitar_respuestas` int(4) DEFAULT '0' COMMENT 'Segundos a limitar respuestas',
  `recuperar_sesion` char(1) DEFAULT 'N' COMMENT 'Recuperar sesión luego de téenico',
  `visor_revisar_cada` int(6) DEFAULT '0' COMMENT 'Visor: Segundos para revisar',
  `logo_arriba` int(6) DEFAULT '0' COMMENT 'Arriba del logo',
  `logo_izquierda` int(6) DEFAULT '0' COMMENT 'Izquierda del logo',
  `logo_ancho` int(6) DEFAULT '0' COMMENT 'Ancho del logo',
  `logo_alto` int(6) DEFAULT '0' COMMENT 'Alto del logo',
  `logo_ruta` varchar(500) DEFAULT NULL COMMENT 'Ruta del logo',
  `mapa_alineacion` varchar(30) DEFAULT NULL COMMENT 'Tipo de alineación del mapa',
  `mapa_fondo` varchar(20) DEFAULT NULL COMMENT 'Color de fondo del mapa',
  `confirmar_reparacion` char(1) DEFAULT 'N' COMMENT 'Confirmar reparación',
  `ruta_programa_mapa` varchar(1000) DEFAULT NULL COMMENT 'Programa para actualizar el mapa',
  `mapa_delay` int(4) DEFAULT NULL COMMENT 'Segundos a mostrar múltiples mapas',
  `mapa_rotacion` int(1) DEFAULT '0' COMMENT 'Muestra de mapas (0: Solo con falla, 1: todos)',
  `tema_principal` int(3) DEFAULT NULL COMMENT 'ID del tema principal',
  `mapa_ultimo` datetime DEFAULT NULL COMMENT 'Fecha de la última actualización del mapa',
  `mapa_solicitud` char(1) DEFAULT NULL COMMENT 'Solicitar actualización de mapa',
  `tema_permitir_crear` char(1) DEFAULT NULL COMMENT 'Permitir a los usuarios crear temas',
  `tema_permitir_cambio` char(1) DEFAULT NULL COMMENT 'Permitir a los usuarios cambiar temas',
  `turno_modo` int(1) DEFAULT NULL COMMENT 'Modo de cambiar de turno (0: Manual, 1: Sugerir, 2: Prompt, 3: Automático)',
  `ver_logo_cronos` char(1) DEFAULT 'S' COMMENT 'Visualizar el logo de cronos',
  `url_cronos` varchar(300) DEFAULT NULL COMMENT 'URL de la página de cronos',
  `audios_activar` char(1) DEFAULT 'N' COMMENT 'Activar generación de audios generales',
  `audios_ruta` varchar(1000) DEFAULT NULL COMMENT 'Ruta a guardar',
  `audios_prefijo` varchar(1000) DEFAULT NULL COMMENT 'Ruta del audio prefijo',
  `ver_ayuda` char(1) DEFAULT 'S' COMMENT 'Ver ayuda en andon',
  `ip_localhost` varchar(100) DEFAULT NULL COMMENT 'IP local (primaria)',
  `mensaje` varchar(300) DEFAULT NULL COMMENT 'Mensaje a enviar como audio',
  `ultimo_audio` datetime DEFAULT NULL COMMENT 'Fecha de la ultima generación de audios',
  `audios_escalamiento` int(1) DEFAULT '0' COMMENT 'Modo de escalamiento de los audios',
  `audios_resolucion` varchar(100) DEFAULT NULL COMMENT 'Informar la resolución',
  `usar_clave_falla` char(1) DEFAULT NULL COMMENT 'Usar clave de falla',
  `mostrar_numero` char(1) DEFAULT NULL COMMENT 'Mostrar el número de reporte',
  `turno_actual` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `dimension` varchar(20) DEFAULT NULL COMMENT 'Dimension de la aplicación',
  `tiempo_andon` int(6) DEFAULT '0' COMMENT 'Tiempo de espera etre repeticionesANDOn',
  `lectura_pendiente` int(1) DEFAULT '0' COMMENT 'Lectura pendiente?',
  `ver_nombre_planta` char(1) DEFAULT NULL COMMENT 'Ver titulo del mapa',
  `oee_mostrar_paro` char(1) DEFAULT 'N' COMMENT 'Mostrar mensaje de paros',
  `carrusel_oee` int(1) DEFAULT '0' COMMENT 'Tipo de carrusel OEE (0=Todas las máquinas, 1=Sólo maquinas produciendo)',
  `carrusel_tiempo` bigint(6) DEFAULT '0' COMMENT 'Tiempo de visualización de la máquina en segundos',
  `esperado_oee` decimal(6,2) DEFAULT '0.00',
  `esperado_ftq` decimal(6,2) DEFAULT '0.00',
  `esperado_efi` decimal(6,2) DEFAULT '0.00',
  `esperado_dis` decimal(6,2) DEFAULT '0.00',
  `esperado_mttr` decimal(15,5) DEFAULT '0.00000',
  `esperado_mtbf` decimal(15,5) DEFAULT '0.00000',
  `reportes_inicial` int(1) DEFAULT '0' COMMENT 'Automaticamente ver todos',
  `turno_oee` bigint(20) DEFAULT '0',
  `pagers_val` char(1) DEFAULT NULL COMMENT 'S',
  `permitir_multiples_reportes` char(1) DEFAULT 'N',
  `ver_reportes_final` char(1) DEFAULT 'N',
  `andon_prorrateado` char(1) DEFAULT 'S',
  `audios_repeticiones` int(3) DEFAULT '0',
  `tiempo_audios` int(4) DEFAULT '0',
  `audio_rate` int(2) DEFAULT '0',
  `audios_externos` char(1) DEFAULT NULL,
  `audios_externos_carpeta` varchar(500) DEFAULT NULL,
  `audios_externos_modo` int(2) DEFAULT '0',
  `audios_externos_pausa` int(4) DEFAULT '0',
  `programacion` bigint(10) DEFAULT '0' COMMENT 'Programacion',
  `hibrido_mostrar_reparacion` char(1) DEFAULT 'N',
  `hibrido_alarmar_ubicacion` char(1) DEFAULT 'N',
  `hibrido_alarmar_reparacion` char(1) DEFAULT 'N',
  `mail_vencimiento` varchar(200) DEFAULT NULL COMMENT 'Mails para notificar el vencimiento',
  `finalizar_sesion` int(4) DEFAULT '0',
  `wip_salto_adelante` char(1) DEFAULT 'S',
  `wip_salto_detras` char(1) DEFAULT 'N',
  `permitir_afectacion` char(1) DEFAULT 'N',
  `idioma_defecto` int(4) DEFAULT '0' COMMENT 'Idioma por defecto',
  `area_change` bigint(20) DEFAULT '0' COMMENT 'ID del área para COver',
  `tipo_change` bigint(20) DEFAULT '0' COMMENT 'ID del concepto para COver',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

/*Data for the table `configuracion` */

insert  into `configuracion`(`id`,`planta`,`rfc`,`licencia`,`tiempo`,`correo_cuenta`,`correo_puerto`,`correo_ssl`,`correo_clave`,`correo_host`,`flag_agregar`,`ejecutando_desde`,`ultima_falla`,`ultima_revision`,`revisar_cada`,`utilizar_arduino`,`puerto_comm1`,`puerto_comm1_par`,`puerto_comm2`,`puerto_comm2_par`,`puerto_comm3`,`puerto_comm3_par`,`puerto_comm4`,`puerto_comm4_par`,`puerto_comm5`,`puerto_comm5_par`,`puerto_comm6`,`puerto_comm6_par`,`ruta_sms`,`ruta_audios`,`optimizar_llamada`,`optimizar_sms`,`optimizar_correo`,`optimizar_mmcall`,`mantener_prioridad`,`voz_predeterminada`,`escape_mmcall`,`escape_mmcall_mensaje`,`escape_mmcall_lista`,`escape_mmcall_cancelar`,`escape_llamadas`,`escape_accion`,`escape_lista`,`escape_mensaje`,`escape_mensaje_propio`,`veces_reproducir`,`gestion_log`,`gestion_meses`,`correo_titulo_falla`,`correo_titulo`,`correo_cuerpo`,`correo_firma`,`timeout_llamadas`,`timeout_sms`,`traducir`,`tiempo_corte`,`ultimo_corte`,`bajo_hasta`,`bajo_color`,`bajo_etiqueta`,`medio_hasta`,`medio_color`,`medio_etiqueta`,`alto_color`,`alto_etiqueta`,`noatendio_color`,`noatendio_etiqueta`,`escaladas_color`,`escaladas_etiqueta`,`flag_monitor`,`ruta_archivos_enviar`,`server_mmcall`,`cad_consolidado`,`ruta_imagenes`,`tiempo_imagen`,`graficas_seleccion`,`graficas_duracion`,`timeout_fallas`,`avisar_segundos`,`color_aviso`,`contar_post`,`color_post`,`escaner_prefijo`,`escaner_sufijo`,`tiempo_holgura`,`tiempo_entre_lecturas`,`tiempo_escaner`,`largo_escaner`,`lote_inspeccion_clave`,`reverso_permitir`,`reverso_referencia`,`dias_programacion`,`holgura_reprogramar`,`tipo_flujo`,`estimado_productividad`,`confirmar_mensaje_mantto`,`url_mmcall`,`accion_mmcall`,`tiempo_reporte`,`be_alarmas_correos`,`be_alarmas_mmcall`,`be_alarmas_llamadas`,`be_alarmas_sms`,`be_envio_reportes`,`be_revision_correos`,`be_revision_mmcall`,`be_revision_arduino`,`be_log_lineas`,`be_log_activar`,`maximo_largo_mmcall`,`separador_mail`,`limitar_inicio`,`limitar_respuestas`,`recuperar_sesion`,`visor_revisar_cada`,`logo_arriba`,`logo_izquierda`,`logo_ancho`,`logo_alto`,`logo_ruta`,`mapa_alineacion`,`mapa_fondo`,`confirmar_reparacion`,`ruta_programa_mapa`,`mapa_delay`,`mapa_rotacion`,`tema_principal`,`mapa_ultimo`,`mapa_solicitud`,`tema_permitir_crear`,`tema_permitir_cambio`,`turno_modo`,`ver_logo_cronos`,`url_cronos`,`audios_activar`,`audios_ruta`,`audios_prefijo`,`ver_ayuda`,`ip_localhost`,`mensaje`,`ultimo_audio`,`audios_escalamiento`,`audios_resolucion`,`usar_clave_falla`,`mostrar_numero`,`turno_actual`,`dimension`,`tiempo_andon`,`lectura_pendiente`,`ver_nombre_planta`,`oee_mostrar_paro`,`carrusel_oee`,`carrusel_tiempo`,`esperado_oee`,`esperado_ftq`,`esperado_efi`,`esperado_dis`,`esperado_mttr`,`esperado_mtbf`,`reportes_inicial`,`turno_oee`,`pagers_val`,`permitir_multiples_reportes`,`ver_reportes_final`,`andon_prorrateado`,`audios_repeticiones`,`tiempo_audios`,`audio_rate`,`audios_externos`,`audios_externos_carpeta`,`audios_externos_modo`,`audios_externos_pausa`,`programacion`,`hibrido_mostrar_reparacion`,`hibrido_alarmar_ubicacion`,`hibrido_alarmar_reparacion`,`mail_vencimiento`,`finalizar_sesion`,`wip_salto_adelante`,`wip_salto_detras`,`permitir_afectacion`,`idioma_defecto`,`area_change`,`tipo_change`) values (1,'CRONOS INTEGRACION S.A. DE C.V.','AAA900101A1A','6FC23EBAB-738F18903-B95C11D0-6DD1B9989F',2,'elvismontezuma@hotmail.com','587','S','Montezum@1974','smtp.live.com','N','2020-07-20 10:03:36',19365,'2019-06-19 22:04:46',2,'S','COM10','19200,8,0,1,2,S','COM3','null','null','null','','null','null','null','null','','C:/Users/PC/Documents/vw monitorSMS','C:/Users/PC/Documents/vw monitoraudios','N','N','S','N','N','DEFAULT','S','TODOS LOS REQUESTERS OCUPADOS',0,'S',3,'N',12,'SE GENERARON 3 LLAMADAS SIN CONTESTAR AL NUMERO ','S',2,'201911',0,'S','Monitor de WIP v1.0','Este mensaje se envía desde una aplicación automática, por favor no lo responda ya que esta cuenta de correos no se revisa.','Gracias por su soporte!',30,10,'S',NULL,'2019-07-12 16:54:53',70,'e8503c','BAJO',85,'F5B041','REGULAR','229954','En tiempo','tomato','Abiertas','orange','Escaladas','S','C:/Users/PC/Documents/vw monitorenviar','http://localhost:8081','NAVE VW 102','C:\\Users\\lap\\Documents\\vw\\carrusel',10,'S;S;S;S;N;N','2;2;2;2;2;2',10,3,'tomato','S','red','','',900,15,3000,0,'S','C','REV1234',15,1800,'C',75,'S','http://localhost:8081/locations/integration/page/number=1','N',300,'S','S','S','S','S',0,0,0,NULL,'S',0,';',0,10,'S',10,8,5,200,50,'./assets/logo.png','xMidYMid meet','FFFFFF','S','""C:\Program Files (x86)\Cronos\SIGMA"" ""server=localhost;database=sigma;user id=root;password=usbw;port=3307;Convert Zero Datetime=True"" ""C:/Users/Dell E7490/Documents/mapas/tenneco2"" ""C:/Users/Public/Documents/MMCall/root/sigma/assets/mapas""',10,1,2,'2020-07-28 21:06:34','A','S','S',3,'S','https://mmcallmexico.mx','N','C:/Users/lap/Documents/ANDON AUDIOS/AREAS','','S','localhost:8081','En la máquina [2] se presentó la falla [4]',NULL,1,'El reporte [0] ha sido resuelto','S','I',0,NULL,90,0,'S','S',0,10,'0.00','0.00','0.00','0.00','2.00000','24.00000',0,1,'N','N','S','S',0,0,0,'n','',0,0,0,'N','N','N',NULL,0,'S','N','N',10,6,47);

/*Table structure for table `consultas_be` */

DROP TABLE IF EXISTS `consultas_be`;

CREATE TABLE `consultas_be` (
  `id` bigint(10) NOT NULL COMMENT 'ID de la consulta',
  `version` varchar(30) NOT NULL COMMENT 'ID de la licencia',
  `orden` int(4) DEFAULT NULL COMMENT 'Orden de la ejecución',
  `sentencia` varchar(5000) DEFAULT NULL COMMENT 'Sentencia a aplicar',
  `estatus` char(1) NOT NULL DEFAULT 'A' COMMENT 'Estatus de la licencia (0=Por aplicar, 1=Aplicado, 9=Error)',
  PRIMARY KEY (`id`,`version`,`estatus`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `consultas_be` */

/*Table structure for table `consultas_cab` */

DROP TABLE IF EXISTS `consultas_cab`;

CREATE TABLE `consultas_cab` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `nombre` varchar(50) COLLATE utf8_bin DEFAULT NULL COMMENT 'Nombre de la consulta',
  `usuario` bigint(20) DEFAULT NULL COMMENT 'ID del usuario',
  `publico` char(1) COLLATE utf8_bin DEFAULT NULL COMMENT 'Es una consulta pública?',
  `periodo` char(2) COLLATE utf8_bin DEFAULT NULL COMMENT 'Período a usar',
  `desde` datetime DEFAULT NULL COMMENT 'Fecha desde',
  `hasta` datetime DEFAULT NULL COMMENT 'Fecha de hasta',
  `defecto` char(1) COLLATE utf8_bin DEFAULT NULL COMMENT 'Consulta por defecto',
  `filtrooper` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total operaciones filtradas',
  `filtronpar` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrolin` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtromaq` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtroare` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrofal` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtromq1` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtromq2` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtromq3` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrofa1` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrofa2` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrotec` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Total Números de parte filtradas',
  `filtrotur` char(1) COLLATE utf8_bin DEFAULT 'N',
  `filtroord` char(1) COLLATE utf8_bin DEFAULT 'N',
  `filtropar` char(1) COLLATE utf8_bin DEFAULT 'N',
  `filtrocla` char(1) COLLATE utf8_bin DEFAULT 'N',
  `filtroori` char(1) COLLATE utf8_bin DEFAULT 'N',
  `visualizar` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Visualizar todo',
  `general` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Consulta general',
  `actualizacion` datetime NOT NULL,
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`usuario`),
  KEY `NewIndex2` (`usuario`,`general`)
) ENGINE=MyISAM AUTO_INCREMENT=570 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de consultas';

/*Data for the table `consultas_cab` */

/*Table structure for table `consultas_det` */

DROP TABLE IF EXISTS `consultas_det`;

CREATE TABLE `consultas_det` (
  `consulta` bigint(20) DEFAULT NULL COMMENT 'ID de la consulta',
  `tabla` int(6) DEFAULT NULL COMMENT 'ID de la tabla',
  `valor` bigint(20) DEFAULT NULL COMMENT 'ID del valor',
  KEY `NewIndex1` (`consulta`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Detalle de la consulta';

/*Data for the table `consultas_det` */

/*Table structure for table `control` */

DROP TABLE IF EXISTS `control`;

CREATE TABLE `control` (
  `fecha` varchar(10) NOT NULL COMMENT 'Fecha  y hora del envío',
  `mensajes` int(8) DEFAULT NULL COMMENT 'Mensajes enviados',
  PRIMARY KEY (`fecha`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Control de mensajes enviados';

/*Data for the table `control` */

/*Table structure for table `control_listas` */

DROP TABLE IF EXISTS `control_listas`;

CREATE TABLE `control_listas` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `lista` bigint(11) DEFAULT NULL COMMENT 'ID de la lista',
  `maestro` bigint(11) DEFAULT NULL COMMENT 'ID del maestro de listas',
  `cierre` datetime DEFAULT NULL COMMENT 'Fecha y hora de cierre',
  `usuario` bigint(11) DEFAULT NULL COMMENT 'ID del usuario',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 COMMENT='Control de listas de verificacion';

/*Data for the table `control_listas` */

/*Table structure for table `defectos` */

DROP TABLE IF EXISTS `defectos`;

CREATE TABLE `defectos` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `defecto` bigint(11) DEFAULT NULL COMMENT 'ID del defecto',
  `retrabajo` char(1) DEFAULT NULL COMMENT 'Se retrabajaron',
  `automatico` char(1) DEFAULT NULL COMMENT 'El defecto fue automático',
  `equipo` bigint(11) DEFAULT NULL COMMENT 'ID del equipo',
  `turno` bigint(11) DEFAULT NULL COMMENT 'ID del turno',
  `tripulacion` bigint(11) DEFAULT NULL COMMENT 'ID de la tripulación',
  `orden` bigint(11) DEFAULT NULL COMMENT 'ID de la orden de producción',
  `material` bigint(11) DEFAULT NULL COMMENT 'ID del material',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha de registro',
  `cantidad` decimal(17,7) DEFAULT NULL COMMENT 'Cantidad de piezas defectuosas',
  `comentarios` varchar(1000) DEFAULT NULL COMMENT 'Comentarios del sistema',
  `atendido` char(1) DEFAULT NULL COMMENT 'El paro fue atendido?',
  `clasificado_por` bigint(11) DEFAULT NULL COMMENT 'Usuario que clasificó el paro',
  `clasificado_fecha` datetime DEFAULT NULL COMMENT 'Fecha de la clasificación',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Defectos sucedidos';

/*Data for the table `defectos` */

/*Table structure for table `defectos_info` */

DROP TABLE IF EXISTS `defectos_info`;

CREATE TABLE `defectos_info` (
  `corte` bigint(20) DEFAULT '0' COMMENT 'ID del corte',
  `piezas` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Tiempo del paro',
  `defecto` bigint(20) DEFAULT '0' COMMENT 'ID del defecto',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha del defecto',
  `dia` date DEFAULT NULL COMMENT 'Dia de control',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del defecto',
  `tipo` char(1) DEFAULT 'C' COMMENT 'A=Acumulado, C=Corte',
  UNIQUE KEY `NewIndex1` (`corte`,`tipo`,`estatus`),
  KEY `NewIndex2` (`corte`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `defectos_info` */

/*Table structure for table `det_calendario` */

DROP TABLE IF EXISTS `det_calendario`;

CREATE TABLE `det_calendario` (
  `calendario` bigint(11) NOT NULL COMMENT 'ID del calendaro',
  `fecha` date NOT NULL COMMENT 'Fecha',
  `descripcion` varchar(100) DEFAULT NULL COMMENT 'Descripción',
  `imagen` varchar(100) DEFAULT NULL COMMENT 'Imagen',
  PRIMARY KEY (`calendario`,`fecha`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Detalle del calendario';

/*Data for the table `det_calendario` */

/*Table structure for table `det_checklist` */

DROP TABLE IF EXISTS `det_checklist`;

CREATE TABLE `det_checklist` (
  `checklist` bigint(20) DEFAULT '0' COMMENT 'ID del checklist',
  `variable` bigint(20) DEFAULT '0' COMMENT 'ID del variable',
  `orden` int(4) DEFAULT '0' COMMENT 'Orden en el checklist',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  KEY `checklist` (`checklist`,`variable`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `det_checklist` */

/*Table structure for table `det_correo` */

DROP TABLE IF EXISTS `det_correo`;

CREATE TABLE `det_correo` (
  `correo` bigint(20) DEFAULT NULL COMMENT 'ID del correo',
  `reporte` bigint(20) DEFAULT NULL COMMENT 'ID del reporte'
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Detalle de los coreos a enviar';

/*Data for the table `det_correo` */

/*Table structure for table `det_disponibilidad` */

DROP TABLE IF EXISTS `det_disponibilidad`;

CREATE TABLE `det_disponibilidad` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `calendario` bigint(20) DEFAULT NULL COMMENT 'ID del calendario',
  `equipo` bigint(20) DEFAULT NULL COMMENT 'ID del equipo',
  `fecha` date DEFAULT NULL COMMENT 'Fecha especifica',
  `disponibilidad` bigint(5) DEFAULT NULL COMMENT 'Disponibilidad en segundos',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`calendario`,`equipo`,`fecha`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Disponibilidad por equipo y fecha';

/*Data for the table `det_disponibilidad` */

/*Table structure for table `det_distribucion` */

DROP TABLE IF EXISTS `det_distribucion`;

CREATE TABLE `det_distribucion` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `distribucion` bigint(11) NOT NULL COMMENT 'ID de la lista de distribución',
  `orden` int(4) NOT NULL COMMENT 'Línea en la lista',
  `tipo` int(2) DEFAULT NULL COMMENT 'Tipo de lista (0: Móvil-Llamada, 10: Móvil-SMS, 30: Móvil-LLamada y SMS, 40: Correo electrónico, 50: Ärea de MMCall)',
  `cadena` varchar(255) DEFAULT NULL COMMENT 'Cadena',
  `alias` varchar(30) DEFAULT NULL COMMENT 'Alias',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus',
  PRIMARY KEY (`id`,`distribucion`,`orden`)
) ENGINE=MyISAM AUTO_INCREMENT=78 DEFAULT CHARSET=latin1 COMMENT='Detalle de la lista de distribucón';

/*Data for the table `det_distribucion` */

/*Table structure for table `det_estacion` */

DROP TABLE IF EXISTS `det_estacion`;

CREATE TABLE `det_estacion` (
  `estacion` bigint(11) NOT NULL COMMENT 'ID de la estación',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `equipo` bigint(11) NOT NULL COMMENT 'ID del equipo',
  PRIMARY KEY (`estacion`,`equipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Detalle de las estaciones';

/*Data for the table `det_estacion` */

/*Table structure for table `det_idiomas` */

DROP TABLE IF EXISTS `det_idiomas`;

CREATE TABLE `det_idiomas` (
  `idioma` int(3) NOT NULL COMMENT 'ID del idioma',
  `linea` int(3) DEFAULT NULL COMMENT 'Línea del idioma',
  `cadena` varchar(3000) DEFAULT NULL COMMENT 'Cadena (separados por coma)',
  PRIMARY KEY (`idioma`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `det_idiomas` */

/*Table structure for table `det_lista` */

DROP TABLE IF EXISTS `det_lista`;

CREATE TABLE `det_lista` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `lista` bigint(11) NOT NULL COMMENT 'ID de la lista',
  `variable` bigint(11) NOT NULL COMMENT 'ID de la variable',
  `orden` int(4) DEFAULT NULL COMMENT 'Orden en la lista',
  `deorigen` char(1) DEFAULT NULL COMMENT 'Tomar los datos de origen',
  `requerido` char(1) DEFAULT NULL COMMENT 'Campo requerido',
  `notas` varchar(500) DEFAULT NULL COMMENT 'Notas de la variable',
  `tipo` int(2) DEFAULT NULL COMMENT 'Tipo de valor (10=numérico, 20=Si/NO, 30=Tabla)',
  `tabla` char(1) DEFAULT NULL COMMENT 'Tomar valor de una tabla',
  `idtabla` bigint(11) DEFAULT NULL COMMENT 'ID de la tabla',
  `unidad` bigint(11) DEFAULT NULL COMMENT 'ID de la unidad de medida',
  `permitido_min` decimal(30,10) DEFAULT NULL COMMENT 'Valor mínimo',
  `permitido_max` decimal(30,10) DEFAULT NULL COMMENT 'Valor máximo',
  `alarma_min` decimal(30,10) DEFAULT NULL COMMENT 'Valor mínimo para generar alarma',
  `alarma_max` decimal(30,10) DEFAULT NULL COMMENT 'Valor máximo para generar alarma',
  `alarma_supervision` char(1) DEFAULT NULL COMMENT 'Requiere supervisión',
  `alarma_regla` char(1) DEFAULT NULL COMMENT 'Requiere regla',
  `alarma_sino` char(1) DEFAULT NULL COMMENT 'Alarma Si/No',
  `color` varchar(20) DEFAULT NULL COMMENT 'Color de fondo',
  `resaltada` char(1) DEFAULT NULL COMMENT 'Resaltar variable',
  `mostrar_rango` char(1) DEFAULT NULL COMMENT 'Mostrar rango en pantalla',
  `confirmar_respuesta` char(1) DEFAULT NULL COMMENT 'Confirmar la respuesta',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha en que se agregó',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha en que se modificó',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Usuario que agregó',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Usuario que modificó',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`lista`,`variable`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1 COMMENT='Detalle de listas de verificación';

/*Data for the table `det_lista` */

/*Table structure for table `det_plan_checklists` */

DROP TABLE IF EXISTS `det_plan_checklists`;

CREATE TABLE `det_plan_checklists` (
  `plan` bigint(20) DEFAULT NULL COMMENT 'ID del plan',
  `checklist` bigint(20) DEFAULT NULL COMMENT 'ID del checklist'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `det_plan_checklists` */

/*Table structure for table `det_procesos` */

DROP TABLE IF EXISTS `det_procesos`;

CREATE TABLE `det_procesos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `capacidad` bigint(12) DEFAULT '0' COMMENT 'Tiempo general de un lote en stock (segundos)',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID del proceso asociado',
  `programar` char(1) DEFAULT 'N' COMMENT 'Permitir la programación',
  `ultimo_parte` bigint(20) DEFAULT '0',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=47 DEFAULT CHARSET=latin1;

/*Data for the table `det_procesos` */

/*Table structure for table `det_regla` */

DROP TABLE IF EXISTS `det_regla`;

CREATE TABLE `det_regla` (
  `lista` bigint(11) DEFAULT NULL COMMENT 'ID de la lista',
  `variable` bigint(11) DEFAULT NULL COMMENT 'ID de la variable',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `periodo` int(2) DEFAULT NULL COMMENT 'período',
  `acumulado` int(4) DEFAULT NULL COMMENT 'Acumulado',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  KEY `NewIndex1` (`lista`),
  KEY `NewIndex2` (`variable`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Reglas de acumulación';

/*Data for the table `det_regla` */

/*Table structure for table `det_rutas` */

DROP TABLE IF EXISTS `det_rutas`;

CREATE TABLE `det_rutas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `ruta` bigint(20) DEFAULT NULL COMMENT 'ID de la ruta',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia de la operación',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'Referencia con el sistema',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `prefijo` varchar(30) DEFAULT NULL COMMENT 'Prefijo del registro',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `tiempo_stock` bigint(12) DEFAULT '0' COMMENT 'Tiempo general de un lote en stock (segundos)',
  `tiempo_proceso` bigint(12) DEFAULT '0' COMMENT 'Tiempo general de un lote en proceso (segundos)',
  `tiempo_setup` bigint(12) DEFAULT '0' COMMENT 'Tiempo de preparación',
  `tiempo_setup_idem` bigint(12) DEFAULT '0' COMMENT 'Tiempo de preparación mismo material',
  `piezas_finalizar_paro` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Piezas para finalizar paro',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID del proceso asociado',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`ruta`,`secuencia`)
) ENGINE=MyISAM AUTO_INCREMENT=750 DEFAULT CHARSET=latin1 COMMENT='Detalle de rutas';

/*Data for the table `det_rutas` */

/*Table structure for table `detalleparos` */

DROP TABLE IF EXISTS `detalleparos`;

CREATE TABLE `detalleparos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `paro` varchar(50) DEFAULT NULL COMMENT 'ID del paro',
  `reporte` bigint(20) NOT NULL DEFAULT '0',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'ID tipo de paro',
  `clase` int(2) DEFAULT '0' COMMENT 'Clase de paro (0=planeado, 1=No planeado, 2=Manual)',
  `area` bigint(20) DEFAULT '0' COMMENT 'ID del área responsable del paro',
  `masivo` char(1) DEFAULT 'N',
  `hora_inicial` varchar(8) DEFAULT NULL,
  `hora_final` varchar(8) DEFAULT NULL,
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `fecha` date DEFAULT NULL COMMENT 'Fecha para reporte',
  `desde` datetime DEFAULT NULL COMMENT 'Fecha y hora de inicio según plan',
  `hasta` datetime DEFAULT NULL COMMENT 'Fecha y hora de fin según plan',
  `tiempo` bigint(12) DEFAULT '0' COMMENT 'Tiempo real del paro',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `inicio` bigint(20) DEFAULT '0' COMMENT 'Usuario que inició',
  `finalizo` bigint(20) DEFAULT '0' COMMENT 'Usuario que finalizó',
  `inicia` datetime DEFAULT NULL COMMENT 'Fecha y hora de inicio real',
  `finaliza` datetime DEFAULT NULL COMMENT 'Fecha y hora de finalización real',
  `finalizo_accion` char(1) DEFAULT 'P' COMMENT 'Forma de finalizar (M=Manual, P=Por conteo de piezas, T=Por tiempo)',
  `contabilizar` char(1) DEFAULT 'S' COMMENT 'Contabilizar paro (S=Si, N=No, A=Por aprobar)',
  `origen` char(1) DEFAULT 'M' COMMENT 'Origen del plan (N=Normal, M=Desde la máquina)',
  `clave` varchar(20) DEFAULT NULL COMMENT 'Clave de Aprobación/Conversión',
  `aprobador` bigint(11) DEFAULT '0' COMMENT 'Usuario que aprobó la conversión',
  `finaliza_sensor` char(1) DEFAULT 'N' COMMENT 'Se finaliza el paro al contar piezas',
  `estado` char(1) DEFAULT 'L' COMMENT 'Estado del paro (P=Preparando, L=Listo, C=En curso, F=Finalizado)',
  `desde_original` datetime DEFAULT NULL,
  `hasta_original` datetime DEFAULT NULL,
  `cambiado_por` bigint(20) DEFAULT '0',
  `cambiado_fecha` datetime DEFAULT NULL,
  `cambiado_causa` bigint(20) DEFAULT '0' COMMENT 'ID de la cancelación',
  `corte` datetime DEFAULT NULL COMMENT 'Fecha y hora del corte',
  `notas` varchar(100) DEFAULT NULL COMMENT 'Detalle del paro',
  `resultados` varchar(100) DEFAULT NULL COMMENT 'Resutado del paro',
  `parte` bigint(20) DEFAULT '0',
  `lote` bigint(20) DEFAULT '0',
  `wip_piezas` decimal(25,10) DEFAULT '0.0000000000',
  `wip_van` decimal(25,10) DEFAULT '0.0000000000',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del paro',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`desde`,`hasta`),
  KEY `NewIndex2` (`maquina`,`tipo`,`estado`),
  KEY `NewIndex3` (`maquina`,`desde`,`hasta`,`estado`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=51801 DEFAULT CHARSET=latin1 COMMENT='Detalle de paros';

/*Data for the table `detalleparos` */

/*Table structure for table `detallerechazos` */

DROP TABLE IF EXISTS `detallerechazos`;

CREATE TABLE `detallerechazos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `rechazo` varchar(100) DEFAULT NULL COMMENT 'Texto del rechazo',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'ID tipo de paro',
  `area` bigint(20) DEFAULT '0' COMMENT 'ID del área responsable del paro',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `fecha` date DEFAULT NULL COMMENT 'Fecha para reporte',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `origen` int(1) DEFAULT '0' COMMENT 'Origen del plan (N=Normal, M=Desde la máquina)',
  `corte` bigint(20) DEFAULT '0' COMMENT 'ID del corte',
  `notas` varchar(100) DEFAULT NULL COMMENT 'Detalle del paro',
  `parte` bigint(20) DEFAULT '0',
  `lote` bigint(20) DEFAULT '0',
  `cantidad` decimal(25,10) DEFAULT '0.0000000000',
  `cantidad_tc` decimal(25,10) DEFAULT '0.0000000000',
  `usuario` bigint(20) DEFAULT '0' COMMENT 'ID del usuario',
  `actualizacion` datetime DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `NewIndex2` (`equipo`,`tipo`),
  KEY `NewIndex3` (`equipo`)
) ENGINE=MyISAM AUTO_INCREMENT=48009 DEFAULT CHARSET=latin1;

/*Data for the table `detallerechazos` */

/*Table structure for table `dias` */

DROP TABLE IF EXISTS `dias`;

CREATE TABLE `dias` (
  `fecha` date NOT NULL COMMENT 'Fecha del día',
  `dia` int(1) NOT NULL,
  PRIMARY KEY (`fecha`,`dia`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Tabla de dias';

/*Data for the table `dias` */

insert  into `dias`(`fecha`,`dia`) values ('2018-12-15',7),('2018-12-16',1),('2018-12-17',2),('2018-12-18',3),('2018-12-19',4),('2018-12-20',5),('2018-12-21',6),('2018-12-22',7),('2018-12-23',1),('2018-12-24',2),('2018-12-25',3),('2018-12-26',4),('2018-12-27',5),('2018-12-28',6),('2018-12-29',7),('2018-12-30',1),('2018-12-31',2),('2019-01-01',3),('2019-01-02',4),('2019-01-03',5),('2019-01-04',6),('2019-01-05',7),('2019-01-06',1),('2019-01-07',2),('2019-01-08',3),('2019-01-09',4),('2019-01-10',5),('2019-01-11',6),('2019-01-12',7),('2019-01-13',1),('2019-01-14',2),('2019-01-15',3),('2019-01-16',4),('2019-01-17',5),('2019-01-18',6),('2019-01-19',7),('2019-01-20',1),('2019-01-21',2),('2019-01-22',3),('2019-01-23',4),('2019-01-24',5),('2019-01-25',6),('2019-01-26',7),('2019-01-27',1),('2019-01-28',2),('2019-01-29',3),('2019-01-30',4),('2019-01-31',5),('2019-02-01',6),('2019-02-02',7),('2019-02-03',1),('2019-02-04',2),('2019-02-05',3),('2019-02-06',4),('2019-02-07',5),('2019-02-08',6),('2019-02-09',7),('2019-02-10',1),('2019-02-11',2),('2019-02-12',3),('2019-02-13',4),('2019-02-14',5),('2019-02-15',6),('2019-02-16',7),('2019-02-17',1),('2019-02-18',2),('2019-02-19',3),('2019-02-20',4),('2019-02-21',5),('2019-02-22',6),('2019-02-23',7),('2019-02-24',1),('2019-02-25',2),('2019-02-26',3),('2019-02-27',4),('2019-02-28',5),('2019-03-01',6),('2019-03-02',7),('2019-03-03',1),('2019-03-04',2),('2019-03-05',3),('2019-03-06',4),('2019-03-07',5),('2019-03-08',6),('2019-03-09',7),('2019-03-10',1),('2019-03-11',2),('2019-03-12',3),('2019-03-13',4),('2019-03-14',5),('2019-03-15',6),('2019-03-16',7),('2019-03-17',1),('2019-03-18',2),('2019-03-19',3),('2019-03-20',4),('2019-03-21',5),('2019-03-22',6),('2019-03-23',7),('2019-03-24',1),('2019-03-25',2),('2019-03-26',3),('2019-03-27',4),('2019-03-28',5),('2019-03-29',6),('2019-03-30',7),('2019-03-31',1),('2019-04-01',2),('2019-04-02',3),('2019-04-03',4),('2019-04-04',5),('2019-04-05',6),('2019-04-06',7),('2019-04-07',1),('2019-04-08',2),('2019-04-09',3),('2019-04-10',4),('2019-04-11',5),('2019-04-12',6),('2019-04-13',7),('2019-04-14',1),('2019-04-15',2),('2019-04-16',3),('2019-04-17',4),('2019-04-18',5),('2019-04-19',6),('2019-04-20',7),('2019-04-21',1),('2019-04-22',2),('2019-04-23',3),('2019-04-24',4),('2019-04-25',5),('2019-04-26',6),('2019-04-27',7),('2019-04-28',1),('2019-04-29',2),('2019-04-30',3),('2019-05-01',4),('2019-05-02',5),('2019-05-03',6),('2019-05-04',7),('2019-05-05',1),('2019-05-06',2),('2019-05-07',3),('2019-05-08',4),('2019-05-09',5),('2019-05-10',6),('2019-05-11',7),('2019-05-12',1),('2019-05-13',2),('2019-05-14',3),('2019-05-15',4),('2019-05-16',5),('2019-05-17',6),('2019-05-18',7),('2019-05-19',1),('2019-05-20',2),('2019-05-21',3),('2019-05-22',4),('2019-05-23',5),('2019-05-24',6),('2019-05-25',7),('2019-05-26',1),('2019-05-27',2),('2019-05-28',3),('2019-05-29',4),('2019-05-30',5),('2019-05-31',6),('2019-06-01',7),('2019-06-02',1),('2019-06-03',2),('2019-06-04',3),('2019-06-05',4),('2019-06-06',5),('2019-06-07',6),('2019-06-08',7),('2019-06-09',1),('2019-06-10',2),('2019-06-11',3),('2019-06-12',4),('2019-06-13',5),('2019-06-14',6),('2019-06-15',7),('2019-06-16',1),('2019-06-17',2),('2019-06-18',3),('2019-06-19',4),('2019-06-20',5),('2019-06-21',6),('2019-06-22',7),('2019-06-23',1),('2019-06-24',2),('2019-06-25',3),('2019-06-26',4),('2019-06-27',5),('2019-06-28',6),('2019-06-29',7),('2019-06-30',1),('2019-07-01',2),('2019-07-02',3),('2019-07-03',4),('2019-07-04',5),('2019-07-05',6),('2019-07-06',7),('2019-07-07',1),('2019-07-08',2),('2019-07-09',3),('2019-07-10',4),('2019-07-11',5),('2019-07-12',6),('2019-07-13',7),('2019-07-14',1),('2019-07-15',2),('2019-07-16',3),('2019-07-17',4),('2019-07-18',5),('2019-07-19',6),('2019-07-20',7),('2019-07-21',1),('2019-07-22',2),('2019-07-23',3),('2019-07-24',4),('2019-07-25',5),('2019-07-26',6),('2019-07-27',7),('2019-07-28',1),('2019-07-29',2),('2019-07-30',3),('2019-07-31',4),('2019-08-01',5),('2019-08-02',6),('2019-08-03',7),('2019-08-04',1),('2019-08-05',2),('2019-08-06',3),('2019-08-07',4),('2019-08-08',5),('2019-08-09',6),('2019-08-10',7),('2019-08-11',1),('2019-08-12',2),('2019-08-13',3),('2019-08-14',4),('2019-08-15',5),('2019-08-16',6),('2019-08-17',7),('2019-08-18',1),('2019-08-19',2),('2019-08-20',3),('2019-08-21',4),('2019-08-22',5),('2019-08-23',6),('2019-08-24',7),('2019-08-25',1),('2019-08-26',2),('2019-08-27',3),('2019-08-28',4),('2019-08-29',5),('2019-08-30',6),('2019-08-31',7),('2019-09-01',1),('2019-09-02',2),('2019-09-03',3),('2019-09-04',4),('2019-09-05',5),('2019-09-06',6),('2019-09-07',7),('2019-09-08',1),('2019-09-09',2),('2019-09-10',3),('2019-09-11',4),('2019-09-12',5),('2019-09-13',6),('2019-09-14',7),('2019-09-15',1),('2019-09-16',2),('2019-09-17',3),('2019-09-18',4),('2019-09-19',5),('2019-09-20',6),('2019-09-21',7),('2019-09-22',1),('2019-09-23',2),('2019-09-24',3),('2019-09-25',4),('2019-09-26',5),('2019-09-27',6),('2019-09-28',7),('2019-09-29',1),('2019-09-30',2),('2019-10-01',3),('2019-10-02',4),('2019-10-03',5),('2019-10-04',6),('2019-10-05',7),('2019-10-06',1),('2019-10-07',2),('2019-10-08',3),('2019-10-09',4),('2019-10-10',5),('2019-10-11',6),('2019-10-12',7),('2019-10-13',1),('2019-10-14',2),('2019-10-15',3),('2019-10-16',4),('2019-10-17',5),('2019-10-18',6),('2019-10-19',7),('2019-10-20',1),('2019-10-21',2),('2019-10-22',3),('2019-10-23',4),('2019-10-24',5),('2019-10-25',6),('2019-10-26',7),('2019-10-27',1),('2019-10-28',2),('2019-10-29',3),('2019-10-30',4),('2019-10-31',5),('2019-11-01',6),('2019-11-02',7),('2019-11-03',1),('2019-11-04',2),('2019-11-05',3),('2019-11-06',4),('2019-11-07',5),('2019-11-08',6),('2019-11-09',7),('2019-11-10',1),('2019-11-11',2),('2019-11-12',3),('2019-11-13',4),('2019-11-14',5),('2019-11-15',6),('2019-11-16',7),('2019-11-17',1),('2019-11-18',2),('2019-11-19',3),('2019-11-20',4),('2019-11-21',5),('2019-11-22',6),('2019-11-23',7),('2019-11-24',1),('2019-11-25',2),('2019-11-26',3),('2019-11-27',4),('2019-11-28',5),('2019-11-29',6),('2019-11-30',7),('2019-12-01',1),('2019-12-02',2),('2019-12-03',3),('2019-12-04',4),('2019-12-05',5),('2019-12-06',6),('2019-12-07',7),('2019-12-08',1),('2019-12-09',2),('2019-12-10',3),('2019-12-11',4),('2019-12-12',5),('2019-12-13',6),('2019-12-14',7),('2019-12-15',1),('2019-12-16',2),('2019-12-17',3),('2019-12-18',4),('2019-12-19',5),('2019-12-20',6),('2019-12-21',7),('2019-12-22',1),('2019-12-23',2),('2019-12-24',3),('2019-12-25',4),('2019-12-26',5),('2019-12-27',6),('2019-12-28',7),('2019-12-29',1),('2019-12-30',2),('2019-12-31',3),('2020-01-01',4),('2020-01-02',5),('2020-01-03',6),('2020-01-04',7),('2020-01-05',1),('2020-01-06',2),('2020-01-07',3),('2020-01-08',4),('2020-01-09',5),('2020-01-10',6),('2020-01-11',7),('2020-01-12',1),('2020-01-13',2),('2020-01-14',3),('2020-01-15',4),('2020-01-16',5),('2020-01-17',6),('2020-01-18',7),('2020-01-19',1),('2020-01-20',2),('2020-01-21',3),('2020-01-22',4),('2020-01-23',5),('2020-01-24',6),('2020-01-25',7),('2020-01-26',1),('2020-01-27',2),('2020-01-28',3),('2020-01-29',4),('2020-01-30',5),('2020-01-31',6),('2020-02-01',7),('2020-02-02',1),('2020-02-03',2),('2020-02-04',3),('2020-02-05',4),('2020-02-06',5),('2020-02-07',6),('2020-02-08',7),('2020-02-09',1),('2020-02-10',2),('2020-02-11',3),('2020-02-12',4),('2020-02-13',5),('2020-02-14',6),('2020-02-15',7),('2020-02-16',1),('2020-02-17',2),('2020-02-18',3),('2020-02-19',4),('2020-02-20',5),('2020-02-21',6),('2020-02-22',7),('2020-02-23',1),('2020-02-24',2),('2020-02-25',3),('2020-02-26',4),('2020-02-27',5),('2020-02-28',6),('2020-02-29',7),('2020-03-01',1),('2020-03-02',2),('2020-03-03',3),('2020-03-04',4),('2020-03-05',5),('2020-03-06',6),('2020-03-07',7),('2020-03-08',1),('2020-03-09',2),('2020-03-10',3),('2020-03-11',4),('2020-03-12',5),('2020-03-13',6),('2020-03-14',7),('2020-03-15',1),('2020-03-16',2),('2020-03-17',3),('2020-03-18',4),('2020-03-19',5),('2020-03-20',6),('2020-03-21',7),('2020-03-22',1),('2020-03-23',2),('2020-03-24',3),('2020-03-25',4),('2020-03-26',5),('2020-03-27',6),('2020-03-28',7),('2020-03-29',1),('2020-03-30',2),('2020-03-31',3),('2020-04-01',4),('2020-04-02',5),('2020-04-03',6),('2020-04-04',7),('2020-04-05',1),('2020-04-06',2),('2020-04-07',3),('2020-04-08',4),('2020-04-09',5),('2020-04-10',6),('2020-04-11',7),('2020-04-12',1),('2020-04-13',2),('2020-04-14',3),('2020-04-15',4),('2020-04-16',5),('2020-04-17',6),('2020-04-18',7),('2020-04-19',1),('2020-04-20',2),('2020-04-21',3),('2020-04-22',4),('2020-04-23',5),('2020-04-24',6),('2020-04-25',7),('2020-04-26',1),('2020-04-27',2),('2020-04-28',3),('2020-04-29',4),('2020-04-30',5),('2020-05-01',6),('2020-05-02',7),('2020-05-03',1),('2020-05-04',2),('2020-05-05',3),('2020-05-06',4),('2020-05-07',5),('2020-05-08',6),('2020-05-09',7),('2020-05-10',1),('2020-05-11',2),('2020-05-12',3),('2020-05-13',4),('2020-05-14',5),('2020-05-15',6),('2020-05-16',7),('2020-05-17',1),('2020-05-18',2),('2020-05-19',3),('2020-05-20',4),('2020-05-21',5),('2020-05-22',6),('2020-05-23',7),('2020-05-24',1),('2020-05-25',2),('2020-05-26',3),('2020-05-27',4),('2020-05-28',5),('2020-05-29',6),('2020-05-30',7),('2020-05-31',1),('2020-06-01',2),('2020-06-02',3),('2020-06-03',4),('2020-06-04',5),('2020-06-05',6),('2020-06-06',7),('2020-06-07',1),('2020-06-08',2),('2020-06-09',3),('2020-06-10',4),('2020-06-11',5),('2020-06-12',6),('2020-06-13',7),('2020-06-14',1),('2020-06-15',2),('2020-06-16',3),('2020-06-17',4),('2020-06-18',5),('2020-06-19',6),('2020-06-20',7),('2020-06-21',1),('2020-06-22',2),('2020-06-23',3),('2020-06-24',4),('2020-06-25',5),('2020-06-26',6),('2020-06-27',7),('2020-06-28',1),('2020-06-29',2),('2020-06-30',3),('2020-07-01',4),('2020-07-02',5),('2020-07-03',6),('2020-07-04',7),('2020-07-05',1),('2020-07-06',2),('2020-07-07',3),('2020-07-08',4),('2020-07-09',5),('2020-07-10',6),('2020-07-11',7),('2020-07-12',1),('2020-07-13',2),('2020-07-14',3),('2020-07-15',4),('2020-07-16',5),('2020-07-17',6),('2020-07-18',7),('2020-07-19',1),('2020-07-20',2),('2020-07-21',3),('2020-07-22',4),('2020-07-23',5),('2020-07-24',6),('2020-07-25',7),('2020-07-26',1),('2020-07-27',2),('2020-07-28',3),('2020-07-29',4),('2020-07-30',5),('2020-07-31',6),('2020-08-01',7),('2020-08-02',1),('2020-08-03',2),('2020-08-04',3),('2020-08-05',4),('2020-08-06',5),('2020-08-07',6),('2020-08-08',7),('2020-08-09',1),('2020-08-10',2),('2020-08-11',3),('2020-08-12',4),('2020-08-13',5),('2020-08-14',6),('2020-08-15',7),('2020-08-16',1),('2020-08-17',2),('2020-08-18',3),('2020-08-19',4),('2020-08-20',5),('2020-08-21',6),('2020-08-22',7),('2020-08-23',1),('2020-08-24',2),('2020-08-25',3),('2020-08-26',4),('2020-08-27',5),('2020-08-28',6),('2020-08-29',7),('2020-08-30',1),('2020-08-31',2),('2020-09-01',3),('2020-09-02',4),('2020-09-03',5),('2020-09-04',6),('2020-09-05',7),('2020-09-06',1),('2020-09-07',2),('2020-09-08',3),('2020-09-09',4),('2020-09-10',5),('2020-09-11',6),('2020-09-12',7),('2020-09-13',1),('2020-09-14',2),('2020-09-15',3),('2020-09-16',4),('2020-09-17',5),('2020-09-18',6),('2020-09-19',7),('2020-09-20',1),('2020-09-21',2),('2020-09-22',3),('2020-09-23',4),('2020-09-24',5),('2020-09-25',6),('2020-09-26',7),('2020-09-27',1),('2020-09-28',2),('2020-09-29',3),('2020-09-30',4),('2020-10-01',5),('2020-10-02',6),('2020-10-03',7),('2020-10-04',1),('2020-10-05',2),('2020-10-06',3),('2020-10-07',4),('2020-10-08',5),('2020-10-09',6),('2020-10-10',7),('2020-10-11',1),('2020-10-12',2),('2020-10-13',3),('2020-10-14',4),('2020-10-15',5),('2020-10-16',6),('2020-10-17',7),('2020-10-18',1),('2020-10-19',2),('2020-10-20',3),('2020-10-21',4),('2020-10-22',5),('2020-10-23',6),('2020-10-24',7),('2020-10-25',1),('2020-10-26',2),('2020-10-27',3),('2020-10-28',4),('2020-10-29',5),('2020-10-30',6),('2020-10-31',7),('2020-11-01',1),('2020-11-02',2),('2020-11-03',3),('2020-11-04',4),('2020-11-05',5),('2020-11-06',6),('2020-11-07',7),('2020-11-08',1),('2020-11-09',2),('2020-11-10',3),('2020-11-11',4),('2020-11-12',5),('2020-11-13',6),('2020-11-14',7),('2020-11-15',1),('2020-11-16',2),('2020-11-17',3),('2020-11-18',4),('2020-11-19',5),('2020-11-20',6),('2020-11-21',7),('2020-11-22',1),('2020-11-23',2),('2020-11-24',3),('2020-11-25',4),('2020-11-26',5),('2020-11-27',6),('2020-11-28',7),('2020-11-29',1),('2020-11-30',2),('2020-12-01',3),('2020-12-02',4),('2020-12-03',5),('2020-12-04',6),('2020-12-05',7),('2020-12-06',1),('2020-12-07',2),('2020-12-08',3),('2020-12-09',4),('2020-12-10',5),('2020-12-11',6),('2020-12-12',7),('2020-12-13',1),('2020-12-14',2),('2020-12-15',3),('2020-12-16',4),('2020-12-17',5),('2020-12-18',6),('2020-12-19',7),('2020-12-20',1),('2020-12-21',2),('2020-12-22',3),('2020-12-23',4),('2020-12-24',5),('2020-12-25',6),('2020-12-26',7),('2020-12-27',1),('2020-12-28',2),('2020-12-29',3),('2020-12-30',4),('2020-12-31',5),('2021-01-01',6),('2021-01-02',7),('2021-01-03',1),('2021-01-04',2),('2021-01-05',3),('2021-01-06',4),('2021-01-07',5),('2021-01-08',6),('2021-01-09',7),('2021-01-10',1),('2021-01-11',2),('2021-01-12',3),('2021-01-13',4),('2021-01-14',5),('2021-01-15',6),('2021-01-16',7),('2021-01-17',1),('2021-01-18',2),('2021-01-19',3),('2021-01-20',4),('2021-01-21',5),('2021-01-22',6),('2021-01-23',7),('2021-01-24',1),('2021-01-25',2),('2021-01-26',3),('2021-01-27',4),('2021-01-28',5),('2021-01-29',6),('2021-01-30',7),('2021-01-31',1),('2021-02-01',2),('2021-02-02',3),('2021-02-03',4),('2021-02-04',5),('2021-02-05',6),('2021-02-06',7),('2021-02-07',1),('2021-02-08',2),('2021-02-09',3),('2021-02-10',4),('2021-02-11',5),('2021-02-12',6),('2021-02-13',7),('2021-02-14',1),('2021-02-15',2),('2021-02-16',3),('2021-02-17',4),('2021-02-18',5),('2021-02-19',6),('2021-02-20',7),('2021-02-21',1),('2021-02-22',2),('2021-02-23',3),('2021-02-24',4),('2021-02-25',5),('2021-02-26',6),('2021-02-27',7),('2021-02-28',1),('2021-03-01',2),('2021-03-02',3),('2021-03-03',4),('2021-03-04',5),('2021-03-05',6),('2021-03-06',7),('2021-03-07',1),('2021-03-08',2),('2021-03-09',3),('2021-03-10',4),('2021-03-11',5),('2021-03-12',6),('2021-03-13',7),('2021-03-14',1),('2021-03-15',2),('2021-03-16',3),('2021-03-17',4),('2021-03-18',5),('2021-03-19',6),('2021-03-20',7),('2021-03-21',1),('2021-03-22',2),('2021-03-23',3),('2021-03-24',4),('2021-03-25',5),('2021-03-26',6),('2021-03-27',7),('2021-03-28',1),('2021-03-29',2),('2021-03-30',3),('2021-03-31',4),('2021-04-01',5),('2021-04-02',6),('2021-04-03',7),('2021-04-04',1),('2021-04-05',2),('2021-04-06',3),('2021-04-07',4),('2021-04-08',5),('2021-04-09',6),('2021-04-10',7),('2021-04-11',1),('2021-04-12',2),('2021-04-13',3),('2021-04-14',4),('2021-04-15',5),('2021-04-16',6),('2021-04-17',7),('2021-04-18',1),('2021-04-19',2),('2021-04-20',3),('2021-04-21',4),('2021-04-22',5),('2021-04-23',6),('2021-04-24',7),('2021-04-25',1),('2021-04-26',2),('2021-04-27',3),('2021-04-28',4),('2021-04-29',5),('2021-04-30',6),('2021-05-01',7),('2021-05-02',1),('2021-05-03',2),('2021-05-04',3),('2021-05-05',4),('2021-05-06',5),('2021-05-07',6),('2021-05-08',7),('2021-05-09',1),('2021-05-10',2),('2021-05-11',3),('2021-05-12',4),('2021-05-13',5),('2021-05-14',6),('2021-05-15',7),('2021-05-16',1),('2021-05-17',2),('2021-05-18',3),('2021-05-19',4),('2021-05-20',5),('2021-05-21',6),('2021-05-22',7),('2021-05-23',1),('2021-05-24',2),('2021-05-25',3),('2021-05-26',4),('2021-05-27',5),('2021-05-28',6),('2021-05-29',7),('2021-05-30',1),('2021-05-31',2),('2021-06-01',3),('2021-06-02',4),('2021-06-03',5),('2021-06-04',6),('2021-06-05',7),('2021-06-06',1),('2021-06-07',2),('2021-06-08',3),('2021-06-09',4),('2021-06-10',5),('2021-06-11',6),('2021-06-12',7),('2021-06-13',1),('2021-06-14',2),('2021-06-15',3),('2021-06-16',4),('2021-06-17',5),('2021-06-18',6),('2021-06-19',7),('2021-06-20',1),('2021-06-21',2),('2021-06-22',3),('2021-06-23',4),('2021-06-24',5),('2021-06-25',6),('2021-06-26',7),('2021-06-27',1),('2021-06-28',2),('2021-06-29',3),('2021-06-30',4),('2021-07-01',5),('2021-07-02',6),('2021-07-03',7),('2021-07-04',1),('2021-07-05',2),('2021-07-06',3),('2021-07-07',4),('2021-07-08',5),('2021-07-09',6),('2021-07-10',7),('2021-07-11',1),('2021-07-12',2),('2021-07-13',3),('2021-07-14',4),('2021-07-15',5),('2021-07-16',6),('2021-07-17',7),('2021-07-18',1),('2021-07-19',2),('2021-07-20',3),('2021-07-21',4),('2021-07-22',5),('2021-07-23',6),('2021-07-24',7),('2021-07-25',1),('2021-07-26',2),('2021-07-27',3),('2021-07-28',4),('2021-07-29',5),('2021-07-30',6),('2021-07-31',7),('2021-08-01',1),('2021-08-02',2),('2021-08-03',3),('2021-08-04',4),('2021-08-05',5),('2021-08-06',6),('2021-08-07',7),('2021-08-08',1),('2021-08-09',2),('2021-08-10',3),('2021-08-11',4),('2021-08-12',5),('2021-08-13',6),('2021-08-14',7),('2021-08-15',1),('2021-08-16',2),('2021-08-17',3),('2021-08-18',4),('2021-08-19',5),('2021-08-20',6),('2021-08-21',7),('2021-08-22',1),('2021-08-23',2),('2021-08-24',3),('2021-08-25',4),('2021-08-26',5),('2021-08-27',6),('2021-08-28',7),('2021-08-29',1),('2021-08-30',2),('2021-08-31',3),('2021-09-01',4),('2021-09-02',5),('2021-09-03',6),('2021-09-04',7),('2021-09-05',1),('2021-09-06',2),('2021-09-07',3),('2021-09-08',4),('2021-09-09',5),('2021-09-10',6),('2021-09-11',7),('2021-09-12',1),('2021-09-13',2),('2021-09-14',3),('2021-09-15',4),('2021-09-16',5),('2021-09-17',6),('2021-09-18',7),('2021-09-19',1),('2021-09-20',2),('2021-09-21',3),('2021-09-22',4),('2021-09-23',5),('2021-09-24',6),('2021-09-25',7),('2021-09-26',1),('2021-09-27',2),('2021-09-28',3),('2021-09-29',4),('2021-09-30',5),('2021-10-01',6),('2021-10-02',7),('2021-10-03',1),('2021-10-04',2),('2021-10-05',3),('2021-10-06',4),('2021-10-07',5),('2021-10-08',6),('2021-10-09',7),('2021-10-10',1),('2021-10-11',2),('2021-10-12',3),('2021-10-13',4),('2021-10-14',5),('2021-10-15',6),('2021-10-16',7),('2021-10-17',1),('2021-10-18',2),('2021-10-19',3),('2021-10-20',4),('2021-10-21',5),('2021-10-22',6),('2021-10-23',7),('2021-10-24',1),('2021-10-25',2),('2021-10-26',3),('2021-10-27',4),('2021-10-28',5),('2021-10-29',6),('2021-10-30',7),('2021-10-31',1),('2021-11-01',2),('2021-11-02',3),('2021-11-03',4),('2021-11-04',5),('2021-11-05',6),('2021-11-06',7),('2021-11-07',1),('2021-11-08',2),('2021-11-09',3),('2021-11-10',4),('2021-11-11',5),('2021-11-12',6),('2021-11-13',7),('2021-11-14',1),('2021-11-15',2),('2021-11-16',3),('2021-11-17',4),('2021-11-18',5),('2021-11-19',6),('2021-11-20',7),('2021-11-21',1),('2021-11-22',2),('2021-11-23',3),('2021-11-24',4),('2021-11-25',5),('2021-11-26',6),('2021-11-27',7),('2021-11-28',1),('2021-11-29',2),('2021-11-30',3),('2021-12-01',4),('2021-12-02',5),('2021-12-03',6),('2021-12-04',7),('2021-12-05',1),('2021-12-06',2),('2021-12-07',3),('2021-12-08',4),('2021-12-09',5),('2021-12-10',6),('2021-12-11',7),('2021-12-12',1),('2021-12-13',2),('2021-12-14',3),('2021-12-15',4),('2021-12-16',5),('2021-12-17',6),('2021-12-18',7),('2021-12-19',1),('2021-12-20',2),('2021-12-21',3),('2021-12-22',4),('2021-12-23',5),('2021-12-24',6),('2021-12-25',7),('2021-12-26',1),('2021-12-27',2),('2021-12-28',3),('2021-12-29',4),('2021-12-30',5),('2021-12-31',6),('2022-01-01',7),('2022-01-02',1),('2022-01-03',2),('2022-01-04',3),('2022-01-05',4),('2022-01-06',5),('2022-01-07',6),('2022-01-08',7),('2022-01-09',1),('2022-01-10',2),('2022-01-11',3),('2022-01-12',4),('2022-01-13',5),('2022-01-14',6),('2022-01-15',7),('2022-01-16',1),('2022-01-17',2),('2022-01-18',3),('2022-01-19',4),('2022-01-20',5),('2022-01-21',6),('2022-01-22',7),('2022-01-23',1),('2022-01-24',2),('2022-01-25',3),('2022-01-26',4),('2022-01-27',5),('2022-01-28',6),('2022-01-29',7),('2022-01-30',1),('2022-01-31',2),('2022-02-01',3),('2022-02-02',4),('2022-02-03',5),('2022-02-04',6),('2022-02-05',7),('2022-02-06',1),('2022-02-07',2),('2022-02-08',3),('2022-02-09',4),('2022-02-10',5),('2022-02-11',6),('2022-02-12',7),('2022-02-13',1),('2022-02-14',2),('2022-02-15',3),('2022-02-16',4),('2022-02-17',5),('2022-02-18',6),('2022-02-19',7),('2022-02-20',1),('2022-02-21',2),('2022-02-22',3),('2022-02-23',4),('2022-02-24',5),('2022-02-25',6),('2022-02-26',7),('2022-02-27',1),('2022-02-28',2),('2022-03-01',3),('2022-03-02',4),('2022-03-03',5),('2022-03-04',6),('2022-03-05',7),('2022-03-06',1),('2022-03-07',2),('2022-03-08',3),('2022-03-09',4),('2022-03-10',5),('2022-03-11',6),('2022-03-12',7),('2022-03-13',1),('2022-03-14',2),('2022-03-15',3),('2022-03-16',4),('2022-03-17',5),('2022-03-18',6),('2022-03-19',7),('2022-03-20',1),('2022-03-21',2),('2022-03-22',3),('2022-03-23',4),('2022-03-24',5),('2022-03-25',6),('2022-03-26',7),('2022-03-27',1),('2022-03-28',2),('2022-03-29',3),('2022-03-30',4),('2022-03-31',5),('2022-04-01',6),('2022-04-02',7),('2022-04-03',1),('2022-04-04',2),('2022-04-05',3),('2022-04-06',4),('2022-04-07',5),('2022-04-08',6),('2022-04-09',7),('2022-04-10',1),('2022-04-11',2),('2022-04-12',3),('2022-04-13',4),('2022-04-14',5),('2022-04-15',6),('2022-04-16',7),('2022-04-17',1),('2022-04-18',2),('2022-04-19',3),('2022-04-20',4),('2022-04-21',5),('2022-04-22',6),('2022-04-23',7),('2022-04-24',1),('2022-04-25',2),('2022-04-26',3),('2022-04-27',4),('2022-04-28',5),('2022-04-29',6),('2022-04-30',7),('2022-05-01',1),('2022-05-02',2),('2022-05-03',3),('2022-05-04',4),('2022-05-05',5),('2022-05-06',6),('2022-05-07',7),('2022-05-08',1),('2022-05-09',2),('2022-05-10',3),('2022-05-11',4),('2022-05-12',5),('2022-05-13',6),('2022-05-14',7),('2022-05-15',1),('2022-05-16',2),('2022-05-17',3),('2022-05-18',4),('2022-05-19',5),('2022-05-20',6),('2022-05-21',7),('2022-05-22',1),('2022-05-23',2),('2022-05-24',3),('2022-05-25',4),('2022-05-26',5),('2022-05-27',6),('2022-05-28',7),('2022-05-29',1),('2022-05-30',2),('2022-05-31',3),('2022-06-01',4),('2022-06-02',5),('2022-06-03',6),('2022-06-04',7),('2022-06-05',1),('2022-06-06',2),('2022-06-07',3),('2022-06-08',4),('2022-06-09',5),('2022-06-10',6),('2022-06-11',7),('2022-06-12',1),('2022-06-13',2),('2022-06-14',3),('2022-06-15',4),('2022-06-16',5),('2022-06-17',6),('2022-06-18',7),('2022-06-19',1),('2022-06-20',2),('2022-06-21',3),('2022-06-22',4),('2022-06-23',5),('2022-06-24',6),('2022-06-25',7),('2022-06-26',1),('2022-06-27',2),('2022-06-28',3),('2022-06-29',4),('2022-06-30',5),('2022-07-01',6),('2022-07-02',7),('2022-07-03',1),('2022-07-04',2),('2022-07-05',3),('2022-07-06',4),('2022-07-07',5),('2022-07-08',6),('2022-07-09',7),('2022-07-10',1),('2022-07-11',2),('2022-07-12',3),('2022-07-13',4),('2022-07-14',5),('2022-07-15',6),('2022-07-16',7),('2022-07-17',1),('2022-07-18',2),('2022-07-19',3),('2022-07-20',4),('2022-07-21',5),('2022-07-22',6),('2022-07-23',7),('2022-07-24',1),('2022-07-25',2),('2022-07-26',3),('2022-07-27',4),('2022-07-28',5),('2022-07-29',6),('2022-07-30',7),('2022-07-31',1),('2022-08-01',2),('2022-08-02',3),('2022-08-03',4),('2022-08-04',5),('2022-08-05',6),('2022-08-06',7),('2022-08-07',1),('2022-08-08',2),('2022-08-09',3),('2022-08-10',4),('2022-08-11',5),('2022-08-12',6),('2022-08-13',7),('2022-08-14',1),('2022-08-15',2),('2022-08-16',3),('2022-08-17',4),('2022-08-18',5),('2022-08-19',6),('2022-08-20',7),('2022-08-21',1),('2022-08-22',2),('2022-08-23',3),('2022-08-24',4),('2022-08-25',5),('2022-08-26',6),('2022-08-27',7),('2022-08-28',1),('2022-08-29',2),('2022-08-30',3),('2022-08-31',4),('2022-09-01',5),('2022-09-02',6),('2022-09-03',7),('2022-09-04',1),('2022-09-05',2),('2022-09-06',3),('2022-09-07',4),('2022-09-08',5),('2022-09-09',6),('2022-09-10',7),('2022-09-11',1),('2022-09-12',2),('2022-09-13',3),('2022-09-14',4),('2022-09-15',5),('2022-09-16',6),('2022-09-17',7),('2022-09-18',1),('2022-09-19',2),('2022-09-20',3),('2022-09-21',4),('2022-09-22',5),('2022-09-23',6),('2022-09-24',7),('2022-09-25',1),('2022-09-26',2),('2022-09-27',3),('2022-09-28',4),('2022-09-29',5),('2022-09-30',6),('2022-10-01',7),('2022-10-02',1),('2022-10-03',2),('2022-10-04',3),('2022-10-05',4),('2022-10-06',5),('2022-10-07',6),('2022-10-08',7),('2022-10-09',1),('2022-10-10',2),('2022-10-11',3),('2022-10-12',4),('2022-10-13',5),('2022-10-14',6),('2022-10-15',7),('2022-10-16',1),('2022-10-17',2),('2022-10-18',3),('2022-10-19',4),('2022-10-20',5),('2022-10-21',6),('2022-10-22',7),('2022-10-23',1),('2022-10-24',2),('2022-10-25',3),('2022-10-26',4),('2022-10-27',5),('2022-10-28',6),('2022-10-29',7),('2022-10-30',1),('2022-10-31',2),('2022-11-01',3),('2022-11-02',4),('2022-11-03',5),('2022-11-04',6),('2022-11-05',7),('2022-11-06',1),('2022-11-07',2),('2022-11-08',3),('2022-11-09',4),('2022-11-10',5),('2022-11-11',6),('2022-11-12',7),('2022-11-13',1),('2022-11-14',2),('2022-11-15',3),('2022-11-16',4),('2022-11-17',5),('2022-11-18',6),('2022-11-19',7),('2022-11-20',1),('2022-11-21',2),('2022-11-22',3),('2022-11-23',4),('2022-11-24',5),('2022-11-25',6),('2022-11-26',7),('2022-11-27',1),('2022-11-28',2),('2022-11-29',3),('2022-11-30',4),('2022-12-01',5),('2022-12-02',6),('2022-12-03',7),('2022-12-04',1),('2022-12-05',2),('2022-12-06',3),('2022-12-07',4),('2022-12-08',5),('2022-12-09',6),('2022-12-10',7),('2022-12-11',1),('2022-12-12',2),('2022-12-13',3),('2022-12-14',4),('2022-12-15',5),('2022-12-16',6),('2022-12-17',7),('2022-12-18',1),('2022-12-19',2),('2022-12-20',3),('2022-12-21',4),('2022-12-22',5),('2022-12-23',6),('2022-12-24',7),('2022-12-25',1),('2022-12-26',2),('2022-12-27',3),('2022-12-28',4),('2022-12-29',5),('2022-12-30',6),('2022-12-31',7),('2023-01-01',1),('2023-01-02',2),('2023-01-03',3),('2023-01-04',4),('2023-01-05',5),('2023-01-06',6),('2023-01-07',7),('2023-01-08',1),('2023-01-09',2),('2023-01-10',3),('2023-01-11',4),('2023-01-12',5),('2023-01-13',6),('2023-01-14',7),('2023-01-15',1),('2023-01-16',2),('2023-01-17',3),('2023-01-18',4),('2023-01-19',5),('2023-01-20',6),('2023-01-21',7),('2023-01-22',1),('2023-01-23',2),('2023-01-24',3),('2023-01-25',4),('2023-01-26',5),('2023-01-27',6),('2023-01-28',7),('2023-01-29',1),('2023-01-30',2),('2023-01-31',3),('2023-02-01',4),('2023-02-02',5),('2023-02-03',6),('2023-02-04',7),('2023-02-05',1),('2023-02-06',2),('2023-02-07',3),('2023-02-08',4),('2023-02-09',5),('2023-02-10',6),('2023-02-11',7),('2023-02-12',1),('2023-02-13',2),('2023-02-14',3),('2023-02-15',4),('2023-02-16',5),('2023-02-17',6),('2023-02-18',7),('2023-02-19',1),('2023-02-20',2),('2023-02-21',3),('2023-02-22',4),('2023-02-23',5),('2023-02-24',6),('2023-02-25',7),('2023-02-26',1),('2023-02-27',2),('2023-02-28',3),('2023-03-01',4),('2023-03-02',5),('2023-03-03',6),('2023-03-04',7),('2023-03-05',1),('2023-03-06',2),('2023-03-07',3),('2023-03-08',4),('2023-03-09',5),('2023-03-10',6),('2023-03-11',7),('2023-03-12',1),('2023-03-13',2),('2023-03-14',3),('2023-03-15',4),('2023-03-16',5),('2023-03-17',6),('2023-03-18',7),('2023-03-19',1),('2023-03-20',2),('2023-03-21',3),('2023-03-22',4),('2023-03-23',5),('2023-03-24',6),('2023-03-25',7),('2023-03-26',1),('2023-03-27',2),('2023-03-28',3),('2023-03-29',4),('2023-03-30',5),('2023-03-31',6),('2023-04-01',7),('2023-04-02',1),('2023-04-03',2),('2023-04-04',3),('2023-04-05',4),('2023-04-06',5),('2023-04-07',6),('2023-04-08',7),('2023-04-09',1),('2023-04-10',2),('2023-04-11',3),('2023-04-12',4),('2023-04-13',5),('2023-04-14',6),('2023-04-15',7),('2023-04-16',1),('2023-04-17',2),('2023-04-18',3),('2023-04-19',4),('2023-04-20',5),('2023-04-21',6),('2023-04-22',7),('2023-04-23',1),('2023-04-24',2),('2023-04-25',3),('2023-04-26',4),('2023-04-27',5),('2023-04-28',6),('2023-04-29',7),('2023-04-30',1),('2023-05-01',2),('2023-05-02',3),('2023-05-03',4),('2023-05-04',5),('2023-05-05',6),('2023-05-06',7),('2023-05-07',1),('2023-05-08',2),('2023-05-09',3),('2023-05-10',4),('2023-05-11',5),('2023-05-12',6),('2023-05-13',7),('2023-05-14',1),('2023-05-15',2),('2023-05-16',3),('2023-05-17',4),('2023-05-18',5),('2023-05-19',6),('2023-05-20',7),('2023-05-21',1),('2023-05-22',2),('2023-05-23',3),('2023-05-24',4),('2023-05-25',5),('2023-05-26',6),('2023-05-27',7),('2023-05-28',1),('2023-05-29',2),('2023-05-30',3),('2023-05-31',4),('2023-06-01',5),('2023-06-02',6),('2023-06-03',7),('2023-06-04',1),('2023-06-05',2),('2023-06-06',3),('2023-06-07',4),('2023-06-08',5),('2023-06-09',6),('2023-06-10',7),('2023-06-11',1),('2023-06-12',2),('2023-06-13',3),('2023-06-14',4),('2023-06-15',5),('2023-06-16',6),('2023-06-17',7),('2023-06-18',1),('2023-06-19',2),('2023-06-20',3),('2023-06-21',4),('2023-06-22',5),('2023-06-23',6),('2023-06-24',7),('2023-06-25',1),('2023-06-26',2),('2023-06-27',3),('2023-06-28',4),('2023-06-29',5),('2023-06-30',6),('2023-07-01',7),('2023-07-02',1),('2023-07-03',2),('2023-07-04',3),('2023-07-05',4),('2023-07-06',5),('2023-07-07',6),('2023-07-08',7),('2023-07-09',1),('2023-07-10',2),('2023-07-11',3),('2023-07-12',4),('2023-07-13',5),('2023-07-14',6),('2023-07-15',7),('2023-07-16',1),('2023-07-17',2),('2023-07-18',3),('2023-07-19',4),('2023-07-20',5),('2023-07-21',6),('2023-07-22',7),('2023-07-23',1),('2023-07-24',2),('2023-07-25',3),('2023-07-26',4),('2023-07-27',5),('2023-07-28',6),('2023-07-29',7),('2023-07-30',1),('2023-07-31',2),('2023-08-01',3),('2023-08-02',4),('2023-08-03',5),('2023-08-04',6),('2023-08-05',7),('2023-08-06',1),('2023-08-07',2),('2023-08-08',3),('2023-08-09',4),('2023-08-10',5),('2023-08-11',6),('2023-08-12',7),('2023-08-13',1),('2023-08-14',2),('2023-08-15',3),('2023-08-16',4),('2023-08-17',5),('2023-08-18',6),('2023-08-19',7),('2023-08-20',1),('2023-08-21',2),('2023-08-22',3),('2023-08-23',4),('2023-08-24',5),('2023-08-25',6),('2023-08-26',7),('2023-08-27',1),('2023-08-28',2),('2023-08-29',3),('2023-08-30',4),('2023-08-31',5),('2023-09-01',6),('2023-09-02',7),('2023-09-03',1),('2023-09-04',2),('2023-09-05',3),('2023-09-06',4),('2023-09-07',5),('2023-09-08',6),('2023-09-09',7),('2023-09-10',1),('2023-09-11',2),('2023-09-12',3),('2023-09-13',4),('2023-09-14',5),('2023-09-15',6),('2023-09-16',7),('2023-09-17',1),('2023-09-18',2),('2023-09-19',3),('2023-09-20',4),('2023-09-21',5),('2023-09-22',6),('2023-09-23',7),('2023-09-24',1),('2023-09-25',2),('2023-09-26',3),('2023-09-27',4),('2023-09-28',5),('2023-09-29',6),('2023-09-30',7),('2023-10-01',1),('2023-10-02',2),('2023-10-03',3),('2023-10-04',4),('2023-10-05',5),('2023-10-06',6),('2023-10-07',7),('2023-10-08',1),('2023-10-09',2),('2023-10-10',3),('2023-10-11',4),('2023-10-12',5),('2023-10-13',6),('2023-10-14',7),('2023-10-15',1),('2023-10-16',2),('2023-10-17',3),('2023-10-18',4),('2023-10-19',5),('2023-10-20',6),('2023-10-21',7),('2023-10-22',1),('2023-10-23',2),('2023-10-24',3),('2023-10-25',4),('2023-10-26',5),('2023-10-27',6),('2023-10-28',7),('2023-10-29',1),('2023-10-30',2),('2023-10-31',3),('2023-11-01',4),('2023-11-02',5),('2023-11-03',6),('2023-11-04',7),('2023-11-05',1),('2023-11-06',2),('2023-11-07',3),('2023-11-08',4),('2023-11-09',5),('2023-11-10',6),('2023-11-11',7),('2023-11-12',1),('2023-11-13',2),('2023-11-14',3),('2023-11-15',4),('2023-11-16',5),('2023-11-17',6),('2023-11-18',7),('2023-11-19',1),('2023-11-20',2),('2023-11-21',3),('2023-11-22',4),('2023-11-23',5),('2023-11-24',6),('2023-11-25',7),('2023-11-26',1),('2023-11-27',2),('2023-11-28',3),('2023-11-29',4),('2023-11-30',5),('2023-12-01',6),('2023-12-02',7),('2023-12-03',1),('2023-12-04',2),('2023-12-05',3),('2023-12-06',4),('2023-12-07',5),('2023-12-08',6),('2023-12-09',7),('2023-12-10',1),('2023-12-11',2),('2023-12-12',3),('2023-12-13',4),('2023-12-14',5),('2023-12-15',6),('2023-12-16',7),('2023-12-17',1),('2023-12-18',2),('2023-12-19',3),('2023-12-20',4),('2023-12-21',5),('2023-12-22',6),('2023-12-23',7),('2023-12-24',1),('2023-12-25',2),('2023-12-26',3),('2023-12-27',4),('2023-12-28',5),('2023-12-29',6),('2023-12-30',7),('2023-12-31',1),('2024-01-01',2),('2024-01-02',3),('2024-01-03',4),('2024-01-04',5),('2024-01-05',6),('2024-01-06',7),('2024-01-07',1),('2024-01-08',2),('2024-01-09',3),('2024-01-10',4),('2024-01-11',5),('2024-01-12',6),('2024-01-13',7),('2024-01-14',1),('2024-01-15',2),('2024-01-16',3),('2024-01-17',4),('2024-01-18',5),('2024-01-19',6),('2024-01-20',7),('2024-01-21',1),('2024-01-22',2),('2024-01-23',3),('2024-01-24',4),('2024-01-25',5),('2024-01-26',6),('2024-01-27',7),('2024-01-28',1),('2024-01-29',2),('2024-01-30',3),('2024-01-31',4),('2024-02-01',5),('2024-02-02',6),('2024-02-03',7),('2024-02-04',1),('2024-02-05',2),('2024-02-06',3),('2024-02-07',4),('2024-02-08',5),('2024-02-09',6),('2024-02-10',7),('2024-02-11',1),('2024-02-12',2),('2024-02-13',3),('2024-02-14',4),('2024-02-15',5),('2024-02-16',6),('2024-02-17',7),('2024-02-18',1),('2024-02-19',2),('2024-02-20',3),('2024-02-21',4),('2024-02-22',5),('2024-02-23',6),('2024-02-24',7),('2024-02-25',1),('2024-02-26',2),('2024-02-27',3),('2024-02-28',4),('2024-02-29',5),('2024-03-01',6),('2024-03-02',7),('2024-03-03',1),('2024-03-04',2),('2024-03-05',3),('2024-03-06',4),('2024-03-07',5),('2024-03-08',6),('2024-03-09',7),('2024-03-10',1),('2024-03-11',2),('2024-03-12',3),('2024-03-13',4),('2024-03-14',5),('2024-03-15',6),('2024-03-16',7),('2024-03-17',1),('2024-03-18',2),('2024-03-19',3),('2024-03-20',4),('2024-03-21',5),('2024-03-22',6),('2024-03-23',7),('2024-03-24',1),('2024-03-25',2),('2024-03-26',3),('2024-03-27',4),('2024-03-28',5),('2024-03-29',6),('2024-03-30',7),('2024-03-31',1),('2024-04-01',2),('2024-04-02',3),('2024-04-03',4),('2024-04-04',5),('2024-04-05',6),('2024-04-06',7),('2024-04-07',1),('2024-04-08',2),('2024-04-09',3),('2024-04-10',4),('2024-04-11',5),('2024-04-12',6),('2024-04-13',7),('2024-04-14',1),('2024-04-15',2),('2024-04-16',3),('2024-04-17',4),('2024-04-18',5),('2024-04-19',6),('2024-04-20',7),('2024-04-21',1),('2024-04-22',2),('2024-04-23',3),('2024-04-24',4),('2024-04-25',5),('2024-04-26',6),('2024-04-27',7),('2024-04-28',1),('2024-04-29',2),('2024-04-30',3),('2024-05-01',4),('2024-05-02',5),('2024-05-03',6),('2024-05-04',7),('2024-05-05',1),('2024-05-06',2),('2024-05-07',3),('2024-05-08',4),('2024-05-09',5),('2024-05-10',6),('2024-05-11',7),('2024-05-12',1),('2024-05-13',2),('2024-05-14',3),('2024-05-15',4),('2024-05-16',5),('2024-05-17',6),('2024-05-18',7),('2024-05-19',1),('2024-05-20',2),('2024-05-21',3),('2024-05-22',4),('2024-05-23',5),('2024-05-24',6),('2024-05-25',7),('2024-05-26',1),('2024-05-27',2),('2024-05-28',3),('2024-05-29',4),('2024-05-30',5),('2024-05-31',6),('2024-06-01',7),('2024-06-02',1),('2024-06-03',2),('2024-06-04',3),('2024-06-05',4),('2024-06-06',5),('2024-06-07',6),('2024-06-08',7),('2024-06-09',1),('2024-06-10',2),('2024-06-11',3),('2024-06-12',4),('2024-06-13',5),('2024-06-14',6),('2024-06-15',7),('2024-06-16',1),('2024-06-17',2),('2024-06-18',3),('2024-06-19',4),('2024-06-20',5),('2024-06-21',6),('2024-06-22',7),('2024-06-23',1),('2024-06-24',2),('2024-06-25',3),('2024-06-26',4),('2024-06-27',5),('2024-06-28',6),('2024-06-29',7),('2024-06-30',1),('2024-07-01',2),('2024-07-02',3),('2024-07-03',4),('2024-07-04',5),('2024-07-05',6),('2024-07-06',7),('2024-07-07',1),('2024-07-08',2),('2024-07-09',3),('2024-07-10',4),('2024-07-11',5),('2024-07-12',6),('2024-07-13',7),('2024-07-14',1),('2024-07-15',2),('2024-07-16',3),('2024-07-17',4),('2024-07-18',5),('2024-07-19',6),('2024-07-20',7),('2024-07-21',1),('2024-07-22',2),('2024-07-23',3),('2024-07-24',4),('2024-07-25',5),('2024-07-26',6),('2024-07-27',7),('2024-07-28',1),('2024-07-29',2),('2024-07-30',3),('2024-07-31',4),('2024-08-01',5),('2024-08-02',6),('2024-08-03',7),('2024-08-04',1),('2024-08-05',2),('2024-08-06',3),('2024-08-07',4),('2024-08-08',5),('2024-08-09',6),('2024-08-10',7),('2024-08-11',1),('2024-08-12',2),('2024-08-13',3),('2024-08-14',4),('2024-08-15',5),('2024-08-16',6),('2024-08-17',7),('2024-08-18',1),('2024-08-19',2),('2024-08-20',3),('2024-08-21',4),('2024-08-22',5),('2024-08-23',6),('2024-08-24',7),('2024-08-25',1),('2024-08-26',2),('2024-08-27',3),('2024-08-28',4),('2024-08-29',5),('2024-08-30',6),('2024-08-31',7),('2024-09-01',1),('2024-09-02',2),('2024-09-03',3),('2024-09-04',4),('2024-09-05',5),('2024-09-06',6),('2024-09-07',7),('2024-09-08',1),('2024-09-09',2),('2024-09-10',3),('2024-09-11',4),('2024-09-12',5),('2024-09-13',6),('2024-09-14',7),('2024-09-15',1),('2024-09-16',2),('2024-09-17',3),('2024-09-18',4),('2024-09-19',5),('2024-09-20',6),('2024-09-21',7),('2024-09-22',1),('2024-09-23',2),('2024-09-24',3),('2024-09-25',4),('2024-09-26',5),('2024-09-27',6),('2024-09-28',7),('2024-09-29',1),('2024-09-30',2),('2024-10-01',3),('2024-10-02',4),('2024-10-03',5),('2024-10-04',6),('2024-10-05',7),('2024-10-06',1),('2024-10-07',2),('2024-10-08',3),('2024-10-09',4),('2024-10-10',5),('2024-10-11',6),('2024-10-12',7),('2024-10-13',1),('2024-10-14',2),('2024-10-15',3),('2024-10-16',4),('2024-10-17',5),('2024-10-18',6),('2024-10-19',7),('2024-10-20',1),('2024-10-21',2),('2024-10-22',3),('2024-10-23',4),('2024-10-24',5),('2024-10-25',6),('2024-10-26',7),('2024-10-27',1),('2024-10-28',2),('2024-10-29',3),('2024-10-30',4),('2024-10-31',5),('2024-11-01',6),('2024-11-02',7),('2024-11-03',1),('2024-11-04',2),('2024-11-05',3),('2024-11-06',4),('2024-11-07',5),('2024-11-08',6),('2024-11-09',7),('2024-11-10',1),('2024-11-11',2),('2024-11-12',3),('2024-11-13',4),('2024-11-14',5),('2024-11-15',6),('2024-11-16',7),('2024-11-17',1),('2024-11-18',2),('2024-11-19',3),('2024-11-20',4),('2024-11-21',5),('2024-11-22',6),('2024-11-23',7),('2024-11-24',1),('2024-11-25',2),('2024-11-26',3),('2024-11-27',4),('2024-11-28',5),('2024-11-29',6),('2024-11-30',7),('2024-12-01',1),('2024-12-02',2),('2024-12-03',3),('2024-12-04',4),('2024-12-05',5),('2024-12-06',6),('2024-12-07',7),('2024-12-08',1),('2024-12-09',2),('2024-12-10',3),('2024-12-11',4),('2024-12-12',5),('2024-12-13',6),('2024-12-14',7),('2024-12-15',1),('2024-12-16',2),('2024-12-17',3),('2024-12-18',4),('2024-12-19',5),('2024-12-20',6),('2024-12-21',7),('2024-12-22',1),('2024-12-23',2),('2024-12-24',3),('2024-12-25',4),('2024-12-26',5),('2024-12-27',6),('2024-12-28',7),('2024-12-29',1),('2024-12-30',2),('2024-12-31',3),('2025-01-01',4),('2025-01-02',5),('2025-01-03',6),('2025-01-04',7),('2025-01-05',1),('2025-01-06',2),('2025-01-07',3),('2025-01-08',4),('2025-01-09',5),('2025-01-10',6),('2025-01-11',7),('2025-01-12',1),('2025-01-13',2),('2025-01-14',3),('2025-01-15',4),('2025-01-16',5),('2025-01-17',6),('2025-01-18',7),('2025-01-19',1),('2025-01-20',2),('2025-01-21',3),('2025-01-22',4),('2025-01-23',5),('2025-01-24',6),('2025-01-25',7),('2025-01-26',1),('2025-01-27',2),('2025-01-28',3),('2025-01-29',4),('2025-01-30',5),('2025-01-31',6),('2025-02-01',7),('2025-02-02',1),('2025-02-03',2),('2025-02-04',3),('2025-02-05',4),('2025-02-06',5),('2025-02-07',6),('2025-02-08',7),('2025-02-09',1),('2025-02-10',2),('2025-02-11',3),('2025-02-12',4),('2025-02-13',5),('2025-02-14',6),('2025-02-15',7),('2025-02-16',1),('2025-02-17',2),('2025-02-18',3),('2025-02-19',4),('2025-02-20',5),('2025-02-21',6),('2025-02-22',7),('2025-02-23',1),('2025-02-24',2),('2025-02-25',3),('2025-02-26',4),('2025-02-27',5),('2025-02-28',6),('2025-03-01',7),('2025-03-02',1),('2025-03-03',2),('2025-03-04',3),('2025-03-05',4),('2025-03-06',5),('2025-03-07',6),('2025-03-08',7),('2025-03-09',1),('2025-03-10',2),('2025-03-11',3),('2025-03-12',4),('2025-03-13',5),('2025-03-14',6),('2025-03-15',7),('2025-03-16',1),('2025-03-17',2),('2025-03-18',3),('2025-03-19',4),('2025-03-20',5),('2025-03-21',6),('2025-03-22',7),('2025-03-23',1),('2025-03-24',2),('2025-03-25',3),('2025-03-26',4),('2025-03-27',5),('2025-03-28',6),('2025-03-29',7),('2025-03-30',1),('2025-03-31',2),('2025-04-01',3),('2025-04-02',4),('2025-04-03',5),('2025-04-04',6),('2025-04-05',7),('2025-04-06',1),('2025-04-07',2),('2025-04-08',3),('2025-04-09',4),('2025-04-10',5),('2025-04-11',6),('2025-04-12',7),('2025-04-13',1),('2025-04-14',2),('2025-04-15',3),('2025-04-16',4),('2025-04-17',5),('2025-04-18',6),('2025-04-19',7),('2025-04-20',1),('2025-04-21',2),('2025-04-22',3),('2025-04-23',4),('2025-04-24',5),('2025-04-25',6),('2025-04-26',7),('2025-04-27',1),('2025-04-28',2),('2025-04-29',3),('2025-04-30',4),('2025-05-01',5),('2025-05-02',6),('2025-05-03',7),('2025-05-04',1),('2025-05-05',2),('2025-05-06',3),('2025-05-07',4),('2025-05-08',5),('2025-05-09',6),('2025-05-10',7),('2025-05-11',1),('2025-05-12',2),('2025-05-13',3),('2025-05-14',4),('2025-05-15',5),('2025-05-16',6),('2025-05-17',7),('2025-05-18',1),('2025-05-19',2),('2025-05-20',3),('2025-05-21',4),('2025-05-22',5),('2025-05-23',6),('2025-05-24',7),('2025-05-25',1),('2025-05-26',2),('2025-05-27',3),('2025-05-28',4),('2025-05-29',5),('2025-05-30',6),('2025-05-31',7),('2025-06-01',1),('2025-06-02',2),('2025-06-03',3),('2025-06-04',4),('2025-06-05',5),('2025-06-06',6),('2025-06-07',7),('2025-06-08',1),('2025-06-09',2),('2025-06-10',3),('2025-06-11',4),('2025-06-12',5),('2025-06-13',6),('2025-06-14',7),('2025-06-15',1),('2025-06-16',2),('2025-06-17',3),('2025-06-18',4),('2025-06-19',5),('2025-06-20',6),('2025-06-21',7),('2025-06-22',1),('2025-06-23',2),('2025-06-24',3),('2025-06-25',4),('2025-06-26',5),('2025-06-27',6),('2025-06-28',7),('2025-06-29',1),('2025-06-30',2),('2025-07-01',3),('2025-07-02',4),('2025-07-03',5),('2025-07-04',6),('2025-07-05',7),('2025-07-06',1),('2025-07-07',2),('2025-07-08',3),('2025-07-09',4),('2025-07-10',5),('2025-07-11',6),('2025-07-12',7),('2025-07-13',1),('2025-07-14',2),('2025-07-15',3),('2025-07-16',4),('2025-07-17',5),('2025-07-18',6),('2025-07-19',7),('2025-07-20',1),('2025-07-21',2),('2025-07-22',3),('2025-07-23',4),('2025-07-24',5),('2025-07-25',6),('2025-07-26',7),('2025-07-27',1),('2025-07-28',2),('2025-07-29',3),('2025-07-30',4),('2025-07-31',5),('2025-08-01',6),('2025-08-02',7),('2025-08-03',1),('2025-08-04',2),('2025-08-05',3),('2025-08-06',4),('2025-08-07',5),('2025-08-08',6),('2025-08-09',7),('2025-08-10',1),('2025-08-11',2),('2025-08-12',3),('2025-08-13',4),('2025-08-14',5),('2025-08-15',6),('2025-08-16',7),('2025-08-17',1),('2025-08-18',2),('2025-08-19',3),('2025-08-20',4),('2025-08-21',5),('2025-08-22',6),('2025-08-23',7),('2025-08-24',1),('2025-08-25',2),('2025-08-26',3),('2025-08-27',4),('2025-08-28',5),('2025-08-29',6),('2025-08-30',7),('2025-08-31',1),('2025-09-01',2),('2025-09-02',3),('2025-09-03',4),('2025-09-04',5),('2025-09-05',6),('2025-09-06',7),('2025-09-07',1),('2025-09-08',2),('2025-09-09',3),('2025-09-10',4),('2025-09-11',5),('2025-09-12',6),('2025-09-13',7),('2025-09-14',1),('2025-09-15',2),('2025-09-16',3),('2025-09-17',4),('2025-09-18',5),('2025-09-19',6),('2025-09-20',7),('2025-09-21',1),('2025-09-22',2),('2025-09-23',3),('2025-09-24',4),('2025-09-25',5),('2025-09-26',6),('2025-09-27',7),('2025-09-28',1),('2025-09-29',2),('2025-09-30',3),('2025-10-01',4),('2025-10-02',5),('2025-10-03',6),('2025-10-04',7),('2025-10-05',1),('2025-10-06',2),('2025-10-07',3),('2025-10-08',4),('2025-10-09',5),('2025-10-10',6),('2025-10-11',7),('2025-10-12',1),('2025-10-13',2),('2025-10-14',3),('2025-10-15',4),('2025-10-16',5),('2025-10-17',6),('2025-10-18',7),('2025-10-19',1),('2025-10-20',2),('2025-10-21',3),('2025-10-22',4),('2025-10-23',5),('2025-10-24',6),('2025-10-25',7),('2025-10-26',1),('2025-10-27',2),('2025-10-28',3),('2025-10-29',4),('2025-10-30',5),('2025-10-31',6),('2025-11-01',7),('2025-11-02',1),('2025-11-03',2),('2025-11-04',3),('2025-11-05',4),('2025-11-06',5),('2025-11-07',6),('2025-11-08',7),('2025-11-09',1),('2025-11-10',2),('2025-11-11',3),('2025-11-12',4),('2025-11-13',5),('2025-11-14',6),('2025-11-15',7),('2025-11-16',1),('2025-11-17',2),('2025-11-18',3),('2025-11-19',4),('2025-11-20',5),('2025-11-21',6),('2025-11-22',7),('2025-11-23',1),('2025-11-24',2),('2025-11-25',3),('2025-11-26',4),('2025-11-27',5),('2025-11-28',6),('2025-11-29',7),('2025-11-30',1),('2025-12-01',2),('2025-12-02',3),('2025-12-03',4),('2025-12-04',5),('2025-12-05',6),('2025-12-06',7),('2025-12-07',1),('2025-12-08',2),('2025-12-09',3),('2025-12-10',4),('2025-12-11',5),('2025-12-12',6),('2025-12-13',7),('2025-12-14',1),('2025-12-15',2),('2025-12-16',3),('2025-12-17',4),('2025-12-18',5),('2025-12-19',6),('2025-12-20',7),('2025-12-21',1),('2025-12-22',2),('2025-12-23',3),('2025-12-24',4),('2025-12-25',5),('2025-12-26',6),('2025-12-27',7),('2025-12-28',1),('2025-12-29',2),('2025-12-30',3),('2025-12-31',4),('2026-01-01',5),('2026-01-02',6),('2026-01-03',7),('2026-01-04',1),('2026-01-05',2),('2026-01-06',3),('2026-01-07',4),('2026-01-08',5),('2026-01-09',6),('2026-01-10',7),('2026-01-11',1),('2026-01-12',2),('2026-01-13',3),('2026-01-14',4),('2026-01-15',5),('2026-01-16',6),('2026-01-17',7),('2026-01-18',1),('2026-01-19',2),('2026-01-20',3),('2026-01-21',4),('2026-01-22',5),('2026-01-23',6),('2026-01-24',7),('2026-01-25',1),('2026-01-26',2),('2026-01-27',3),('2026-01-28',4),('2026-01-29',5),('2026-01-30',6),('2026-01-31',7),('2026-02-01',1),('2026-02-02',2),('2026-02-03',3),('2026-02-04',4),('2026-02-05',5),('2026-02-06',6),('2026-02-07',7),('2026-02-08',1),('2026-02-09',2),('2026-02-10',3),('2026-02-11',4),('2026-02-12',5),('2026-02-13',6),('2026-02-14',7),('2026-02-15',1),('2026-02-16',2),('2026-02-17',3),('2026-02-18',4),('2026-02-19',5),('2026-02-20',6),('2026-02-21',7),('2026-02-22',1),('2026-02-23',2),('2026-02-24',3),('2026-02-25',4),('2026-02-26',5),('2026-02-27',6),('2026-02-28',7),('2026-03-01',1),('2026-03-02',2),('2026-03-03',3),('2026-03-04',4),('2026-03-05',5),('2026-03-06',6),('2026-03-07',7),('2026-03-08',1),('2026-03-09',2),('2026-03-10',3),('2026-03-11',4),('2026-03-12',5),('2026-03-13',6),('2026-03-14',7),('2026-03-15',1),('2026-03-16',2),('2026-03-17',3),('2026-03-18',4),('2026-03-19',5),('2026-03-20',6),('2026-03-21',7),('2026-03-22',1),('2026-03-23',2),('2026-03-24',3),('2026-03-25',4),('2026-03-26',5),('2026-03-27',6),('2026-03-28',7),('2026-03-29',1),('2026-03-30',2),('2026-03-31',3),('2026-04-01',4),('2026-04-02',5),('2026-04-03',6),('2026-04-04',7),('2026-04-05',1),('2026-04-06',2),('2026-04-07',3),('2026-04-08',4),('2026-04-09',5),('2026-04-10',6),('2026-04-11',7),('2026-04-12',1),('2026-04-13',2),('2026-04-14',3),('2026-04-15',4),('2026-04-16',5),('2026-04-17',6),('2026-04-18',7),('2026-04-19',1),('2026-04-20',2),('2026-04-21',3),('2026-04-22',4),('2026-04-23',5),('2026-04-24',6),('2026-04-25',7),('2026-04-26',1),('2026-04-27',2),('2026-04-28',3),('2026-04-29',4),('2026-04-30',5),('2026-05-01',6),('2026-05-02',7),('2026-05-03',1),('2026-05-04',2),('2026-05-05',3),('2026-05-06',4),('2026-05-07',5),('2026-05-08',6),('2026-05-09',7),('2026-05-10',1),('2026-05-11',2),('2026-05-12',3),('2026-05-13',4),('2026-05-14',5),('2026-05-15',6),('2026-05-16',7),('2026-05-17',1),('2026-05-18',2),('2026-05-19',3),('2026-05-20',4),('2026-05-21',5),('2026-05-22',6),('2026-05-23',7),('2026-05-24',1),('2026-05-25',2),('2026-05-26',3),('2026-05-27',4),('2026-05-28',5),('2026-05-29',6),('2026-05-30',7),('2026-05-31',1),('2026-06-01',2),('2026-06-02',3),('2026-06-03',4),('2026-06-04',5),('2026-06-05',6),('2026-06-06',7),('2026-06-07',1),('2026-06-08',2),('2026-06-09',3),('2026-06-10',4),('2026-06-11',5),('2026-06-12',6),('2026-06-13',7),('2026-06-14',1),('2026-06-15',2),('2026-06-16',3),('2026-06-17',4),('2026-06-18',5),('2026-06-19',6),('2026-06-20',7),('2026-06-21',1),('2026-06-22',2),('2026-06-23',3),('2026-06-24',4),('2026-06-25',5),('2026-06-26',6),('2026-06-27',7),('2026-06-28',1),('2026-06-29',2),('2026-06-30',3),('2026-07-01',4),('2026-07-02',5),('2026-07-03',6),('2026-07-04',7),('2026-07-05',1),('2026-07-06',2),('2026-07-07',3),('2026-07-08',4),('2026-07-09',5),('2026-07-10',6),('2026-07-11',7),('2026-07-12',1),('2026-07-13',2),('2026-07-14',3),('2026-07-15',4),('2026-07-16',5),('2026-07-17',6),('2026-07-18',7),('2026-07-19',1),('2026-07-20',2),('2026-07-21',3),('2026-07-22',4),('2026-07-23',5),('2026-07-24',6),('2026-07-25',7),('2026-07-26',1),('2026-07-27',2),('2026-07-28',3),('2026-07-29',4),('2026-07-30',5),('2026-07-31',6),('2026-08-01',7),('2026-08-02',1),('2026-08-03',2),('2026-08-04',3),('2026-08-05',4),('2026-08-06',5),('2026-08-07',6),('2026-08-08',7),('2026-08-09',1),('2026-08-10',2),('2026-08-11',3),('2026-08-12',4),('2026-08-13',5),('2026-08-14',6),('2026-08-15',7),('2026-08-16',1),('2026-08-17',2),('2026-08-18',3),('2026-08-19',4),('2026-08-20',5),('2026-08-21',6),('2026-08-22',7),('2026-08-23',1),('2026-08-24',2),('2026-08-25',3),('2026-08-26',4),('2026-08-27',5),('2026-08-28',6),('2026-08-29',7),('2026-08-30',1),('2026-08-31',2),('2026-09-01',3),('2026-09-02',4),('2026-09-03',5),('2026-09-04',6),('2026-09-05',7),('2026-09-06',1),('2026-09-07',2),('2026-09-08',3),('2026-09-09',4),('2026-09-10',5),('2026-09-11',6),('2026-09-12',7),('2026-09-13',1),('2026-09-14',2),('2026-09-15',3),('2026-09-16',4),('2026-09-17',5),('2026-09-18',6),('2026-09-19',7),('2026-09-20',1),('2026-09-21',2),('2026-09-22',3),('2026-09-23',4),('2026-09-24',5),('2026-09-25',6),('2026-09-26',7),('2026-09-27',1),('2026-09-28',2),('2026-09-29',3),('2026-09-30',4),('2026-10-01',5),('2026-10-02',6),('2026-10-03',7),('2026-10-04',1),('2026-10-05',2),('2026-10-06',3),('2026-10-07',4),('2026-10-08',5),('2026-10-09',6),('2026-10-10',7),('2026-10-11',1),('2026-10-12',2),('2026-10-13',3),('2026-10-14',4),('2026-10-15',5),('2026-10-16',6),('2026-10-17',7),('2026-10-18',1),('2026-10-19',2),('2026-10-20',3),('2026-10-21',4),('2026-10-22',5),('2026-10-23',6),('2026-10-24',7),('2026-10-25',1),('2026-10-26',2),('2026-10-27',3),('2026-10-28',4),('2026-10-29',5),('2026-10-30',6),('2026-10-31',7),('2026-11-01',1),('2026-11-02',2),('2026-11-03',3),('2026-11-04',4),('2026-11-05',5),('2026-11-06',6),('2026-11-07',7),('2026-11-08',1),('2026-11-09',2),('2026-11-10',3),('2026-11-11',4),('2026-11-12',5),('2026-11-13',6),('2026-11-14',7),('2026-11-15',1),('2026-11-16',2),('2026-11-17',3),('2026-11-18',4),('2026-11-19',5),('2026-11-20',6),('2026-11-21',7),('2026-11-22',1),('2026-11-23',2),('2026-11-24',3),('2026-11-25',4),('2026-11-26',5),('2026-11-27',6),('2026-11-28',7),('2026-11-29',1),('2026-11-30',2),('2026-12-01',3),('2026-12-02',4),('2026-12-03',5),('2026-12-04',6),('2026-12-05',7),('2026-12-06',1),('2026-12-07',2),('2026-12-08',3),('2026-12-09',4),('2026-12-10',5),('2026-12-11',6),('2026-12-12',7),('2026-12-13',1),('2026-12-14',2),('2026-12-15',3),('2026-12-16',4),('2026-12-17',5),('2026-12-18',6),('2026-12-19',7),('2026-12-20',1),('2026-12-21',2),('2026-12-22',3),('2026-12-23',4),('2026-12-24',5),('2026-12-25',6),('2026-12-26',7),('2026-12-27',1),('2026-12-28',2),('2026-12-29',3),('2026-12-30',4),('2026-12-31',5),('2027-01-01',6),('2027-01-02',7),('2027-01-03',1),('2027-01-04',2),('2027-01-05',3),('2027-01-06',4),('2027-01-07',5),('2027-01-08',6),('2027-01-09',7),('2027-01-10',1),('2027-01-11',2),('2027-01-12',3),('2027-01-13',4),('2027-01-14',5),('2027-01-15',6),('2027-01-16',7),('2027-01-17',1),('2027-01-18',2),('2027-01-19',3),('2027-01-20',4),('2027-01-21',5),('2027-01-22',6),('2027-01-23',7),('2027-01-24',1),('2027-01-25',2),('2027-01-26',3),('2027-01-27',4),('2027-01-28',5),('2027-01-29',6),('2027-01-30',7),('2027-01-31',1),('2027-02-01',2),('2027-02-02',3),('2027-02-03',4),('2027-02-04',5),('2027-02-05',6),('2027-02-06',7),('2027-02-07',1),('2027-02-08',2),('2027-02-09',3),('2027-02-10',4),('2027-02-11',5),('2027-02-12',6),('2027-02-13',7),('2027-02-14',1),('2027-02-15',2),('2027-02-16',3),('2027-02-17',4),('2027-02-18',5),('2027-02-19',6),('2027-02-20',7),('2027-02-21',1),('2027-02-22',2),('2027-02-23',3),('2027-02-24',4),('2027-02-25',5),('2027-02-26',6),('2027-02-27',7),('2027-02-28',1),('2027-03-01',2),('2027-03-02',3),('2027-03-03',4),('2027-03-04',5),('2027-03-05',6),('2027-03-06',7),('2027-03-07',1),('2027-03-08',2),('2027-03-09',3),('2027-03-10',4),('2027-03-11',5),('2027-03-12',6),('2027-03-13',7),('2027-03-14',1),('2027-03-15',2),('2027-03-16',3),('2027-03-17',4),('2027-03-18',5),('2027-03-19',6),('2027-03-20',7),('2027-03-21',1),('2027-03-22',2),('2027-03-23',3),('2027-03-24',4),('2027-03-25',5),('2027-03-26',6),('2027-03-27',7),('2027-03-28',1),('2027-03-29',2),('2027-03-30',3),('2027-03-31',4),('2027-04-01',5),('2027-04-02',6),('2027-04-03',7),('2027-04-04',1),('2027-04-05',2),('2027-04-06',3),('2027-04-07',4),('2027-04-08',5),('2027-04-09',6),('2027-04-10',7),('2027-04-11',1),('2027-04-12',2),('2027-04-13',3),('2027-04-14',4),('2027-04-15',5),('2027-04-16',6),('2027-04-17',7),('2027-04-18',1),('2027-04-19',2),('2027-04-20',3),('2027-04-21',4),('2027-04-22',5),('2027-04-23',6),('2027-04-24',7),('2027-04-25',1),('2027-04-26',2),('2027-04-27',3),('2027-04-28',4),('2027-04-29',5),('2027-04-30',6),('2027-05-01',7),('2027-05-02',1),('2027-05-03',2),('2027-05-04',3),('2027-05-05',4),('2027-05-06',5),('2027-05-07',6),('2027-05-08',7),('2027-05-09',1),('2027-05-10',2),('2027-05-11',3),('2027-05-12',4),('2027-05-13',5),('2027-05-14',6),('2027-05-15',7),('2027-05-16',1),('2027-05-17',2),('2027-05-18',3),('2027-05-19',4),('2027-05-20',5),('2027-05-21',6),('2027-05-22',7),('2027-05-23',1),('2027-05-24',2),('2027-05-25',3),('2027-05-26',4),('2027-05-27',5),('2027-05-28',6),('2027-05-29',7),('2027-05-30',1),('2027-05-31',2),('2027-06-01',3),('2027-06-02',4),('2027-06-03',5),('2027-06-04',6),('2027-06-05',7),('2027-06-06',1),('2027-06-07',2),('2027-06-08',3),('2027-06-09',4),('2027-06-10',5),('2027-06-11',6),('2027-06-12',7),('2027-06-13',1),('2027-06-14',2),('2027-06-15',3),('2027-06-16',4),('2027-06-17',5),('2027-06-18',6),('2027-06-19',7),('2027-06-20',1),('2027-06-21',2),('2027-06-22',3),('2027-06-23',4),('2027-06-24',5),('2027-06-25',6),('2027-06-26',7),('2027-06-27',1),('2027-06-28',2),('2027-06-29',3),('2027-06-30',4),('2027-07-01',5),('2027-07-02',6),('2027-07-03',7),('2027-07-04',1),('2027-07-05',2),('2027-07-06',3),('2027-07-07',4),('2027-07-08',5),('2027-07-09',6),('2027-07-10',7),('2027-07-11',1),('2027-07-12',2),('2027-07-13',3),('2027-07-14',4),('2027-07-15',5),('2027-07-16',6),('2027-07-17',7),('2027-07-18',1),('2027-07-19',2),('2027-07-20',3),('2027-07-21',4),('2027-07-22',5),('2027-07-23',6),('2027-07-24',7),('2027-07-25',1),('2027-07-26',2),('2027-07-27',3),('2027-07-28',4),('2027-07-29',5),('2027-07-30',6),('2027-07-31',7),('2027-08-01',1),('2027-08-02',2),('2027-08-03',3),('2027-08-04',4),('2027-08-05',5),('2027-08-06',6),('2027-08-07',7),('2027-08-08',1),('2027-08-09',2),('2027-08-10',3),('2027-08-11',4),('2027-08-12',5),('2027-08-13',6),('2027-08-14',7),('2027-08-15',1),('2027-08-16',2),('2027-08-17',3),('2027-08-18',4),('2027-08-19',5),('2027-08-20',6),('2027-08-21',7),('2027-08-22',1),('2027-08-23',2),('2027-08-24',3),('2027-08-25',4),('2027-08-26',5),('2027-08-27',6),('2027-08-28',7),('2027-08-29',1),('2027-08-30',2),('2027-08-31',3),('2027-09-01',4),('2027-09-02',5),('2027-09-03',6),('2027-09-04',7),('2027-09-05',1),('2027-09-06',2),('2027-09-07',3),('2027-09-08',4),('2027-09-09',5),('2027-09-10',6),('2027-09-11',7),('2027-09-12',1),('2027-09-13',2),('2027-09-14',3),('2027-09-15',4),('2027-09-16',5),('2027-09-17',6),('2027-09-18',7),('2027-09-19',1),('2027-09-20',2),('2027-09-21',3),('2027-09-22',4),('2027-09-23',5),('2027-09-24',6),('2027-09-25',7),('2027-09-26',1),('2027-09-27',2),('2027-09-28',3),('2027-09-29',4),('2027-09-30',5),('2027-10-01',6),('2027-10-02',7),('2027-10-03',1),('2027-10-04',2),('2027-10-05',3),('2027-10-06',4),('2027-10-07',5),('2027-10-08',6),('2027-10-09',7),('2027-10-10',1),('2027-10-11',2),('2027-10-12',3),('2027-10-13',4),('2027-10-14',5),('2027-10-15',6),('2027-10-16',7),('2027-10-17',1),('2027-10-18',2),('2027-10-19',3),('2027-10-20',4),('2027-10-21',5),('2027-10-22',6),('2027-10-23',7),('2027-10-24',1),('2027-10-25',2),('2027-10-26',3),('2027-10-27',4),('2027-10-28',5),('2027-10-29',6),('2027-10-30',7),('2027-10-31',1),('2027-11-01',2),('2027-11-02',3),('2027-11-03',4),('2027-11-04',5),('2027-11-05',6),('2027-11-06',7),('2027-11-07',1),('2027-11-08',2),('2027-11-09',3),('2027-11-10',4),('2027-11-11',5),('2027-11-12',6),('2027-11-13',7),('2027-11-14',1),('2027-11-15',2),('2027-11-16',3),('2027-11-17',4),('2027-11-18',5),('2027-11-19',6),('2027-11-20',7),('2027-11-21',1),('2027-11-22',2),('2027-11-23',3),('2027-11-24',4),('2027-11-25',5),('2027-11-26',6),('2027-11-27',7),('2027-11-28',1),('2027-11-29',2),('2027-11-30',3),('2027-12-01',4),('2027-12-02',5),('2027-12-03',6),('2027-12-04',7),('2027-12-05',1),('2027-12-06',2),('2027-12-07',3),('2027-12-08',4),('2027-12-09',5),('2027-12-10',6),('2027-12-11',7),('2027-12-12',1),('2027-12-13',2),('2027-12-14',3),('2027-12-15',4),('2027-12-16',5),('2027-12-17',6),('2027-12-18',7),('2027-12-19',1),('2027-12-20',2),('2027-12-21',3),('2027-12-22',4),('2027-12-23',5),('2027-12-24',6),('2027-12-25',7),('2027-12-26',1),('2027-12-27',2),('2027-12-28',3),('2027-12-29',4),('2027-12-30',5),('2027-12-31',6),('2028-01-01',7),('2028-01-02',1),('2028-01-03',2),('2028-01-04',3),('2028-01-05',4),('2028-01-06',5),('2028-01-07',6),('2028-01-08',7),('2028-01-09',1),('2028-01-10',2),('2028-01-11',3),('2028-01-12',4),('2028-01-13',5),('2028-01-14',6),('2028-01-15',7),('2028-01-16',1),('2028-01-17',2),('2028-01-18',3),('2028-01-19',4),('2028-01-20',5),('2028-01-21',6),('2028-01-22',7),('2028-01-23',1),('2028-01-24',2),('2028-01-25',3),('2028-01-26',4),('2028-01-27',5),('2028-01-28',6),('2028-01-29',7),('2028-01-30',1),('2028-01-31',2),('2028-02-01',3),('2028-02-02',4),('2028-02-03',5),('2028-02-04',6),('2028-02-05',7),('2028-02-06',1),('2028-02-07',2),('2028-02-08',3),('2028-02-09',4),('2028-02-10',5),('2028-02-11',6),('2028-02-12',7),('2028-02-13',1),('2028-02-14',2),('2028-02-15',3),('2028-02-16',4),('2028-02-17',5),('2028-02-18',6),('2028-02-19',7),('2028-02-20',1),('2028-02-21',2),('2028-02-22',3),('2028-02-23',4),('2028-02-24',5),('2028-02-25',6),('2028-02-26',7),('2028-02-27',1),('2028-02-28',2),('2028-02-29',3),('2028-03-01',4),('2028-03-02',5),('2028-03-03',6),('2028-03-04',7),('2028-03-05',1),('2028-03-06',2),('2028-03-07',3),('2028-03-08',4),('2028-03-09',5),('2028-03-10',6),('2028-03-11',7),('2028-03-12',1),('2028-03-13',2),('2028-03-14',3),('2028-03-15',4),('2028-03-16',5),('2028-03-17',6),('2028-03-18',7),('2028-03-19',1),('2028-03-20',2),('2028-03-21',3),('2028-03-22',4),('2028-03-23',5),('2028-03-24',6),('2028-03-25',7),('2028-03-26',1),('2028-03-27',2),('2028-03-28',3),('2028-03-29',4),('2028-03-30',5),('2028-03-31',6),('2028-04-01',7),('2028-04-02',1),('2028-04-03',2),('2028-04-04',3),('2028-04-05',4),('2028-04-06',5),('2028-04-07',6),('2028-04-08',7),('2028-04-09',1),('2028-04-10',2),('2028-04-11',3),('2028-04-12',4),('2028-04-13',5),('2028-04-14',6),('2028-04-15',7),('2028-04-16',1),('2028-04-17',2),('2028-04-18',3),('2028-04-19',4),('2028-04-20',5),('2028-04-21',6),('2028-04-22',7),('2028-04-23',1),('2028-04-24',2),('2028-04-25',3),('2028-04-26',4),('2028-04-27',5),('2028-04-28',6),('2028-04-29',7),('2028-04-30',1),('2028-05-01',2),('2028-05-02',3),('2028-05-03',4),('2028-05-04',5),('2028-05-05',6),('2028-05-06',7),('2028-05-07',1),('2028-05-08',2),('2028-05-09',3),('2028-05-10',4),('2028-05-11',5),('2028-05-12',6),('2028-05-13',7),('2028-05-14',1),('2028-05-15',2),('2028-05-16',3),('2028-05-17',4),('2028-05-18',5),('2028-05-19',6),('2028-05-20',7),('2028-05-21',1),('2028-05-22',2),('2028-05-23',3),('2028-05-24',4),('2028-05-25',5),('2028-05-26',6),('2028-05-27',7),('2028-05-28',1),('2028-05-29',2),('2028-05-30',3),('2028-05-31',4),('2028-06-01',5),('2028-06-02',6),('2028-06-03',7),('2028-06-04',1),('2028-06-05',2),('2028-06-06',3),('2028-06-07',4),('2028-06-08',5),('2028-06-09',6),('2028-06-10',7),('2028-06-11',1),('2028-06-12',2),('2028-06-13',3),('2028-06-14',4),('2028-06-15',5),('2028-06-16',6),('2028-06-17',7),('2028-06-18',1),('2028-06-19',2),('2028-06-20',3),('2028-06-21',4),('2028-06-22',5),('2028-06-23',6),('2028-06-24',7),('2028-06-25',1),('2028-06-26',2),('2028-06-27',3),('2028-06-28',4),('2028-06-29',5),('2028-06-30',6),('2028-07-01',7),('2028-07-02',1),('2028-07-03',2),('2028-07-04',3),('2028-07-05',4),('2028-07-06',5),('2028-07-07',6),('2028-07-08',7),('2028-07-09',1),('2028-07-10',2),('2028-07-11',3),('2028-07-12',4),('2028-07-13',5),('2028-07-14',6),('2028-07-15',7),('2028-07-16',1),('2028-07-17',2),('2028-07-18',3),('2028-07-19',4),('2028-07-20',5),('2028-07-21',6),('2028-07-22',7),('2028-07-23',1),('2028-07-24',2),('2028-07-25',3),('2028-07-26',4),('2028-07-27',5),('2028-07-28',6),('2028-07-29',7),('2028-07-30',1),('2028-07-31',2),('2028-08-01',3),('2028-08-02',4),('2028-08-03',5),('2028-08-04',6),('2028-08-05',7),('2028-08-06',1),('2028-08-07',2),('2028-08-08',3),('2028-08-09',4),('2028-08-10',5),('2028-08-11',6),('2028-08-12',7),('2028-08-13',1),('2028-08-14',2),('2028-08-15',3),('2028-08-16',4),('2028-08-17',5),('2028-08-18',6),('2028-08-19',7),('2028-08-20',1),('2028-08-21',2),('2028-08-22',3),('2028-08-23',4),('2028-08-24',5),('2028-08-25',6),('2028-08-26',7),('2028-08-27',1),('2028-08-28',2),('2028-08-29',3),('2028-08-30',4),('2028-08-31',5),('2028-09-01',6),('2028-09-02',7),('2028-09-03',1),('2028-09-04',2),('2028-09-05',3),('2028-09-06',4),('2028-09-07',5),('2028-09-08',6),('2028-09-09',7),('2028-09-10',1),('2028-09-11',2),('2028-09-12',3),('2028-09-13',4),('2028-09-14',5),('2028-09-15',6),('2028-09-16',7),('2028-09-17',1),('2028-09-18',2),('2028-09-19',3),('2028-09-20',4),('2028-09-21',5),('2028-09-22',6),('2028-09-23',7),('2028-09-24',1),('2028-09-25',2),('2028-09-26',3),('2028-09-27',4),('2028-09-28',5),('2028-09-29',6),('2028-09-30',7),('2028-10-01',1),('2028-10-02',2),('2028-10-03',3),('2028-10-04',4),('2028-10-05',5),('2028-10-06',6),('2028-10-07',7),('2028-10-08',1),('2028-10-09',2),('2028-10-10',3),('2028-10-11',4),('2028-10-12',5),('2028-10-13',6),('2028-10-14',7),('2028-10-15',1),('2028-10-16',2),('2028-10-17',3),('2028-10-18',4),('2028-10-19',5),('2028-10-20',6),('2028-10-21',7),('2028-10-22',1),('2028-10-23',2),('2028-10-24',3),('2028-10-25',4),('2028-10-26',5),('2028-10-27',6),('2028-10-28',7),('2028-10-29',1),('2028-10-30',2),('2028-10-31',3),('2028-11-01',4),('2028-11-02',5),('2028-11-03',6),('2028-11-04',7),('2028-11-05',1),('2028-11-06',2),('2028-11-07',3),('2028-11-08',4),('2028-11-09',5),('2028-11-10',6),('2028-11-11',7),('2028-11-12',1),('2028-11-13',2),('2028-11-14',3),('2028-11-15',4),('2028-11-16',5),('2028-11-17',6),('2028-11-18',7),('2028-11-19',1),('2028-11-20',2),('2028-11-21',3),('2028-11-22',4),('2028-11-23',5),('2028-11-24',6),('2028-11-25',7),('2028-11-26',1),('2028-11-27',2),('2028-11-28',3),('2028-11-29',4),('2028-11-30',5),('2028-12-01',6),('2028-12-02',7),('2028-12-03',1),('2028-12-04',2),('2028-12-05',3),('2028-12-06',4),('2028-12-07',5),('2028-12-08',6),('2028-12-09',7),('2028-12-10',1),('2028-12-11',2),('2028-12-12',3),('2028-12-13',4),('2028-12-14',5),('2028-12-15',6),('2028-12-16',7),('2028-12-17',1),('2028-12-18',2),('2028-12-19',3),('2028-12-20',4),('2028-12-21',5),('2028-12-22',6),('2028-12-23',7),('2028-12-24',1),('2028-12-25',2),('2028-12-26',3),('2028-12-27',4),('2028-12-28',5),('2028-12-29',6),('2028-12-30',7),('2028-12-31',1),('2029-01-01',2),('2029-01-02',3),('2029-01-03',4),('2029-01-04',5),('2029-01-05',6),('2029-01-06',7),('2029-01-07',1),('2029-01-08',2),('2029-01-09',3),('2029-01-10',4),('2029-01-11',5),('2029-01-12',6),('2029-01-13',7),('2029-01-14',1),('2029-01-15',2),('2029-01-16',3),('2029-01-17',4),('2029-01-18',5),('2029-01-19',6),('2029-01-20',7),('2029-01-21',1),('2029-01-22',2),('2029-01-23',3),('2029-01-24',4),('2029-01-25',5),('2029-01-26',6),('2029-01-27',7),('2029-01-28',1),('2029-01-29',2),('2029-01-30',3),('2029-01-31',4),('2029-02-01',5),('2029-02-02',6),('2029-02-03',7),('2029-02-04',1),('2029-02-05',2),('2029-02-06',3),('2029-02-07',4),('2029-02-08',5),('2029-02-09',6),('2029-02-10',7),('2029-02-11',1),('2029-02-12',2),('2029-02-13',3),('2029-02-14',4),('2029-02-15',5),('2029-02-16',6),('2029-02-17',7),('2029-02-18',1),('2029-02-19',2),('2029-02-20',3),('2029-02-21',4),('2029-02-22',5),('2029-02-23',6),('2029-02-24',7),('2029-02-25',1),('2029-02-26',2),('2029-02-27',3),('2029-02-28',4),('2029-03-01',5),('2029-03-02',6),('2029-03-03',7),('2029-03-04',1),('2029-03-05',2),('2029-03-06',3),('2029-03-07',4),('2029-03-08',5),('2029-03-09',6),('2029-03-10',7),('2029-03-11',1),('2029-03-12',2),('2029-03-13',3),('2029-03-14',4),('2029-03-15',5),('2029-03-16',6),('2029-03-17',7),('2029-03-18',1),('2029-03-19',2),('2029-03-20',3),('2029-03-21',4),('2029-03-22',5),('2029-03-23',6),('2029-03-24',7),('2029-03-25',1),('2029-03-26',2),('2029-03-27',3),('2029-03-28',4),('2029-03-29',5),('2029-03-30',6),('2029-03-31',7),('2029-04-01',1),('2029-04-02',2),('2029-04-03',3),('2029-04-04',4),('2029-04-05',5),('2029-04-06',6),('2029-04-07',7),('2029-04-08',1),('2029-04-09',2),('2029-04-10',3),('2029-04-11',4),('2029-04-12',5),('2029-04-13',6),('2029-04-14',7),('2029-04-15',1),('2029-04-16',2),('2029-04-17',3),('2029-04-18',4),('2029-04-19',5),('2029-04-20',6),('2029-04-21',7),('2029-04-22',1),('2029-04-23',2),('2029-04-24',3),('2029-04-25',4),('2029-04-26',5),('2029-04-27',6),('2029-04-28',7),('2029-04-29',1),('2029-04-30',2),('2029-05-01',3),('2029-05-02',4),('2029-05-03',5),('2029-05-04',6),('2029-05-05',7),('2029-05-06',1),('2029-05-07',2),('2029-05-08',3),('2029-05-09',4),('2029-05-10',5),('2029-05-11',6),('2029-05-12',7),('2029-05-13',1),('2029-05-14',2),('2029-05-15',3),('2029-05-16',4),('2029-05-17',5),('2029-05-18',6),('2029-05-19',7),('2029-05-20',1),('2029-05-21',2),('2029-05-22',3),('2029-05-23',4),('2029-05-24',5),('2029-05-25',6),('2029-05-26',7),('2029-05-27',1),('2029-05-28',2),('2029-05-29',3),('2029-05-30',4),('2029-05-31',5),('2029-06-01',6),('2029-06-02',7),('2029-06-03',1),('2029-06-04',2),('2029-06-05',3),('2029-06-06',4),('2029-06-07',5),('2029-06-08',6),('2029-06-09',7),('2029-06-10',1),('2029-06-11',2),('2029-06-12',3),('2029-06-13',4),('2029-06-14',5),('2029-06-15',6),('2029-06-16',7),('2029-06-17',1),('2029-06-18',2),('2029-06-19',3),('2029-06-20',4),('2029-06-21',5),('2029-06-22',6),('2029-06-23',7),('2029-06-24',1),('2029-06-25',2),('2029-06-26',3),('2029-06-27',4),('2029-06-28',5),('2029-06-29',6),('2029-06-30',7),('2029-07-01',1),('2029-07-02',2),('2029-07-03',3),('2029-07-04',4),('2029-07-05',5),('2029-07-06',6),('2029-07-07',7),('2029-07-08',1),('2029-07-09',2),('2029-07-10',3),('2029-07-11',4),('2029-07-12',5),('2029-07-13',6),('2029-07-14',7),('2029-07-15',1),('2029-07-16',2),('2029-07-17',3),('2029-07-18',4),('2029-07-19',5),('2029-07-20',6),('2029-07-21',7),('2029-07-22',1),('2029-07-23',2),('2029-07-24',3),('2029-07-25',4),('2029-07-26',5),('2029-07-27',6),('2029-07-28',7),('2029-07-29',1),('2029-07-30',2),('2029-07-31',3),('2029-08-01',4),('2029-08-02',5),('2029-08-03',6),('2029-08-04',7),('2029-08-05',1),('2029-08-06',2),('2029-08-07',3),('2029-08-08',4),('2029-08-09',5),('2029-08-10',6),('2029-08-11',7),('2029-08-12',1),('2029-08-13',2),('2029-08-14',3),('2029-08-15',4),('2029-08-16',5),('2029-08-17',6),('2029-08-18',7),('2029-08-19',1),('2029-08-20',2),('2029-08-21',3),('2029-08-22',4),('2029-08-23',5),('2029-08-24',6),('2029-08-25',7),('2029-08-26',1),('2029-08-27',2),('2029-08-28',3),('2029-08-29',4),('2029-08-30',5),('2029-08-31',6),('2029-09-01',7),('2029-09-02',1),('2029-09-03',2),('2029-09-04',3),('2029-09-05',4),('2029-09-06',5),('2029-09-07',6),('2029-09-08',7),('2029-09-09',1),('2029-09-10',2),('2029-09-11',3),('2029-09-12',4),('2029-09-13',5),('2029-09-14',6),('2029-09-15',7),('2029-09-16',1),('2029-09-17',2),('2029-09-18',3),('2029-09-19',4),('2029-09-20',5),('2029-09-21',6),('2029-09-22',7),('2029-09-23',1),('2029-09-24',2),('2029-09-25',3),('2029-09-26',4),('2029-09-27',5),('2029-09-28',6),('2029-09-29',7),('2029-09-30',1),('2029-10-01',2),('2029-10-02',3),('2029-10-03',4),('2029-10-04',5),('2029-10-05',6),('2029-10-06',7),('2029-10-07',1),('2029-10-08',2),('2029-10-09',3),('2029-10-10',4),('2029-10-11',5),('2029-10-12',6),('2029-10-13',7),('2029-10-14',1),('2029-10-15',2),('2029-10-16',3),('2029-10-17',4),('2029-10-18',5),('2029-10-19',6),('2029-10-20',7),('2029-10-21',1),('2029-10-22',2),('2029-10-23',3),('2029-10-24',4),('2029-10-25',5),('2029-10-26',6),('2029-10-27',7),('2029-10-28',1),('2029-10-29',2),('2029-10-30',3),('2029-10-31',4),('2029-11-01',5),('2029-11-02',6),('2029-11-03',7),('2029-11-04',1),('2029-11-05',2),('2029-11-06',3),('2029-11-07',4),('2029-11-08',5),('2029-11-09',6),('2029-11-10',7),('2029-11-11',1),('2029-11-12',2),('2029-11-13',3),('2029-11-14',4),('2029-11-15',5),('2029-11-16',6),('2029-11-17',7),('2029-11-18',1),('2029-11-19',2),('2029-11-20',3),('2029-11-21',4),('2029-11-22',5),('2029-11-23',6),('2029-11-24',7),('2029-11-25',1),('2029-11-26',2),('2029-11-27',3),('2029-11-28',4),('2029-11-29',5),('2029-11-30',6),('2029-12-01',7),('2029-12-02',1),('2029-12-03',2),('2029-12-04',3),('2029-12-05',4),('2029-12-06',5),('2029-12-07',6),('2029-12-08',7),('2029-12-09',1),('2029-12-10',2),('2029-12-11',3),('2029-12-12',4),('2029-12-13',5),('2029-12-14',6),('2029-12-15',7),('2029-12-16',1),('2029-12-17',2),('2029-12-18',3),('2029-12-19',4),('2029-12-20',5),('2029-12-21',6),('2029-12-22',7),('2029-12-23',1),('2029-12-24',2),('2029-12-25',3),('2029-12-26',4),('2029-12-27',5),('2029-12-28',6),('2029-12-29',7),('2029-12-30',1),('2029-12-31',2),('2030-01-01',3),('2030-01-02',4),('2030-01-03',5),('2030-01-04',6),('2030-01-05',7),('2030-01-06',1),('2030-01-07',2),('2030-01-08',3),('2030-01-09',4),('2030-01-10',5),('2030-01-11',6),('2030-01-12',7),('2030-01-13',1),('2030-01-14',2),('2030-01-15',3),('2030-01-16',4),('2030-01-17',5),('2030-01-18',6),('2030-01-19',7),('2030-01-20',1),('2030-01-21',2),('2030-01-22',3),('2030-01-23',4),('2030-01-24',5),('2030-01-25',6),('2030-01-26',7),('2030-01-27',1),('2030-01-28',2),('2030-01-29',3),('2030-01-30',4),('2030-01-31',5),('2030-02-01',6),('2030-02-02',7),('2030-02-03',1),('2030-02-04',2),('2030-02-05',3),('2030-02-06',4),('2030-02-07',5),('2030-02-08',6),('2030-02-09',7),('2030-02-10',1),('2030-02-11',2),('2030-02-12',3),('2030-02-13',4),('2030-02-14',5),('2030-02-15',6),('2030-02-16',7),('2030-02-17',1),('2030-02-18',2),('2030-02-19',3),('2030-02-20',4),('2030-02-21',5),('2030-02-22',6),('2030-02-23',7),('2030-02-24',1),('2030-02-25',2),('2030-02-26',3),('2030-02-27',4),('2030-02-28',5),('2030-03-01',6),('2030-03-02',7),('2030-03-03',1),('2030-03-04',2),('2030-03-05',3),('2030-03-06',4),('2030-03-07',5),('2030-03-08',6),('2030-03-09',7),('2030-03-10',1),('2030-03-11',2),('2030-03-12',3),('2030-03-13',4),('2030-03-14',5),('2030-03-15',6),('2030-03-16',7),('2030-03-17',1),('2030-03-18',2),('2030-03-19',3),('2030-03-20',4),('2030-03-21',5),('2030-03-22',6),('2030-03-23',7),('2030-03-24',1),('2030-03-25',2),('2030-03-26',3),('2030-03-27',4),('2030-03-28',5),('2030-03-29',6),('2030-03-30',7),('2030-03-31',1),('2030-04-01',2),('2030-04-02',3),('2030-04-03',4),('2030-04-04',5),('2030-04-05',6),('2030-04-06',7),('2030-04-07',1),('2030-04-08',2),('2030-04-09',3),('2030-04-10',4),('2030-04-11',5),('2030-04-12',6),('2030-04-13',7),('2030-04-14',1),('2030-04-15',2),('2030-04-16',3),('2030-04-17',4),('2030-04-18',5),('2030-04-19',6),('2030-04-20',7),('2030-04-21',1),('2030-04-22',2),('2030-04-23',3),('2030-04-24',4),('2030-04-25',5),('2030-04-26',6),('2030-04-27',7),('2030-04-28',1),('2030-04-29',2),('2030-04-30',3),('2030-05-01',4),('2030-05-02',5),('2030-05-03',6),('2030-05-04',7),('2030-05-05',1),('2030-05-06',2),('2030-05-07',3),('2030-05-08',4),('2030-05-09',5),('2030-05-10',6),('2030-05-11',7),('2030-05-12',1),('2030-05-13',2),('2030-05-14',3),('2030-05-15',4),('2030-05-16',5),('2030-05-17',6),('2030-05-18',7),('2030-05-19',1),('2030-05-20',2),('2030-05-21',3),('2030-05-22',4),('2030-05-23',5),('2030-05-24',6),('2030-05-25',7),('2030-05-26',1),('2030-05-27',2),('2030-05-28',3),('2030-05-29',4),('2030-05-30',5),('2030-05-31',6),('2030-06-01',7),('2030-06-02',1),('2030-06-03',2),('2030-06-04',3),('2030-06-05',4),('2030-06-06',5),('2030-06-07',6),('2030-06-08',7),('2030-06-09',1),('2030-06-10',2),('2030-06-11',3),('2030-06-12',4),('2030-06-13',5),('2030-06-14',6),('2030-06-15',7),('2030-06-16',1),('2030-06-17',2),('2030-06-18',3),('2030-06-19',4),('2030-06-20',5),('2030-06-21',6),('2030-06-22',7),('2030-06-23',1),('2030-06-24',2),('2030-06-25',3),('2030-06-26',4),('2030-06-27',5),('2030-06-28',6),('2030-06-29',7),('2030-06-30',1),('2030-07-01',2),('2030-07-02',3),('2030-07-03',4),('2030-07-04',5),('2030-07-05',6),('2030-07-06',7),('2030-07-07',1),('2030-07-08',2),('2030-07-09',3),('2030-07-10',4),('2030-07-11',5),('2030-07-12',6),('2030-07-13',7),('2030-07-14',1),('2030-07-15',2),('2030-07-16',3),('2030-07-17',4),('2030-07-18',5),('2030-07-19',6),('2030-07-20',7),('2030-07-21',1),('2030-07-22',2),('2030-07-23',3),('2030-07-24',4),('2030-07-25',5),('2030-07-26',6),('2030-07-27',7),('2030-07-28',1),('2030-07-29',2),('2030-07-30',3),('2030-07-31',4),('2030-08-01',5),('2030-08-02',6),('2030-08-03',7),('2030-08-04',1),('2030-08-05',2),('2030-08-06',3),('2030-08-07',4),('2030-08-08',5),('2030-08-09',6),('2030-08-10',7),('2030-08-11',1),('2030-08-12',2),('2030-08-13',3),('2030-08-14',4),('2030-08-15',5),('2030-08-16',6),('2030-08-17',7),('2030-08-18',1),('2030-08-19',2),('2030-08-20',3),('2030-08-21',4),('2030-08-22',5),('2030-08-23',6),('2030-08-24',7),('2030-08-25',1),('2030-08-26',2),('2030-08-27',3),('2030-08-28',4),('2030-08-29',5),('2030-08-30',6),('2030-08-31',7),('2030-09-01',1),('2030-09-02',2),('2030-09-03',3),('2030-09-04',4),('2030-09-05',5),('2030-09-06',6),('2030-09-07',7),('2030-09-08',1),('2030-09-09',2),('2030-09-10',3),('2030-09-11',4),('2030-09-12',5),('2030-09-13',6),('2030-09-14',7),('2030-09-15',1),('2030-09-16',2),('2030-09-17',3),('2030-09-18',4),('2030-09-19',5),('2030-09-20',6),('2030-09-21',7),('2030-09-22',1),('2030-09-23',2),('2030-09-24',3),('2030-09-25',4),('2030-09-26',5),('2030-09-27',6),('2030-09-28',7),('2030-09-29',1),('2030-09-30',2),('2030-10-01',3),('2030-10-02',4),('2030-10-03',5),('2030-10-04',6),('2030-10-05',7),('2030-10-06',1),('2030-10-07',2),('2030-10-08',3),('2030-10-09',4),('2030-10-10',5),('2030-10-11',6),('2030-10-12',7),('2030-10-13',1),('2030-10-14',2),('2030-10-15',3),('2030-10-16',4),('2030-10-17',5),('2030-10-18',6),('2030-10-19',7),('2030-10-20',1),('2030-10-21',2),('2030-10-22',3),('2030-10-23',4),('2030-10-24',5),('2030-10-25',6),('2030-10-26',7),('2030-10-27',1),('2030-10-28',2),('2030-10-29',3),('2030-10-30',4),('2030-10-31',5),('2030-11-01',6),('2030-11-02',7),('2030-11-03',1),('2030-11-04',2),('2030-11-05',3),('2030-11-06',4),('2030-11-07',5),('2030-11-08',6),('2030-11-09',7),('2030-11-10',1),('2030-11-11',2),('2030-11-12',3),('2030-11-13',4),('2030-11-14',5),('2030-11-15',6),('2030-11-16',7),('2030-11-17',1),('2030-11-18',2),('2030-11-19',3),('2030-11-20',4),('2030-11-21',5),('2030-11-22',6),('2030-11-23',7),('2030-11-24',1),('2030-11-25',2),('2030-11-26',3),('2030-11-27',4),('2030-11-28',5),('2030-11-29',6),('2030-11-30',7),('2030-12-01',1),('2030-12-02',2),('2030-12-03',3),('2030-12-04',4),('2030-12-05',5),('2030-12-06',6),('2030-12-07',7),('2030-12-08',1),('2030-12-09',2),('2030-12-10',3),('2030-12-11',4),('2030-12-12',5),('2030-12-13',6),('2030-12-14',7),('2030-12-15',1),('2030-12-16',2),('2030-12-17',3),('2030-12-18',4),('2030-12-19',5),('2030-12-20',6),('2030-12-21',7),('2030-12-22',1),('2030-12-23',2),('2030-12-24',3),('2030-12-25',4),('2030-12-26',5),('2030-12-27',6),('2030-12-28',7),('2030-12-29',1),('2030-12-30',2),('2030-12-31',3),('2031-01-01',4),('2031-01-02',5),('2031-01-03',6),('2031-01-04',7),('2031-01-05',1),('2031-01-06',2),('2031-01-07',3),('2031-01-08',4),('2031-01-09',5),('2031-01-10',6),('2031-01-11',7),('2031-01-12',1),('2031-01-13',2),('2031-01-14',3),('2031-01-15',4),('2031-01-16',5),('2031-01-17',6),('2031-01-18',7),('2031-01-19',1),('2031-01-20',2),('2031-01-21',3),('2031-01-22',4),('2031-01-23',5),('2031-01-24',6),('2031-01-25',7),('2031-01-26',1),('2031-01-27',2),('2031-01-28',3),('2031-01-29',4),('2031-01-30',5),('2031-01-31',6),('2031-02-01',7),('2031-02-02',1),('2031-02-03',2),('2031-02-04',3),('2031-02-05',4),('2031-02-06',5),('2031-02-07',6),('2031-02-08',7),('2031-02-09',1),('2031-02-10',2),('2031-02-11',3),('2031-02-12',4),('2031-02-13',5),('2031-02-14',6),('2031-02-15',7),('2031-02-16',1),('2031-02-17',2),('2031-02-18',3),('2031-02-19',4),('2031-02-20',5),('2031-02-21',6),('2031-02-22',7),('2031-02-23',1),('2031-02-24',2),('2031-02-25',3),('2031-02-26',4),('2031-02-27',5),('2031-02-28',6),('2031-03-01',7),('2031-03-02',1),('2031-03-03',2),('2031-03-04',3),('2031-03-05',4),('2031-03-06',5),('2031-03-07',6),('2031-03-08',7),('2031-03-09',1),('2031-03-10',2),('2031-03-11',3),('2031-03-12',4),('2031-03-13',5),('2031-03-14',6),('2031-03-15',7),('2031-03-16',1),('2031-03-17',2),('2031-03-18',3),('2031-03-19',4),('2031-03-20',5),('2031-03-21',6),('2031-03-22',7),('2031-03-23',1),('2031-03-24',2),('2031-03-25',3),('2031-03-26',4),('2031-03-27',5),('2031-03-28',6),('2031-03-29',7),('2031-03-30',1),('2031-03-31',2),('2031-04-01',3),('2031-04-02',4),('2031-04-03',5),('2031-04-04',6),('2031-04-05',7),('2031-04-06',1),('2031-04-07',2),('2031-04-08',3),('2031-04-09',4),('2031-04-10',5),('2031-04-11',6),('2031-04-12',7),('2031-04-13',1),('2031-04-14',2),('2031-04-15',3),('2031-04-16',4),('2031-04-17',5),('2031-04-18',6),('2031-04-19',7),('2031-04-20',1),('2031-04-21',2),('2031-04-22',3),('2031-04-23',4),('2031-04-24',5),('2031-04-25',6),('2031-04-26',7),('2031-04-27',1),('2031-04-28',2),('2031-04-29',3),('2031-04-30',4),('2031-05-01',5),('2031-05-02',6),('2031-05-03',7),('2031-05-04',1),('2031-05-05',2),('2031-05-06',3),('2031-05-07',4),('2031-05-08',5),('2031-05-09',6),('2031-05-10',7),('2031-05-11',1),('2031-05-12',2),('2031-05-13',3),('2031-05-14',4),('2031-05-15',5),('2031-05-16',6),('2031-05-17',7),('2031-05-18',1),('2031-05-19',2),('2031-05-20',3),('2031-05-21',4),('2031-05-22',5),('2031-05-23',6),('2031-05-24',7),('2031-05-25',1),('2031-05-26',2),('2031-05-27',3),('2031-05-28',4),('2031-05-29',5),('2031-05-30',6),('2031-05-31',7),('2031-06-01',1),('2031-06-02',2),('2031-06-03',3),('2031-06-04',4),('2031-06-05',5),('2031-06-06',6),('2031-06-07',7),('2031-06-08',1),('2031-06-09',2),('2031-06-10',3),('2031-06-11',4),('2031-06-12',5),('2031-06-13',6),('2031-06-14',7),('2031-06-15',1),('2031-06-16',2),('2031-06-17',3),('2031-06-18',4),('2031-06-19',5),('2031-06-20',6),('2031-06-21',7),('2031-06-22',1),('2031-06-23',2),('2031-06-24',3),('2031-06-25',4),('2031-06-26',5),('2031-06-27',6),('2031-06-28',7),('2031-06-29',1),('2031-06-30',2),('2031-07-01',3),('2031-07-02',4),('2031-07-03',5),('2031-07-04',6),('2031-07-05',7),('2031-07-06',1),('2031-07-07',2),('2031-07-08',3),('2031-07-09',4),('2031-07-10',5),('2031-07-11',6),('2031-07-12',7),('2031-07-13',1),('2031-07-14',2),('2031-07-15',3),('2031-07-16',4),('2031-07-17',5),('2031-07-18',6),('2031-07-19',7),('2031-07-20',1),('2031-07-21',2),('2031-07-22',3),('2031-07-23',4),('2031-07-24',5),('2031-07-25',6),('2031-07-26',7),('2031-07-27',1),('2031-07-28',2),('2031-07-29',3),('2031-07-30',4),('2031-07-31',5),('2031-08-01',6),('2031-08-02',7),('2031-08-03',1),('2031-08-04',2),('2031-08-05',3),('2031-08-06',4),('2031-08-07',5),('2031-08-08',6),('2031-08-09',7),('2031-08-10',1),('2031-08-11',2),('2031-08-12',3),('2031-08-13',4),('2031-08-14',5),('2031-08-15',6),('2031-08-16',7),('2031-08-17',1),('2031-08-18',2),('2031-08-19',3),('2031-08-20',4),('2031-08-21',5),('2031-08-22',6),('2031-08-23',7),('2031-08-24',1),('2031-08-25',2),('2031-08-26',3),('2031-08-27',4),('2031-08-28',5),('2031-08-29',6),('2031-08-30',7),('2031-08-31',1),('2031-09-01',2),('2031-09-02',3),('2031-09-03',4),('2031-09-04',5),('2031-09-05',6),('2031-09-06',7),('2031-09-07',1),('2031-09-08',2),('2031-09-09',3),('2031-09-10',4),('2031-09-11',5),('2031-09-12',6),('2031-09-13',7),('2031-09-14',1),('2031-09-15',2),('2031-09-16',3),('2031-09-17',4),('2031-09-18',5),('2031-09-19',6),('2031-09-20',7),('2031-09-21',1),('2031-09-22',2),('2031-09-23',3),('2031-09-24',4),('2031-09-25',5),('2031-09-26',6),('2031-09-27',7),('2031-09-28',1),('2031-09-29',2),('2031-09-30',3),('2031-10-01',4),('2031-10-02',5),('2031-10-03',6),('2031-10-04',7),('2031-10-05',1),('2031-10-06',2),('2031-10-07',3),('2031-10-08',4),('2031-10-09',5),('2031-10-10',6),('2031-10-11',7),('2031-10-12',1),('2031-10-13',2),('2031-10-14',3),('2031-10-15',4),('2031-10-16',5),('2031-10-17',6),('2031-10-18',7),('2031-10-19',1),('2031-10-20',2),('2031-10-21',3),('2031-10-22',4),('2031-10-23',5),('2031-10-24',6),('2031-10-25',7),('2031-10-26',1),('2031-10-27',2),('2031-10-28',3),('2031-10-29',4),('2031-10-30',5),('2031-10-31',6),('2031-11-01',7),('2031-11-02',1),('2031-11-03',2),('2031-11-04',3),('2031-11-05',4),('2031-11-06',5),('2031-11-07',6),('2031-11-08',7),('2031-11-09',1),('2031-11-10',2),('2031-11-11',3),('2031-11-12',4),('2031-11-13',5),('2031-11-14',6),('2031-11-15',7),('2031-11-16',1),('2031-11-17',2),('2031-11-18',3),('2031-11-19',4),('2031-11-20',5),('2031-11-21',6),('2031-11-22',7),('2031-11-23',1),('2031-11-24',2),('2031-11-25',3),('2031-11-26',4),('2031-11-27',5),('2031-11-28',6),('2031-11-29',7),('2031-11-30',1),('2031-12-01',2),('2031-12-02',3),('2031-12-03',4),('2031-12-04',5),('2031-12-05',6),('2031-12-06',7),('2031-12-07',1),('2031-12-08',2),('2031-12-09',3),('2031-12-10',4),('2031-12-11',5),('2031-12-12',6),('2031-12-13',7),('2031-12-14',1),('2031-12-15',2),('2031-12-16',3),('2031-12-17',4),('2031-12-18',5),('2031-12-19',6),('2031-12-20',7),('2031-12-21',1),('2031-12-22',2),('2031-12-23',3),('2031-12-24',4),('2031-12-25',5),('2031-12-26',6),('2031-12-27',7),('2031-12-28',1),('2031-12-29',2),('2031-12-30',3),('2031-12-31',4),('2032-01-01',5),('2032-01-02',6),('2032-01-03',7),('2032-01-04',1),('2032-01-05',2),('2032-01-06',3),('2032-01-07',4),('2032-01-08',5),('2032-01-09',6),('2032-01-10',7),('2032-01-11',1),('2032-01-12',2),('2032-01-13',3),('2032-01-14',4),('2032-01-15',5),('2032-01-16',6),('2032-01-17',7),('2032-01-18',1),('2032-01-19',2),('2032-01-20',3),('2032-01-21',4),('2032-01-22',5),('2032-01-23',6),('2032-01-24',7),('2032-01-25',1),('2032-01-26',2),('2032-01-27',3),('2032-01-28',4),('2032-01-29',5),('2032-01-30',6),('2032-01-31',7),('2032-02-01',1),('2032-02-02',2),('2032-02-03',3),('2032-02-04',4),('2032-02-05',5),('2032-02-06',6),('2032-02-07',7),('2032-02-08',1),('2032-02-09',2),('2032-02-10',3),('2032-02-11',4),('2032-02-12',5),('2032-02-13',6),('2032-02-14',7),('2032-02-15',1),('2032-02-16',2),('2032-02-17',3),('2032-02-18',4),('2032-02-19',5),('2032-02-20',6),('2032-02-21',7),('2032-02-22',1),('2032-02-23',2),('2032-02-24',3),('2032-02-25',4),('2032-02-26',5),('2032-02-27',6),('2032-02-28',7),('2032-02-29',1),('2032-03-01',2),('2032-03-02',3),('2032-03-03',4),('2032-03-04',5),('2032-03-05',6),('2032-03-06',7),('2032-03-07',1),('2032-03-08',2),('2032-03-09',3),('2032-03-10',4),('2032-03-11',5),('2032-03-12',6),('2032-03-13',7),('2032-03-14',1),('2032-03-15',2),('2032-03-16',3),('2032-03-17',4),('2032-03-18',5),('2032-03-19',6),('2032-03-20',7),('2032-03-21',1),('2032-03-22',2),('2032-03-23',3),('2032-03-24',4),('2032-03-25',5),('2032-03-26',6),('2032-03-27',7),('2032-03-28',1),('2032-03-29',2),('2032-03-30',3),('2032-03-31',4),('2032-04-01',5),('2032-04-02',6),('2032-04-03',7),('2032-04-04',1),('2032-04-05',2),('2032-04-06',3),('2032-04-07',4),('2032-04-08',5),('2032-04-09',6),('2032-04-10',7),('2032-04-11',1),('2032-04-12',2),('2032-04-13',3),('2032-04-14',4),('2032-04-15',5),('2032-04-16',6),('2032-04-17',7),('2032-04-18',1),('2032-04-19',2),('2032-04-20',3),('2032-04-21',4),('2032-04-22',5),('2032-04-23',6),('2032-04-24',7),('2032-04-25',1),('2032-04-26',2),('2032-04-27',3),('2032-04-28',4),('2032-04-29',5),('2032-04-30',6),('2032-05-01',7),('2032-05-02',1),('2032-05-03',2),('2032-05-04',3),('2032-05-05',4),('2032-05-06',5),('2032-05-07',6),('2032-05-08',7),('2032-05-09',1),('2032-05-10',2),('2032-05-11',3),('2032-05-12',4),('2032-05-13',5),('2032-05-14',6),('2032-05-15',7),('2032-05-16',1),('2032-05-17',2),('2032-05-18',3),('2032-05-19',4),('2032-05-20',5),('2032-05-21',6),('2032-05-22',7),('2032-05-23',1),('2032-05-24',2),('2032-05-25',3),('2032-05-26',4),('2032-05-27',5),('2032-05-28',6),('2032-05-29',7),('2032-05-30',1),('2032-05-31',2),('2032-06-01',3),('2032-06-02',4),('2032-06-03',5),('2032-06-04',6),('2032-06-05',7),('2032-06-06',1),('2032-06-07',2),('2032-06-08',3),('2032-06-09',4),('2032-06-10',5),('2032-06-11',6),('2032-06-12',7),('2032-06-13',1),('2032-06-14',2),('2032-06-15',3),('2032-06-16',4),('2032-06-17',5),('2032-06-18',6),('2032-06-19',7),('2032-06-20',1),('2032-06-21',2),('2032-06-22',3),('2032-06-23',4),('2032-06-24',5),('2032-06-25',6),('2032-06-26',7),('2032-06-27',1),('2032-06-28',2),('2032-06-29',3),('2032-06-30',4),('2032-07-01',5),('2032-07-02',6),('2032-07-03',7),('2032-07-04',1),('2032-07-05',2),('2032-07-06',3),('2032-07-07',4),('2032-07-08',5),('2032-07-09',6),('2032-07-10',7),('2032-07-11',1),('2032-07-12',2),('2032-07-13',3),('2032-07-14',4),('2032-07-15',5),('2032-07-16',6),('2032-07-17',7),('2032-07-18',1),('2032-07-19',2),('2032-07-20',3),('2032-07-21',4),('2032-07-22',5),('2032-07-23',6),('2032-07-24',7),('2032-07-25',1),('2032-07-26',2),('2032-07-27',3),('2032-07-28',4),('2032-07-29',5),('2032-07-30',6),('2032-07-31',7),('2032-08-01',1),('2032-08-02',2),('2032-08-03',3),('2032-08-04',4),('2032-08-05',5),('2032-08-06',6),('2032-08-07',7),('2032-08-08',1),('2032-08-09',2),('2032-08-10',3),('2032-08-11',4),('2032-08-12',5),('2032-08-13',6),('2032-08-14',7),('2032-08-15',1),('2032-08-16',2),('2032-08-17',3),('2032-08-18',4),('2032-08-19',5),('2032-08-20',6),('2032-08-21',7),('2032-08-22',1),('2032-08-23',2),('2032-08-24',3),('2032-08-25',4),('2032-08-26',5),('2032-08-27',6),('2032-08-28',7),('2032-08-29',1),('2032-08-30',2),('2032-08-31',3),('2032-09-01',4),('2032-09-02',5),('2032-09-03',6),('2032-09-04',7),('2032-09-05',1),('2032-09-06',2),('2032-09-07',3),('2032-09-08',4),('2032-09-09',5),('2032-09-10',6),('2032-09-11',7),('2032-09-12',1),('2032-09-13',2),('2032-09-14',3),('2032-09-15',4),('2032-09-16',5),('2032-09-17',6),('2032-09-18',7),('2032-09-19',1),('2032-09-20',2),('2032-09-21',3),('2032-09-22',4),('2032-09-23',5),('2032-09-24',6),('2032-09-25',7),('2032-09-26',1),('2032-09-27',2),('2032-09-28',3),('2032-09-29',4),('2032-09-30',5),('2032-10-01',6),('2032-10-02',7),('2032-10-03',1),('2032-10-04',2),('2032-10-05',3),('2032-10-06',4),('2032-10-07',5),('2032-10-08',6),('2032-10-09',7),('2032-10-10',1),('2032-10-11',2),('2032-10-12',3),('2032-10-13',4),('2032-10-14',5),('2032-10-15',6),('2032-10-16',7),('2032-10-17',1),('2032-10-18',2),('2032-10-19',3),('2032-10-20',4),('2032-10-21',5),('2032-10-22',6),('2032-10-23',7),('2032-10-24',1),('2032-10-25',2),('2032-10-26',3),('2032-10-27',4),('2032-10-28',5),('2032-10-29',6),('2032-10-30',7),('2032-10-31',1),('2032-11-01',2),('2032-11-02',3),('2032-11-03',4),('2032-11-04',5),('2032-11-05',6),('2032-11-06',7),('2032-11-07',1),('2032-11-08',2),('2032-11-09',3),('2032-11-10',4),('2032-11-11',5),('2032-11-12',6),('2032-11-13',7),('2032-11-14',1),('2032-11-15',2),('2032-11-16',3),('2032-11-17',4),('2032-11-18',5),('2032-11-19',6),('2032-11-20',7),('2032-11-21',1),('2032-11-22',2),('2032-11-23',3),('2032-11-24',4),('2032-11-25',5),('2032-11-26',6),('2032-11-27',7),('2032-11-28',1),('2032-11-29',2),('2032-11-30',3),('2032-12-01',4),('2032-12-02',5),('2032-12-03',6),('2032-12-04',7),('2032-12-05',1),('2032-12-06',2),('2032-12-07',3),('2032-12-08',4),('2032-12-09',5),('2032-12-10',6),('2032-12-11',7),('2032-12-12',1),('2032-12-13',2),('2032-12-14',3),('2032-12-15',4),('2032-12-16',5),('2032-12-17',6),('2032-12-18',7),('2032-12-19',1),('2032-12-20',2),('2032-12-21',3),('2032-12-22',4),('2032-12-23',5),('2032-12-24',6),('2032-12-25',7),('2032-12-26',1),('2032-12-27',2),('2032-12-28',3),('2032-12-29',4),('2032-12-30',5),('2032-12-31',6),('2033-01-01',7),('2033-01-02',1),('2033-01-03',2),('2033-01-04',3),('2033-01-05',4),('2033-01-06',5),('2033-01-07',6),('2033-01-08',7),('2033-01-09',1),('2033-01-10',2),('2033-01-11',3),('2033-01-12',4),('2033-01-13',5),('2033-01-14',6),('2033-01-15',7),('2033-01-16',1),('2033-01-17',2),('2033-01-18',3),('2033-01-19',4),('2033-01-20',5),('2033-01-21',6),('2033-01-22',7),('2033-01-23',1),('2033-01-24',2),('2033-01-25',3),('2033-01-26',4),('2033-01-27',5),('2033-01-28',6),('2033-01-29',7),('2033-01-30',1),('2033-01-31',2),('2033-02-01',3),('2033-02-02',4),('2033-02-03',5),('2033-02-04',6),('2033-02-05',7),('2033-02-06',1),('2033-02-07',2),('2033-02-08',3),('2033-02-09',4),('2033-02-10',5),('2033-02-11',6),('2033-02-12',7),('2033-02-13',1),('2033-02-14',2),('2033-02-15',3),('2033-02-16',4),('2033-02-17',5),('2033-02-18',6),('2033-02-19',7),('2033-02-20',1),('2033-02-21',2),('2033-02-22',3),('2033-02-23',4),('2033-02-24',5),('2033-02-25',6),('2033-02-26',7),('2033-02-27',1),('2033-02-28',2),('2033-03-01',3),('2033-03-02',4),('2033-03-03',5),('2033-03-04',6),('2033-03-05',7),('2033-03-06',1),('2033-03-07',2),('2033-03-08',3),('2033-03-09',4),('2033-03-10',5),('2033-03-11',6),('2033-03-12',7),('2033-03-13',1),('2033-03-14',2),('2033-03-15',3),('2033-03-16',4),('2033-03-17',5),('2033-03-18',6),('2033-03-19',7),('2033-03-20',1),('2033-03-21',2),('2033-03-22',3),('2033-03-23',4),('2033-03-24',5),('2033-03-25',6),('2033-03-26',7),('2033-03-27',1),('2033-03-28',2),('2033-03-29',3),('2033-03-30',4),('2033-03-31',5),('2033-04-01',6),('2033-04-02',7),('2033-04-03',1),('2033-04-04',2),('2033-04-05',3),('2033-04-06',4),('2033-04-07',5),('2033-04-08',6),('2033-04-09',7),('2033-04-10',1),('2033-04-11',2),('2033-04-12',3),('2033-04-13',4),('2033-04-14',5),('2033-04-15',6),('2033-04-16',7),('2033-04-17',1),('2033-04-18',2),('2033-04-19',3),('2033-04-20',4),('2033-04-21',5),('2033-04-22',6),('2033-04-23',7),('2033-04-24',1),('2033-04-25',2),('2033-04-26',3),('2033-04-27',4),('2033-04-28',5),('2033-04-29',6),('2033-04-30',7),('2033-05-01',1),('2033-05-02',2),('2033-05-03',3),('2033-05-04',4),('2033-05-05',5),('2033-05-06',6),('2033-05-07',7),('2033-05-08',1),('2033-05-09',2),('2033-05-10',3),('2033-05-11',4),('2033-05-12',5),('2033-05-13',6),('2033-05-14',7),('2033-05-15',1),('2033-05-16',2),('2033-05-17',3),('2033-05-18',4),('2033-05-19',5),('2033-05-20',6),('2033-05-21',7),('2033-05-22',1),('2033-05-23',2),('2033-05-24',3),('2033-05-25',4),('2033-05-26',5),('2033-05-27',6),('2033-05-28',7),('2033-05-29',1),('2033-05-30',2),('2033-05-31',3),('2033-06-01',4),('2033-06-02',5),('2033-06-03',6),('2033-06-04',7),('2033-06-05',1),('2033-06-06',2),('2033-06-07',3),('2033-06-08',4),('2033-06-09',5),('2033-06-10',6),('2033-06-11',7),('2033-06-12',1),('2033-06-13',2),('2033-06-14',3),('2033-06-15',4),('2033-06-16',5),('2033-06-17',6),('2033-06-18',7),('2033-06-19',1),('2033-06-20',2),('2033-06-21',3),('2033-06-22',4),('2033-06-23',5),('2033-06-24',6),('2033-06-25',7),('2033-06-26',1),('2033-06-27',2),('2033-06-28',3),('2033-06-29',4),('2033-06-30',5),('2033-07-01',6),('2033-07-02',7),('2033-07-03',1),('2033-07-04',2),('2033-07-05',3),('2033-07-06',4),('2033-07-07',5),('2033-07-08',6),('2033-07-09',7),('2033-07-10',1),('2033-07-11',2),('2033-07-12',3),('2033-07-13',4),('2033-07-14',5),('2033-07-15',6),('2033-07-16',7),('2033-07-17',1),('2033-07-18',2),('2033-07-19',3),('2033-07-20',4),('2033-07-21',5),('2033-07-22',6),('2033-07-23',7),('2033-07-24',1),('2033-07-25',2),('2033-07-26',3),('2033-07-27',4),('2033-07-28',5),('2033-07-29',6),('2033-07-30',7),('2033-07-31',1),('2033-08-01',2),('2033-08-02',3),('2033-08-03',4),('2033-08-04',5),('2033-08-05',6),('2033-08-06',7),('2033-08-07',1),('2033-08-08',2),('2033-08-09',3),('2033-08-10',4),('2033-08-11',5),('2033-08-12',6),('2033-08-13',7),('2033-08-14',1),('2033-08-15',2),('2033-08-16',3),('2033-08-17',4),('2033-08-18',5),('2033-08-19',6),('2033-08-20',7),('2033-08-21',1),('2033-08-22',2),('2033-08-23',3),('2033-08-24',4),('2033-08-25',5),('2033-08-26',6),('2033-08-27',7),('2033-08-28',1),('2033-08-29',2),('2033-08-30',3),('2033-08-31',4),('2033-09-01',5),('2033-09-02',6),('2033-09-03',7),('2033-09-04',1),('2033-09-05',2),('2033-09-06',3),('2033-09-07',4),('2033-09-08',5),('2033-09-09',6),('2033-09-10',7),('2033-09-11',1),('2033-09-12',2),('2033-09-13',3),('2033-09-14',4),('2033-09-15',5),('2033-09-16',6),('2033-09-17',7),('2033-09-18',1),('2033-09-19',2),('2033-09-20',3),('2033-09-21',4),('2033-09-22',5),('2033-09-23',6),('2033-09-24',7),('2033-09-25',1),('2033-09-26',2),('2033-09-27',3),('2033-09-28',4),('2033-09-29',5),('2033-09-30',6),('2033-10-01',7),('2033-10-02',1),('2033-10-03',2),('2033-10-04',3),('2033-10-05',4),('2033-10-06',5),('2033-10-07',6),('2033-10-08',7),('2033-10-09',1),('2033-10-10',2),('2033-10-11',3),('2033-10-12',4),('2033-10-13',5),('2033-10-14',6),('2033-10-15',7),('2033-10-16',1),('2033-10-17',2),('2033-10-18',3),('2033-10-19',4),('2033-10-20',5),('2033-10-21',6),('2033-10-22',7),('2033-10-23',1),('2033-10-24',2),('2033-10-25',3),('2033-10-26',4),('2033-10-27',5),('2033-10-28',6),('2033-10-29',7),('2033-10-30',1),('2033-10-31',2),('2033-11-01',3),('2033-11-02',4),('2033-11-03',5),('2033-11-04',6),('2033-11-05',7),('2033-11-06',1),('2033-11-07',2),('2033-11-08',3),('2033-11-09',4),('2033-11-10',5),('2033-11-11',6),('2033-11-12',7),('2033-11-13',1),('2033-11-14',2),('2033-11-15',3),('2033-11-16',4),('2033-11-17',5),('2033-11-18',6),('2033-11-19',7),('2033-11-20',1),('2033-11-21',2),('2033-11-22',3),('2033-11-23',4),('2033-11-24',5),('2033-11-25',6),('2033-11-26',7),('2033-11-27',1),('2033-11-28',2),('2033-11-29',3),('2033-11-30',4),('2033-12-01',5),('2033-12-02',6),('2033-12-03',7),('2033-12-04',1),('2033-12-05',2),('2033-12-06',3),('2033-12-07',4),('2033-12-08',5),('2033-12-09',6),('2033-12-10',7),('2033-12-11',1),('2033-12-12',2),('2033-12-13',3),('2033-12-14',4),('2033-12-15',5),('2033-12-16',6),('2033-12-17',7),('2033-12-18',1),('2033-12-19',2),('2033-12-20',3),('2033-12-21',4),('2033-12-22',5),('2033-12-23',6),('2033-12-24',7),('2033-12-25',1),('2033-12-26',2),('2033-12-27',3),('2033-12-28',4),('2033-12-29',5),('2033-12-30',6),('2033-12-31',7),('2034-01-01',1),('2034-01-02',2),('2034-01-03',3),('2034-01-04',4),('2034-01-05',5),('2034-01-06',6),('2034-01-07',7),('2034-01-08',1),('2034-01-09',2),('2034-01-10',3),('2034-01-11',4),('2034-01-12',5),('2034-01-13',6),('2034-01-14',7),('2034-01-15',1),('2034-01-16',2),('2034-01-17',3),('2034-01-18',4),('2034-01-19',5),('2034-01-20',6),('2034-01-21',7),('2034-01-22',1),('2034-01-23',2),('2034-01-24',3),('2034-01-25',4),('2034-01-26',5),('2034-01-27',6),('2034-01-28',7),('2034-01-29',1),('2034-01-30',2),('2034-01-31',3),('2034-02-01',4),('2034-02-02',5),('2034-02-03',6),('2034-02-04',7),('2034-02-05',1),('2034-02-06',2),('2034-02-07',3),('2034-02-08',4),('2034-02-09',5),('2034-02-10',6),('2034-02-11',7),('2034-02-12',1),('2034-02-13',2),('2034-02-14',3),('2034-02-15',4),('2034-02-16',5),('2034-02-17',6),('2034-02-18',7),('2034-02-19',1),('2034-02-20',2),('2034-02-21',3),('2034-02-22',4),('2034-02-23',5),('2034-02-24',6),('2034-02-25',7),('2034-02-26',1),('2034-02-27',2),('2034-02-28',3),('2034-03-01',4),('2034-03-02',5),('2034-03-03',6),('2034-03-04',7),('2034-03-05',1),('2034-03-06',2),('2034-03-07',3),('2034-03-08',4),('2034-03-09',5),('2034-03-10',6),('2034-03-11',7),('2034-03-12',1),('2034-03-13',2),('2034-03-14',3),('2034-03-15',4),('2034-03-16',5),('2034-03-17',6),('2034-03-18',7),('2034-03-19',1),('2034-03-20',2),('2034-03-21',3),('2034-03-22',4),('2034-03-23',5),('2034-03-24',6),('2034-03-25',7),('2034-03-26',1),('2034-03-27',2),('2034-03-28',3),('2034-03-29',4),('2034-03-30',5),('2034-03-31',6),('2034-04-01',7),('2034-04-02',1),('2034-04-03',2),('2034-04-04',3),('2034-04-05',4),('2034-04-06',5),('2034-04-07',6),('2034-04-08',7),('2034-04-09',1),('2034-04-10',2),('2034-04-11',3),('2034-04-12',4),('2034-04-13',5),('2034-04-14',6),('2034-04-15',7),('2034-04-16',1),('2034-04-17',2),('2034-04-18',3),('2034-04-19',4),('2034-04-20',5),('2034-04-21',6),('2034-04-22',7),('2034-04-23',1),('2034-04-24',2),('2034-04-25',3),('2034-04-26',4),('2034-04-27',5),('2034-04-28',6),('2034-04-29',7),('2034-04-30',1),('2034-05-01',2),('2034-05-02',3),('2034-05-03',4),('2034-05-04',5),('2034-05-05',6),('2034-05-06',7),('2034-05-07',1),('2034-05-08',2),('2034-05-09',3),('2034-05-10',4),('2034-05-11',5),('2034-05-12',6),('2034-05-13',7),('2034-05-14',1),('2034-05-15',2),('2034-05-16',3),('2034-05-17',4),('2034-05-18',5),('2034-05-19',6),('2034-05-20',7),('2034-05-21',1),('2034-05-22',2),('2034-05-23',3),('2034-05-24',4),('2034-05-25',5),('2034-05-26',6),('2034-05-27',7),('2034-05-28',1),('2034-05-29',2),('2034-05-30',3),('2034-05-31',4),('2034-06-01',5),('2034-06-02',6),('2034-06-03',7),('2034-06-04',1),('2034-06-05',2),('2034-06-06',3),('2034-06-07',4),('2034-06-08',5),('2034-06-09',6),('2034-06-10',7),('2034-06-11',1),('2034-06-12',2),('2034-06-13',3),('2034-06-14',4),('2034-06-15',5),('2034-06-16',6),('2034-06-17',7),('2034-06-18',1),('2034-06-19',2),('2034-06-20',3),('2034-06-21',4),('2034-06-22',5),('2034-06-23',6),('2034-06-24',7),('2034-06-25',1),('2034-06-26',2),('2034-06-27',3),('2034-06-28',4),('2034-06-29',5),('2034-06-30',6),('2034-07-01',7),('2034-07-02',1),('2034-07-03',2),('2034-07-04',3),('2034-07-05',4),('2034-07-06',5),('2034-07-07',6),('2034-07-08',7),('2034-07-09',1),('2034-07-10',2),('2034-07-11',3),('2034-07-12',4),('2034-07-13',5),('2034-07-14',6),('2034-07-15',7),('2034-07-16',1),('2034-07-17',2),('2034-07-18',3),('2034-07-19',4),('2034-07-20',5),('2034-07-21',6),('2034-07-22',7),('2034-07-23',1),('2034-07-24',2),('2034-07-25',3),('2034-07-26',4),('2034-07-27',5),('2034-07-28',6),('2034-07-29',7),('2034-07-30',1),('2034-07-31',2),('2034-08-01',3),('2034-08-02',4),('2034-08-03',5),('2034-08-04',6),('2034-08-05',7),('2034-08-06',1),('2034-08-07',2),('2034-08-08',3),('2034-08-09',4),('2034-08-10',5),('2034-08-11',6),('2034-08-12',7),('2034-08-13',1),('2034-08-14',2),('2034-08-15',3),('2034-08-16',4),('2034-08-17',5),('2034-08-18',6),('2034-08-19',7),('2034-08-20',1),('2034-08-21',2),('2034-08-22',3),('2034-08-23',4),('2034-08-24',5),('2034-08-25',6),('2034-08-26',7),('2034-08-27',1),('2034-08-28',2),('2034-08-29',3),('2034-08-30',4),('2034-08-31',5),('2034-09-01',6),('2034-09-02',7),('2034-09-03',1),('2034-09-04',2),('2034-09-05',3),('2034-09-06',4),('2034-09-07',5),('2034-09-08',6),('2034-09-09',7),('2034-09-10',1),('2034-09-11',2),('2034-09-12',3),('2034-09-13',4),('2034-09-14',5),('2034-09-15',6),('2034-09-16',7),('2034-09-17',1),('2034-09-18',2),('2034-09-19',3),('2034-09-20',4),('2034-09-21',5),('2034-09-22',6),('2034-09-23',7),('2034-09-24',1),('2034-09-25',2),('2034-09-26',3),('2034-09-27',4),('2034-09-28',5),('2034-09-29',6),('2034-09-30',7),('2034-10-01',1),('2034-10-02',2),('2034-10-03',3),('2034-10-04',4),('2034-10-05',5),('2034-10-06',6),('2034-10-07',7),('2034-10-08',1),('2034-10-09',2),('2034-10-10',3),('2034-10-11',4),('2034-10-12',5),('2034-10-13',6),('2034-10-14',7),('2034-10-15',1),('2034-10-16',2),('2034-10-17',3),('2034-10-18',4),('2034-10-19',5),('2034-10-20',6),('2034-10-21',7),('2034-10-22',1),('2034-10-23',2),('2034-10-24',3),('2034-10-25',4),('2034-10-26',5),('2034-10-27',6),('2034-10-28',7),('2034-10-29',1),('2034-10-30',2),('2034-10-31',3),('2034-11-01',4),('2034-11-02',5),('2034-11-03',6),('2034-11-04',7),('2034-11-05',1),('2034-11-06',2),('2034-11-07',3),('2034-11-08',4),('2034-11-09',5),('2034-11-10',6),('2034-11-11',7),('2034-11-12',1),('2034-11-13',2),('2034-11-14',3),('2034-11-15',4),('2034-11-16',5),('2034-11-17',6),('2034-11-18',7),('2034-11-19',1),('2034-11-20',2),('2034-11-21',3),('2034-11-22',4),('2034-11-23',5),('2034-11-24',6),('2034-11-25',7),('2034-11-26',1),('2034-11-27',2),('2034-11-28',3),('2034-11-29',4),('2034-11-30',5),('2034-12-01',6),('2034-12-02',7),('2034-12-03',1),('2034-12-04',2),('2034-12-05',3),('2034-12-06',4),('2034-12-07',5),('2034-12-08',6),('2034-12-09',7),('2034-12-10',1),('2034-12-11',2),('2034-12-12',3),('2034-12-13',4),('2034-12-14',5),('2034-12-15',6),('2034-12-16',7),('2034-12-17',1),('2034-12-18',2),('2034-12-19',3),('2034-12-20',4),('2034-12-21',5),('2034-12-22',6),('2034-12-23',7),('2034-12-24',1),('2034-12-25',2),('2034-12-26',3),('2034-12-27',4),('2034-12-28',5),('2034-12-29',6),('2034-12-30',7),('2034-12-31',1),('2035-01-01',2),('2035-01-02',3),('2035-01-03',4),('2035-01-04',5),('2035-01-05',6),('2035-01-06',7),('2035-01-07',1),('2035-01-08',2),('2035-01-09',3),('2035-01-10',4),('2035-01-11',5),('2035-01-12',6),('2035-01-13',7),('2035-01-14',1),('2035-01-15',2),('2035-01-16',3),('2035-01-17',4),('2035-01-18',5),('2035-01-19',6),('2035-01-20',7),('2035-01-21',1),('2035-01-22',2),('2035-01-23',3),('2035-01-24',4),('2035-01-25',5),('2035-01-26',6),('2035-01-27',7),('2035-01-28',1),('2035-01-29',2),('2035-01-30',3),('2035-01-31',4),('2035-02-01',5),('2035-02-02',6),('2035-02-03',7),('2035-02-04',1),('2035-02-05',2),('2035-02-06',3),('2035-02-07',4),('2035-02-08',5),('2035-02-09',6),('2035-02-10',7),('2035-02-11',1),('2035-02-12',2),('2035-02-13',3),('2035-02-14',4),('2035-02-15',5),('2035-02-16',6),('2035-02-17',7),('2035-02-18',1),('2035-02-19',2),('2035-02-20',3),('2035-02-21',4),('2035-02-22',5),('2035-02-23',6),('2035-02-24',7),('2035-02-25',1),('2035-02-26',2),('2035-02-27',3),('2035-02-28',4),('2035-03-01',5),('2035-03-02',6),('2035-03-03',7),('2035-03-04',1),('2035-03-05',2),('2035-03-06',3),('2035-03-07',4),('2035-03-08',5),('2035-03-09',6),('2035-03-10',7),('2035-03-11',1),('2035-03-12',2),('2035-03-13',3),('2035-03-14',4),('2035-03-15',5),('2035-03-16',6),('2035-03-17',7),('2035-03-18',1),('2035-03-19',2),('2035-03-20',3),('2035-03-21',4),('2035-03-22',5),('2035-03-23',6),('2035-03-24',7),('2035-03-25',1),('2035-03-26',2),('2035-03-27',3),('2035-03-28',4),('2035-03-29',5),('2035-03-30',6),('2035-03-31',7),('2035-04-01',1),('2035-04-02',2),('2035-04-03',3),('2035-04-04',4),('2035-04-05',5),('2035-04-06',6),('2035-04-07',7),('2035-04-08',1),('2035-04-09',2),('2035-04-10',3),('2035-04-11',4),('2035-04-12',5),('2035-04-13',6),('2035-04-14',7),('2035-04-15',1),('2035-04-16',2),('2035-04-17',3),('2035-04-18',4),('2035-04-19',5),('2035-04-20',6),('2035-04-21',7),('2035-04-22',1),('2035-04-23',2),('2035-04-24',3),('2035-04-25',4),('2035-04-26',5),('2035-04-27',6),('2035-04-28',7),('2035-04-29',1),('2035-04-30',2),('2035-05-01',3),('2035-05-02',4),('2035-05-03',5),('2035-05-04',6),('2035-05-05',7),('2035-05-06',1),('2035-05-07',2),('2035-05-08',3),('2035-05-09',4),('2035-05-10',5),('2035-05-11',6),('2035-05-12',7),('2035-05-13',1),('2035-05-14',2),('2035-05-15',3),('2035-05-16',4),('2035-05-17',5),('2035-05-18',6),('2035-05-19',7),('2035-05-20',1),('2035-05-21',2),('2035-05-22',3),('2035-05-23',4),('2035-05-24',5),('2035-05-25',6),('2035-05-26',7),('2035-05-27',1),('2035-05-28',2),('2035-05-29',3),('2035-05-30',4),('2035-05-31',5),('2035-06-01',6),('2035-06-02',7),('2035-06-03',1),('2035-06-04',2),('2035-06-05',3),('2035-06-06',4),('2035-06-07',5),('2035-06-08',6),('2035-06-09',7),('2035-06-10',1),('2035-06-11',2),('2035-06-12',3),('2035-06-13',4),('2035-06-14',5),('2035-06-15',6),('2035-06-16',7),('2035-06-17',1),('2035-06-18',2),('2035-06-19',3),('2035-06-20',4),('2035-06-21',5),('2035-06-22',6),('2035-06-23',7),('2035-06-24',1),('2035-06-25',2),('2035-06-26',3),('2035-06-27',4),('2035-06-28',5),('2035-06-29',6),('2035-06-30',7),('2035-07-01',1),('2035-07-02',2),('2035-07-03',3),('2035-07-04',4),('2035-07-05',5),('2035-07-06',6),('2035-07-07',7),('2035-07-08',1),('2035-07-09',2),('2035-07-10',3),('2035-07-11',4),('2035-07-12',5),('2035-07-13',6),('2035-07-14',7),('2035-07-15',1),('2035-07-16',2),('2035-07-17',3),('2035-07-18',4),('2035-07-19',5),('2035-07-20',6),('2035-07-21',7),('2035-07-22',1),('2035-07-23',2),('2035-07-24',3),('2035-07-25',4),('2035-07-26',5),('2035-07-27',6),('2035-07-28',7),('2035-07-29',1),('2035-07-30',2),('2035-07-31',3),('2035-08-01',4),('2035-08-02',5),('2035-08-03',6),('2035-08-04',7),('2035-08-05',1),('2035-08-06',2),('2035-08-07',3),('2035-08-08',4),('2035-08-09',5),('2035-08-10',6),('2035-08-11',7),('2035-08-12',1),('2035-08-13',2),('2035-08-14',3),('2035-08-15',4),('2035-08-16',5),('2035-08-17',6),('2035-08-18',7),('2035-08-19',1),('2035-08-20',2),('2035-08-21',3),('2035-08-22',4),('2035-08-23',5),('2035-08-24',6),('2035-08-25',7),('2035-08-26',1),('2035-08-27',2),('2035-08-28',3),('2035-08-29',4),('2035-08-30',5),('2035-08-31',6),('2035-09-01',7),('2035-09-02',1),('2035-09-03',2),('2035-09-04',3),('2035-09-05',4),('2035-09-06',5),('2035-09-07',6),('2035-09-08',7),('2035-09-09',1),('2035-09-10',2),('2035-09-11',3),('2035-09-12',4),('2035-09-13',5),('2035-09-14',6),('2035-09-15',7),('2035-09-16',1),('2035-09-17',2),('2035-09-18',3),('2035-09-19',4),('2035-09-20',5),('2035-09-21',6),('2035-09-22',7),('2035-09-23',1),('2035-09-24',2),('2035-09-25',3),('2035-09-26',4),('2035-09-27',5),('2035-09-28',6),('2035-09-29',7),('2035-09-30',1),('2035-10-01',2),('2035-10-02',3),('2035-10-03',4),('2035-10-04',5),('2035-10-05',6),('2035-10-06',7),('2035-10-07',1),('2035-10-08',2),('2035-10-09',3),('2035-10-10',4),('2035-10-11',5),('2035-10-12',6),('2035-10-13',7),('2035-10-14',1),('2035-10-15',2),('2035-10-16',3),('2035-10-17',4),('2035-10-18',5),('2035-10-19',6),('2035-10-20',7),('2035-10-21',1),('2035-10-22',2),('2035-10-23',3),('2035-10-24',4),('2035-10-25',5),('2035-10-26',6),('2035-10-27',7),('2035-10-28',1),('2035-10-29',2),('2035-10-30',3),('2035-10-31',4),('2035-11-01',5),('2035-11-02',6),('2035-11-03',7),('2035-11-04',1),('2035-11-05',2),('2035-11-06',3),('2035-11-07',4),('2035-11-08',5),('2035-11-09',6),('2035-11-10',7),('2035-11-11',1),('2035-11-12',2),('2035-11-13',3),('2035-11-14',4),('2035-11-15',5),('2035-11-16',6),('2035-11-17',7),('2035-11-18',1),('2035-11-19',2),('2035-11-20',3),('2035-11-21',4),('2035-11-22',5),('2035-11-23',6),('2035-11-24',7),('2035-11-25',1),('2035-11-26',2),('2035-11-27',3),('2035-11-28',4),('2035-11-29',5),('2035-11-30',6),('2035-12-01',7),('2035-12-02',1),('2035-12-03',2),('2035-12-04',3),('2035-12-05',4),('2035-12-06',5),('2035-12-07',6),('2035-12-08',7),('2035-12-09',1),('2035-12-10',2),('2035-12-11',3),('2035-12-12',4),('2035-12-13',5),('2035-12-14',6),('2035-12-15',7),('2035-12-16',1),('2035-12-17',2),('2035-12-18',3),('2035-12-19',4),('2035-12-20',5),('2035-12-21',6),('2035-12-22',7),('2035-12-23',1),('2035-12-24',2),('2035-12-25',3),('2035-12-26',4),('2035-12-27',5),('2035-12-28',6),('2035-12-29',7),('2035-12-30',1),('2035-12-31',2),('2036-01-01',3),('2036-01-02',4),('2036-01-03',5),('2036-01-04',6),('2036-01-05',7),('2036-01-06',1),('2036-01-07',2),('2036-01-08',3),('2036-01-09',4),('2036-01-10',5),('2036-01-11',6),('2036-01-12',7),('2036-01-13',1),('2036-01-14',2),('2036-01-15',3),('2036-01-16',4),('2036-01-17',5),('2036-01-18',6),('2036-01-19',7),('2036-01-20',1),('2036-01-21',2),('2036-01-22',3),('2036-01-23',4),('2036-01-24',5),('2036-01-25',6),('2036-01-26',7),('2036-01-27',1),('2036-01-28',2),('2036-01-29',3),('2036-01-30',4),('2036-01-31',5),('2036-02-01',6),('2036-02-02',7),('2036-02-03',1),('2036-02-04',2),('2036-02-05',3),('2036-02-06',4),('2036-02-07',5),('2036-02-08',6),('2036-02-09',7),('2036-02-10',1),('2036-02-11',2),('2036-02-12',3),('2036-02-13',4),('2036-02-14',5),('2036-02-15',6),('2036-02-16',7),('2036-02-17',1),('2036-02-18',2),('2036-02-19',3),('2036-02-20',4),('2036-02-21',5),('2036-02-22',6),('2036-02-23',7),('2036-02-24',1),('2036-02-25',2),('2036-02-26',3),('2036-02-27',4),('2036-02-28',5),('2036-02-29',6),('2036-03-01',7),('2036-03-02',1),('2036-03-03',2),('2036-03-04',3),('2036-03-05',4),('2036-03-06',5),('2036-03-07',6),('2036-03-08',7),('2036-03-09',1),('2036-03-10',2),('2036-03-11',3),('2036-03-12',4),('2036-03-13',5),('2036-03-14',6),('2036-03-15',7),('2036-03-16',1),('2036-03-17',2),('2036-03-18',3),('2036-03-19',4),('2036-03-20',5),('2036-03-21',6),('2036-03-22',7),('2036-03-23',1),('2036-03-24',2),('2036-03-25',3),('2036-03-26',4),('2036-03-27',5),('2036-03-28',6),('2036-03-29',7),('2036-03-30',1),('2036-03-31',2),('2036-04-01',3),('2036-04-02',4),('2036-04-03',5),('2036-04-04',6),('2036-04-05',7),('2036-04-06',1),('2036-04-07',2),('2036-04-08',3),('2036-04-09',4),('2036-04-10',5),('2036-04-11',6),('2036-04-12',7),('2036-04-13',1),('2036-04-14',2),('2036-04-15',3),('2036-04-16',4),('2036-04-17',5),('2036-04-18',6),('2036-04-19',7),('2036-04-20',1),('2036-04-21',2),('2036-04-22',3),('2036-04-23',4),('2036-04-24',5),('2036-04-25',6),('2036-04-26',7),('2036-04-27',1),('2036-04-28',2),('2036-04-29',3),('2036-04-30',4),('2036-05-01',5),('2036-05-02',6),('2036-05-03',7),('2036-05-04',1),('2036-05-05',2),('2036-05-06',3),('2036-05-07',4),('2036-05-08',5),('2036-05-09',6),('2036-05-10',7),('2036-05-11',1),('2036-05-12',2),('2036-05-13',3),('2036-05-14',4),('2036-05-15',5),('2036-05-16',6),('2036-05-17',7),('2036-05-18',1),('2036-05-19',2),('2036-05-20',3),('2036-05-21',4),('2036-05-22',5),('2036-05-23',6),('2036-05-24',7),('2036-05-25',1),('2036-05-26',2),('2036-05-27',3),('2036-05-28',4),('2036-05-29',5),('2036-05-30',6),('2036-05-31',7),('2036-06-01',1),('2036-06-02',2),('2036-06-03',3),('2036-06-04',4),('2036-06-05',5),('2036-06-06',6),('2036-06-07',7),('2036-06-08',1),('2036-06-09',2),('2036-06-10',3),('2036-06-11',4),('2036-06-12',5),('2036-06-13',6),('2036-06-14',7),('2036-06-15',1),('2036-06-16',2),('2036-06-17',3),('2036-06-18',4),('2036-06-19',5),('2036-06-20',6),('2036-06-21',7),('2036-06-22',1),('2036-06-23',2),('2036-06-24',3),('2036-06-25',4),('2036-06-26',5),('2036-06-27',6),('2036-06-28',7),('2036-06-29',1),('2036-06-30',2),('2036-07-01',3),('2036-07-02',4),('2036-07-03',5),('2036-07-04',6),('2036-07-05',7),('2036-07-06',1),('2036-07-07',2),('2036-07-08',3),('2036-07-09',4),('2036-07-10',5),('2036-07-11',6),('2036-07-12',7),('2036-07-13',1),('2036-07-14',2),('2036-07-15',3),('2036-07-16',4),('2036-07-17',5),('2036-07-18',6),('2036-07-19',7),('2036-07-20',1),('2036-07-21',2),('2036-07-22',3),('2036-07-23',4),('2036-07-24',5),('2036-07-25',6),('2036-07-26',7),('2036-07-27',1),('2036-07-28',2),('2036-07-29',3),('2036-07-30',4),('2036-07-31',5),('2036-08-01',6),('2036-08-02',7),('2036-08-03',1),('2036-08-04',2),('2036-08-05',3),('2036-08-06',4),('2036-08-07',5),('2036-08-08',6),('2036-08-09',7),('2036-08-10',1),('2036-08-11',2),('2036-08-12',3),('2036-08-13',4),('2036-08-14',5),('2036-08-15',6),('2036-08-16',7),('2036-08-17',1),('2036-08-18',2),('2036-08-19',3),('2036-08-20',4),('2036-08-21',5),('2036-08-22',6),('2036-08-23',7),('2036-08-24',1),('2036-08-25',2),('2036-08-26',3),('2036-08-27',4),('2036-08-28',5),('2036-08-29',6),('2036-08-30',7),('2036-08-31',1),('2036-09-01',2),('2036-09-02',3),('2036-09-03',4),('2036-09-04',5),('2036-09-05',6),('2036-09-06',7),('2036-09-07',1),('2036-09-08',2),('2036-09-09',3),('2036-09-10',4),('2036-09-11',5),('2036-09-12',6),('2036-09-13',7),('2036-09-14',1),('2036-09-15',2),('2036-09-16',3),('2036-09-17',4),('2036-09-18',5),('2036-09-19',6),('2036-09-20',7),('2036-09-21',1),('2036-09-22',2),('2036-09-23',3),('2036-09-24',4),('2036-09-25',5),('2036-09-26',6),('2036-09-27',7),('2036-09-28',1),('2036-09-29',2),('2036-09-30',3),('2036-10-01',4),('2036-10-02',5),('2036-10-03',6),('2036-10-04',7),('2036-10-05',1),('2036-10-06',2),('2036-10-07',3),('2036-10-08',4),('2036-10-09',5),('2036-10-10',6),('2036-10-11',7),('2036-10-12',1),('2036-10-13',2),('2036-10-14',3),('2036-10-15',4),('2036-10-16',5),('2036-10-17',6),('2036-10-18',7),('2036-10-19',1),('2036-10-20',2),('2036-10-21',3),('2036-10-22',4),('2036-10-23',5),('2036-10-24',6),('2036-10-25',7),('2036-10-26',1),('2036-10-27',2),('2036-10-28',3),('2036-10-29',4),('2036-10-30',5),('2036-10-31',6),('2036-11-01',7),('2036-11-02',1),('2036-11-03',2),('2036-11-04',3),('2036-11-05',4),('2036-11-06',5),('2036-11-07',6),('2036-11-08',7),('2036-11-09',1),('2036-11-10',2),('2036-11-11',3),('2036-11-12',4),('2036-11-13',5),('2036-11-14',6),('2036-11-15',7),('2036-11-16',1),('2036-11-17',2),('2036-11-18',3),('2036-11-19',4),('2036-11-20',5),('2036-11-21',6),('2036-11-22',7),('2036-11-23',1),('2036-11-24',2),('2036-11-25',3),('2036-11-26',4),('2036-11-27',5),('2036-11-28',6),('2036-11-29',7),('2036-11-30',1),('2036-12-01',2),('2036-12-02',3),('2036-12-03',4),('2036-12-04',5),('2036-12-05',6),('2036-12-06',7),('2036-12-07',1),('2036-12-08',2),('2036-12-09',3),('2036-12-10',4),('2036-12-11',5),('2036-12-12',6),('2036-12-13',7),('2036-12-14',1),('2036-12-15',2),('2036-12-16',3),('2036-12-17',4),('2036-12-18',5),('2036-12-19',6),('2036-12-20',7),('2036-12-21',1),('2036-12-22',2),('2036-12-23',3),('2036-12-24',4),('2036-12-25',5),('2036-12-26',6),('2036-12-27',7),('2036-12-28',1),('2036-12-29',2),('2036-12-30',3),('2036-12-31',4),('2037-01-01',5),('2037-01-02',6),('2037-01-03',7),('2037-01-04',1),('2037-01-05',2),('2037-01-06',3),('2037-01-07',4),('2037-01-08',5),('2037-01-09',6),('2037-01-10',7),('2037-01-11',1),('2037-01-12',2),('2037-01-13',3),('2037-01-14',4),('2037-01-15',5),('2037-01-16',6),('2037-01-17',7),('2037-01-18',1),('2037-01-19',2),('2037-01-20',3),('2037-01-21',4),('2037-01-22',5),('2037-01-23',6),('2037-01-24',7),('2037-01-25',1),('2037-01-26',2),('2037-01-27',3),('2037-01-28',4),('2037-01-29',5),('2037-01-30',6),('2037-01-31',7),('2037-02-01',1),('2037-02-02',2),('2037-02-03',3),('2037-02-04',4),('2037-02-05',5),('2037-02-06',6),('2037-02-07',7),('2037-02-08',1),('2037-02-09',2),('2037-02-10',3),('2037-02-11',4),('2037-02-12',5),('2037-02-13',6),('2037-02-14',7),('2037-02-15',1),('2037-02-16',2),('2037-02-17',3),('2037-02-18',4),('2037-02-19',5),('2037-02-20',6),('2037-02-21',7),('2037-02-22',1),('2037-02-23',2),('2037-02-24',3),('2037-02-25',4),('2037-02-26',5),('2037-02-27',6),('2037-02-28',7),('2037-03-01',1),('2037-03-02',2),('2037-03-03',3),('2037-03-04',4),('2037-03-05',5),('2037-03-06',6),('2037-03-07',7),('2037-03-08',1),('2037-03-09',2),('2037-03-10',3),('2037-03-11',4),('2037-03-12',5),('2037-03-13',6),('2037-03-14',7),('2037-03-15',1),('2037-03-16',2),('2037-03-17',3),('2037-03-18',4),('2037-03-19',5),('2037-03-20',6),('2037-03-21',7),('2037-03-22',1),('2037-03-23',2),('2037-03-24',3),('2037-03-25',4),('2037-03-26',5),('2037-03-27',6),('2037-03-28',7),('2037-03-29',1),('2037-03-30',2),('2037-03-31',3),('2037-04-01',4),('2037-04-02',5),('2037-04-03',6),('2037-04-04',7),('2037-04-05',1),('2037-04-06',2),('2037-04-07',3),('2037-04-08',4),('2037-04-09',5),('2037-04-10',6),('2037-04-11',7),('2037-04-12',1),('2037-04-13',2),('2037-04-14',3),('2037-04-15',4),('2037-04-16',5),('2037-04-17',6),('2037-04-18',7),('2037-04-19',1),('2037-04-20',2),('2037-04-21',3),('2037-04-22',4),('2037-04-23',5),('2037-04-24',6),('2037-04-25',7),('2037-04-26',1),('2037-04-27',2),('2037-04-28',3),('2037-04-29',4),('2037-04-30',5),('2037-05-01',6),('2037-05-02',7),('2037-05-03',1),('2037-05-04',2),('2037-05-05',3),('2037-05-06',4),('2037-05-07',5),('2037-05-08',6),('2037-05-09',7),('2037-05-10',1),('2037-05-11',2),('2037-05-12',3),('2037-05-13',4),('2037-05-14',5),('2037-05-15',6),('2037-05-16',7),('2037-05-17',1),('2037-05-18',2),('2037-05-19',3),('2037-05-20',4),('2037-05-21',5),('2037-05-22',6),('2037-05-23',7),('2037-05-24',1),('2037-05-25',2),('2037-05-26',3),('2037-05-27',4),('2037-05-28',5),('2037-05-29',6),('2037-05-30',7),('2037-05-31',1),('2037-06-01',2),('2037-06-02',3),('2037-06-03',4),('2037-06-04',5),('2037-06-05',6),('2037-06-06',7),('2037-06-07',1),('2037-06-08',2),('2037-06-09',3),('2037-06-10',4),('2037-06-11',5),('2037-06-12',6),('2037-06-13',7),('2037-06-14',1),('2037-06-15',2),('2037-06-16',3),('2037-06-17',4),('2037-06-18',5),('2037-06-19',6),('2037-06-20',7),('2037-06-21',1),('2037-06-22',2),('2037-06-23',3),('2037-06-24',4),('2037-06-25',5),('2037-06-26',6),('2037-06-27',7),('2037-06-28',1),('2037-06-29',2),('2037-06-30',3),('2037-07-01',4),('2037-07-02',5),('2037-07-03',6),('2037-07-04',7),('2037-07-05',1),('2037-07-06',2),('2037-07-07',3),('2037-07-08',4),('2037-07-09',5),('2037-07-10',6),('2037-07-11',7),('2037-07-12',1),('2037-07-13',2),('2037-07-14',3),('2037-07-15',4),('2037-07-16',5),('2037-07-17',6),('2037-07-18',7),('2037-07-19',1),('2037-07-20',2),('2037-07-21',3),('2037-07-22',4),('2037-07-23',5),('2037-07-24',6),('2037-07-25',7),('2037-07-26',1),('2037-07-27',2),('2037-07-28',3),('2037-07-29',4),('2037-07-30',5),('2037-07-31',6),('2037-08-01',7),('2037-08-02',1),('2037-08-03',2),('2037-08-04',3),('2037-08-05',4),('2037-08-06',5),('2037-08-07',6),('2037-08-08',7),('2037-08-09',1),('2037-08-10',2),('2037-08-11',3),('2037-08-12',4),('2037-08-13',5),('2037-08-14',6),('2037-08-15',7),('2037-08-16',1),('2037-08-17',2),('2037-08-18',3),('2037-08-19',4),('2037-08-20',5),('2037-08-21',6),('2037-08-22',7),('2037-08-23',1),('2037-08-24',2),('2037-08-25',3),('2037-08-26',4),('2037-08-27',5),('2037-08-28',6),('2037-08-29',7),('2037-08-30',1),('2037-08-31',2),('2037-09-01',3),('2037-09-02',4),('2037-09-03',5),('2037-09-04',6),('2037-09-05',7),('2037-09-06',1),('2037-09-07',2),('2037-09-08',3),('2037-09-09',4),('2037-09-10',5),('2037-09-11',6),('2037-09-12',7),('2037-09-13',1),('2037-09-14',2),('2037-09-15',3),('2037-09-16',4),('2037-09-17',5),('2037-09-18',6),('2037-09-19',7),('2037-09-20',1),('2037-09-21',2),('2037-09-22',3),('2037-09-23',4),('2037-09-24',5),('2037-09-25',6),('2037-09-26',7),('2037-09-27',1),('2037-09-28',2),('2037-09-29',3),('2037-09-30',4),('2037-10-01',5),('2037-10-02',6),('2037-10-03',7),('2037-10-04',1),('2037-10-05',2),('2037-10-06',3),('2037-10-07',4),('2037-10-08',5),('2037-10-09',6),('2037-10-10',7),('2037-10-11',1),('2037-10-12',2),('2037-10-13',3),('2037-10-14',4),('2037-10-15',5),('2037-10-16',6),('2037-10-17',7),('2037-10-18',1),('2037-10-19',2),('2037-10-20',3),('2037-10-21',4),('2037-10-22',5),('2037-10-23',6),('2037-10-24',7),('2037-10-25',1),('2037-10-26',2),('2037-10-27',3),('2037-10-28',4),('2037-10-29',5),('2037-10-30',6),('2037-10-31',7),('2037-11-01',1),('2037-11-02',2),('2037-11-03',3),('2037-11-04',4),('2037-11-05',5),('2037-11-06',6),('2037-11-07',7),('2037-11-08',1),('2037-11-09',2),('2037-11-10',3),('2037-11-11',4),('2037-11-12',5),('2037-11-13',6),('2037-11-14',7),('2037-11-15',1),('2037-11-16',2),('2037-11-17',3),('2037-11-18',4),('2037-11-19',5),('2037-11-20',6),('2037-11-21',7),('2037-11-22',1),('2037-11-23',2),('2037-11-24',3),('2037-11-25',4),('2037-11-26',5),('2037-11-27',6),('2037-11-28',7),('2037-11-29',1),('2037-11-30',2),('2037-12-01',3),('2037-12-02',4),('2037-12-03',5),('2037-12-04',6),('2037-12-05',7),('2037-12-06',1),('2037-12-07',2),('2037-12-08',3),('2037-12-09',4),('2037-12-10',5),('2037-12-11',6),('2037-12-12',7),('2037-12-13',1),('2037-12-14',2),('2037-12-15',3),('2037-12-16',4),('2037-12-17',5),('2037-12-18',6),('2037-12-19',7),('2037-12-20',1),('2037-12-21',2),('2037-12-22',3),('2037-12-23',4),('2037-12-24',5),('2037-12-25',6),('2037-12-26',7),('2037-12-27',1),('2037-12-28',2),('2037-12-29',3),('2037-12-30',4),('2037-12-31',5),('2038-01-01',6),('2038-01-02',7),('2038-01-03',1),('2038-01-04',2),('2038-01-05',3),('2038-01-06',4),('2038-01-07',5),('2038-01-08',6),('2038-01-09',7),('2038-01-10',1),('2038-01-11',2),('2038-01-12',3),('2038-01-13',4),('2038-01-14',5),('2038-01-15',6),('2038-01-16',7),('2038-01-17',1),('2038-01-18',2),('2038-01-19',3),('2038-01-20',4),('2038-01-21',5),('2038-01-22',6),('2038-01-23',7),('2038-01-24',1),('2038-01-25',2),('2038-01-26',3),('2038-01-27',4),('2038-01-28',5),('2038-01-29',6),('2038-01-30',7),('2038-01-31',1),('2038-02-01',2),('2038-02-02',3),('2038-02-03',4),('2038-02-04',5),('2038-02-05',6),('2038-02-06',7),('2038-02-07',1),('2038-02-08',2),('2038-02-09',3),('2038-02-10',4),('2038-02-11',5),('2038-02-12',6),('2038-02-13',7),('2038-02-14',1),('2038-02-15',2),('2038-02-16',3),('2038-02-17',4),('2038-02-18',5),('2038-02-19',6),('2038-02-20',7),('2038-02-21',1),('2038-02-22',2),('2038-02-23',3),('2038-02-24',4),('2038-02-25',5),('2038-02-26',6),('2038-02-27',7),('2038-02-28',1),('2038-03-01',2),('2038-03-02',3),('2038-03-03',4),('2038-03-04',5),('2038-03-05',6),('2038-03-06',7),('2038-03-07',1),('2038-03-08',2),('2038-03-09',3),('2038-03-10',4),('2038-03-11',5),('2038-03-12',6),('2038-03-13',7),('2038-03-14',1),('2038-03-15',2),('2038-03-16',3),('2038-03-17',4),('2038-03-18',5),('2038-03-19',6),('2038-03-20',7),('2038-03-21',1),('2038-03-22',2),('2038-03-23',3),('2038-03-24',4),('2038-03-25',5),('2038-03-26',6),('2038-03-27',7),('2038-03-28',1),('2038-03-29',2),('2038-03-30',3),('2038-03-31',4),('2038-04-01',5),('2038-04-02',6),('2038-04-03',7),('2038-04-04',1),('2038-04-05',2),('2038-04-06',3),('2038-04-07',4),('2038-04-08',5),('2038-04-09',6),('2038-04-10',7),('2038-04-11',1),('2038-04-12',2),('2038-04-13',3),('2038-04-14',4),('2038-04-15',5),('2038-04-16',6),('2038-04-17',7),('2038-04-18',1),('2038-04-19',2),('2038-04-20',3),('2038-04-21',4),('2038-04-22',5),('2038-04-23',6),('2038-04-24',7),('2038-04-25',1),('2038-04-26',2),('2038-04-27',3),('2038-04-28',4),('2038-04-29',5),('2038-04-30',6),('2038-05-01',7),('2038-05-02',1),('2038-05-03',2),('2038-05-04',3),('2038-05-05',4),('2038-05-06',5),('2038-05-07',6),('2038-05-08',7),('2038-05-09',1),('2038-05-10',2),('2038-05-11',3),('2038-05-12',4),('2038-05-13',5),('2038-05-14',6),('2038-05-15',7),('2038-05-16',1),('2038-05-17',2),('2038-05-18',3),('2038-05-19',4),('2038-05-20',5),('2038-05-21',6),('2038-05-22',7),('2038-05-23',1),('2038-05-24',2),('2038-05-25',3),('2038-05-26',4),('2038-05-27',5),('2038-05-28',6),('2038-05-29',7),('2038-05-30',1),('2038-05-31',2),('2038-06-01',3),('2038-06-02',4),('2038-06-03',5),('2038-06-04',6),('2038-06-05',7),('2038-06-06',1),('2038-06-07',2),('2038-06-08',3),('2038-06-09',4),('2038-06-10',5),('2038-06-11',6),('2038-06-12',7),('2038-06-13',1),('2038-06-14',2),('2038-06-15',3),('2038-06-16',4),('2038-06-17',5),('2038-06-18',6),('2038-06-19',7),('2038-06-20',1),('2038-06-21',2),('2038-06-22',3),('2038-06-23',4),('2038-06-24',5),('2038-06-25',6),('2038-06-26',7),('2038-06-27',1),('2038-06-28',2),('2038-06-29',3),('2038-06-30',4),('2038-07-01',5),('2038-07-02',6),('2038-07-03',7),('2038-07-04',1),('2038-07-05',2),('2038-07-06',3),('2038-07-07',4),('2038-07-08',5),('2038-07-09',6),('2038-07-10',7),('2038-07-11',1),('2038-07-12',2),('2038-07-13',3),('2038-07-14',4),('2038-07-15',5),('2038-07-16',6),('2038-07-17',7),('2038-07-18',1),('2038-07-19',2),('2038-07-20',3),('2038-07-21',4),('2038-07-22',5),('2038-07-23',6),('2038-07-24',7),('2038-07-25',1),('2038-07-26',2),('2038-07-27',3),('2038-07-28',4),('2038-07-29',5),('2038-07-30',6),('2038-07-31',7),('2038-08-01',1),('2038-08-02',2),('2038-08-03',3),('2038-08-04',4),('2038-08-05',5),('2038-08-06',6),('2038-08-07',7),('2038-08-08',1),('2038-08-09',2),('2038-08-10',3),('2038-08-11',4),('2038-08-12',5),('2038-08-13',6),('2038-08-14',7),('2038-08-15',1),('2038-08-16',2),('2038-08-17',3),('2038-08-18',4),('2038-08-19',5),('2038-08-20',6),('2038-08-21',7),('2038-08-22',1),('2038-08-23',2),('2038-08-24',3),('2038-08-25',4),('2038-08-26',5),('2038-08-27',6),('2038-08-28',7),('2038-08-29',1),('2038-08-30',2),('2038-08-31',3),('2038-09-01',4),('2038-09-02',5),('2038-09-03',6),('2038-09-04',7),('2038-09-05',1),('2038-09-06',2),('2038-09-07',3),('2038-09-08',4),('2038-09-09',5),('2038-09-10',6),('2038-09-11',7),('2038-09-12',1),('2038-09-13',2),('2038-09-14',3),('2038-09-15',4),('2038-09-16',5),('2038-09-17',6),('2038-09-18',7),('2038-09-19',1),('2038-09-20',2),('2038-09-21',3),('2038-09-22',4),('2038-09-23',5),('2038-09-24',6),('2038-09-25',7),('2038-09-26',1),('2038-09-27',2),('2038-09-28',3),('2038-09-29',4),('2038-09-30',5),('2038-10-01',6),('2038-10-02',7),('2038-10-03',1),('2038-10-04',2),('2038-10-05',3),('2038-10-06',4),('2038-10-07',5),('2038-10-08',6),('2038-10-09',7),('2038-10-10',1),('2038-10-11',2),('2038-10-12',3),('2038-10-13',4),('2038-10-14',5),('2038-10-15',6),('2038-10-16',7),('2038-10-17',1),('2038-10-18',2),('2038-10-19',3),('2038-10-20',4),('2038-10-21',5),('2038-10-22',6),('2038-10-23',7),('2038-10-24',1),('2038-10-25',2),('2038-10-26',3),('2038-10-27',4),('2038-10-28',5),('2038-10-29',6),('2038-10-30',7),('2038-10-31',1),('2038-11-01',2),('2038-11-02',3),('2038-11-03',4),('2038-11-04',5),('2038-11-05',6),('2038-11-06',7),('2038-11-07',1),('2038-11-08',2),('2038-11-09',3),('2038-11-10',4),('2038-11-11',5),('2038-11-12',6),('2038-11-13',7),('2038-11-14',1),('2038-11-15',2),('2038-11-16',3),('2038-11-17',4),('2038-11-18',5),('2038-11-19',6),('2038-11-20',7),('2038-11-21',1),('2038-11-22',2),('2038-11-23',3),('2038-11-24',4),('2038-11-25',5),('2038-11-26',6),('2038-11-27',7),('2038-11-28',1),('2038-11-29',2),('2038-11-30',3),('2038-12-01',4),('2038-12-02',5),('2038-12-03',6),('2038-12-04',7),('2038-12-05',1),('2038-12-06',2),('2038-12-07',3),('2038-12-08',4),('2038-12-09',5),('2038-12-10',6),('2038-12-11',7),('2038-12-12',1),('2038-12-13',2),('2038-12-14',3),('2038-12-15',4),('2038-12-16',5),('2038-12-17',6),('2038-12-18',7),('2038-12-19',1),('2038-12-20',2),('2038-12-21',3),('2038-12-22',4),('2038-12-23',5),('2038-12-24',6),('2038-12-25',7),('2038-12-26',1),('2038-12-27',2),('2038-12-28',3),('2038-12-29',4),('2038-12-30',5),('2038-12-31',6),('2039-01-01',7),('2039-01-02',1),('2039-01-03',2),('2039-01-04',3),('2039-01-05',4),('2039-01-06',5),('2039-01-07',6),('2039-01-08',7),('2039-01-09',1),('2039-01-10',2),('2039-01-11',3),('2039-01-12',4),('2039-01-13',5),('2039-01-14',6),('2039-01-15',7),('2039-01-16',1),('2039-01-17',2),('2039-01-18',3),('2039-01-19',4),('2039-01-20',5),('2039-01-21',6),('2039-01-22',7),('2039-01-23',1),('2039-01-24',2),('2039-01-25',3),('2039-01-26',4),('2039-01-27',5),('2039-01-28',6),('2039-01-29',7),('2039-01-30',1),('2039-01-31',2),('2039-02-01',3),('2039-02-02',4),('2039-02-03',5),('2039-02-04',6),('2039-02-05',7),('2039-02-06',1),('2039-02-07',2),('2039-02-08',3),('2039-02-09',4),('2039-02-10',5),('2039-02-11',6),('2039-02-12',7),('2039-02-13',1),('2039-02-14',2),('2039-02-15',3),('2039-02-16',4),('2039-02-17',5),('2039-02-18',6),('2039-02-19',7),('2039-02-20',1),('2039-02-21',2),('2039-02-22',3),('2039-02-23',4),('2039-02-24',5),('2039-02-25',6),('2039-02-26',7),('2039-02-27',1),('2039-02-28',2),('2039-03-01',3),('2039-03-02',4),('2039-03-03',5),('2039-03-04',6),('2039-03-05',7),('2039-03-06',1),('2039-03-07',2),('2039-03-08',3),('2039-03-09',4),('2039-03-10',5),('2039-03-11',6),('2039-03-12',7),('2039-03-13',1),('2039-03-14',2),('2039-03-15',3),('2039-03-16',4),('2039-03-17',5),('2039-03-18',6),('2039-03-19',7),('2039-03-20',1),('2039-03-21',2),('2039-03-22',3),('2039-03-23',4),('2039-03-24',5),('2039-03-25',6),('2039-03-26',7),('2039-03-27',1),('2039-03-28',2),('2039-03-29',3),('2039-03-30',4),('2039-03-31',5),('2039-04-01',6),('2039-04-02',7),('2039-04-03',1),('2039-04-04',2),('2039-04-05',3),('2039-04-06',4),('2039-04-07',5),('2039-04-08',6),('2039-04-09',7),('2039-04-10',1),('2039-04-11',2),('2039-04-12',3),('2039-04-13',4),('2039-04-14',5),('2039-04-15',6),('2039-04-16',7),('2039-04-17',1),('2039-04-18',2),('2039-04-19',3),('2039-04-20',4),('2039-04-21',5),('2039-04-22',6),('2039-04-23',7),('2039-04-24',1),('2039-04-25',2),('2039-04-26',3),('2039-04-27',4),('2039-04-28',5),('2039-04-29',6),('2039-04-30',7),('2039-05-01',1),('2039-05-02',2),('2039-05-03',3),('2039-05-04',4),('2039-05-05',5),('2039-05-06',6),('2039-05-07',7),('2039-05-08',1),('2039-05-09',2),('2039-05-10',3),('2039-05-11',4),('2039-05-12',5),('2039-05-13',6),('2039-05-14',7),('2039-05-15',1),('2039-05-16',2),('2039-05-17',3),('2039-05-18',4),('2039-05-19',5),('2039-05-20',6),('2039-05-21',7),('2039-05-22',1),('2039-05-23',2),('2039-05-24',3),('2039-05-25',4),('2039-05-26',5),('2039-05-27',6),('2039-05-28',7),('2039-05-29',1),('2039-05-30',2),('2039-05-31',3),('2039-06-01',4),('2039-06-02',5),('2039-06-03',6),('2039-06-04',7),('2039-06-05',1),('2039-06-06',2),('2039-06-07',3),('2039-06-08',4),('2039-06-09',5),('2039-06-10',6),('2039-06-11',7),('2039-06-12',1),('2039-06-13',2),('2039-06-14',3),('2039-06-15',4),('2039-06-16',5),('2039-06-17',6),('2039-06-18',7),('2039-06-19',1),('2039-06-20',2),('2039-06-21',3),('2039-06-22',4),('2039-06-23',5),('2039-06-24',6),('2039-06-25',7),('2039-06-26',1),('2039-06-27',2),('2039-06-28',3),('2039-06-29',4),('2039-06-30',5),('2039-07-01',6),('2039-07-02',7),('2039-07-03',1),('2039-07-04',2),('2039-07-05',3),('2039-07-06',4),('2039-07-07',5),('2039-07-08',6),('2039-07-09',7),('2039-07-10',1),('2039-07-11',2),('2039-07-12',3),('2039-07-13',4),('2039-07-14',5),('2039-07-15',6),('2039-07-16',7),('2039-07-17',1),('2039-07-18',2),('2039-07-19',3),('2039-07-20',4),('2039-07-21',5),('2039-07-22',6),('2039-07-23',7),('2039-07-24',1),('2039-07-25',2),('2039-07-26',3),('2039-07-27',4),('2039-07-28',5),('2039-07-29',6),('2039-07-30',7),('2039-07-31',1),('2039-08-01',2),('2039-08-02',3),('2039-08-03',4),('2039-08-04',5),('2039-08-05',6),('2039-08-06',7),('2039-08-07',1),('2039-08-08',2),('2039-08-09',3),('2039-08-10',4),('2039-08-11',5),('2039-08-12',6),('2039-08-13',7),('2039-08-14',1),('2039-08-15',2),('2039-08-16',3),('2039-08-17',4),('2039-08-18',5),('2039-08-19',6),('2039-08-20',7),('2039-08-21',1),('2039-08-22',2),('2039-08-23',3),('2039-08-24',4),('2039-08-25',5),('2039-08-26',6),('2039-08-27',7),('2039-08-28',1),('2039-08-29',2),('2039-08-30',3),('2039-08-31',4),('2039-09-01',5),('2039-09-02',6),('2039-09-03',7),('2039-09-04',1),('2039-09-05',2),('2039-09-06',3),('2039-09-07',4),('2039-09-08',5),('2039-09-09',6),('2039-09-10',7),('2039-09-11',1),('2039-09-12',2),('2039-09-13',3),('2039-09-14',4),('2039-09-15',5),('2039-09-16',6),('2039-09-17',7),('2039-09-18',1),('2039-09-19',2),('2039-09-20',3),('2039-09-21',4),('2039-09-22',5),('2039-09-23',6),('2039-09-24',7),('2039-09-25',1),('2039-09-26',2),('2039-09-27',3),('2039-09-28',4),('2039-09-29',5),('2039-09-30',6),('2039-10-01',7),('2039-10-02',1),('2039-10-03',2),('2039-10-04',3),('2039-10-05',4),('2039-10-06',5),('2039-10-07',6),('2039-10-08',7),('2039-10-09',1),('2039-10-10',2),('2039-10-11',3),('2039-10-12',4),('2039-10-13',5),('2039-10-14',6),('2039-10-15',7),('2039-10-16',1),('2039-10-17',2),('2039-10-18',3),('2039-10-19',4),('2039-10-20',5),('2039-10-21',6),('2039-10-22',7),('2039-10-23',1),('2039-10-24',2),('2039-10-25',3),('2039-10-26',4),('2039-10-27',5),('2039-10-28',6),('2039-10-29',7),('2039-10-30',1),('2039-10-31',2),('2039-11-01',3),('2039-11-02',4),('2039-11-03',5),('2039-11-04',6),('2039-11-05',7),('2039-11-06',1),('2039-11-07',2),('2039-11-08',3),('2039-11-09',4),('2039-11-10',5),('2039-11-11',6),('2039-11-12',7),('2039-11-13',1),('2039-11-14',2),('2039-11-15',3),('2039-11-16',4),('2039-11-17',5),('2039-11-18',6),('2039-11-19',7),('2039-11-20',1),('2039-11-21',2),('2039-11-22',3),('2039-11-23',4),('2039-11-24',5),('2039-11-25',6),('2039-11-26',7),('2039-11-27',1),('2039-11-28',2),('2039-11-29',3),('2039-11-30',4),('2039-12-01',5),('2039-12-02',6),('2039-12-03',7),('2039-12-04',1),('2039-12-05',2),('2039-12-06',3),('2039-12-07',4),('2039-12-08',5),('2039-12-09',6),('2039-12-10',7),('2039-12-11',1),('2039-12-12',2),('2039-12-13',3),('2039-12-14',4),('2039-12-15',5),('2039-12-16',6),('2039-12-17',7),('2039-12-18',1),('2039-12-19',2),('2039-12-20',3),('2039-12-21',4),('2039-12-22',5),('2039-12-23',6),('2039-12-24',7),('2039-12-25',1),('2039-12-26',2),('2039-12-27',3),('2039-12-28',4),('2039-12-29',5),('2039-12-30',6),('2039-12-31',7),('2040-01-01',1),('2040-01-02',2),('2040-01-03',3),('2040-01-04',4),('2040-01-05',5),('2040-01-06',6),('2040-01-07',7),('2040-01-08',1),('2040-01-09',2),('2040-01-10',3),('2040-01-11',4),('2040-01-12',5),('2040-01-13',6),('2040-01-14',7);

/*Table structure for table `disponibilidad` */

DROP TABLE IF EXISTS `disponibilidad`;

CREATE TABLE `disponibilidad` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la linea',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `lunes` int(5) DEFAULT NULL COMMENT 'Disponibilidad lunes',
  `martes` int(5) DEFAULT NULL COMMENT 'Disponibilidad martes',
  `miercoles` int(5) DEFAULT NULL COMMENT 'Disponibilidad miércoles',
  `jueves` int(5) DEFAULT NULL COMMENT 'Disponibilidad jueves',
  `viernes` int(5) DEFAULT NULL COMMENT 'Disponibilidad viernes',
  `sabado` int(5) DEFAULT NULL COMMENT 'Disponibilidad sábado',
  `domingo` int(5) DEFAULT NULL COMMENT 'Disponibilidad domingo',
  `estatus` char(1) COLLATE utf8_bin DEFAULT 'A' COMMENT 'Estatus del registro',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=232 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de disponibilidad';

/*Data for the table `disponibilidad` */

/*Table structure for table `equipos_objetivo` */

DROP TABLE IF EXISTS `equipos_objetivo`;

CREATE TABLE `equipos_objetivo` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID de la ruta',
  `lote` bigint(20) DEFAULT '0' COMMENT 'ID de la Orden o lote',
  `fijo` char(1) DEFAULT 'S' COMMENT 'Objetivo fijo',
  `desde` date DEFAULT NULL COMMENT 'Fecha',
  `hasta` date DEFAULT NULL,
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `objetivo` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Piezas por segundo',
  `reinicio` int(2) DEFAULT '0' COMMENT 'Tipo de reinicio (0=No reiniciar, 1=Al llegar, 2=Cambio de producto, 3=Hora, 4=turno, 5=dia, 6=semana, 7=mes, 8=anual)',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`parte`,`equipo`,`fijo`)
) ENGINE=MyISAM AUTO_INCREMENT=4 DEFAULT CHARSET=latin1 COMMENT='Objetivos por equipo';

/*Data for the table `equipos_objetivo` */

/*Table structure for table `estimados` */

DROP TABLE IF EXISTS `estimados`;

CREATE TABLE `estimados` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la línea',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `equipo` varbinary(20) DEFAULT '0' COMMENT 'ID del equipo',
  `fijo` char(1) DEFAULT 'S',
  `desde` date DEFAULT NULL COMMENT 'Desde',
  `hasta` date DEFAULT NULL COMMENT 'Mes del año',
  `oee_minimo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Mínimo de OEE',
  `oee` decimal(10,5) DEFAULT '0.00000',
  `oee_maximo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Máximo de OEE',
  `ftq_minimo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Mínimo de FTQ',
  `ftq` decimal(10,5) DEFAULT '0.00000',
  `ftq_maximo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Máximo de FTQ',
  `efi_minimo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Mínimo de Eficiencia',
  `efi` decimal(10,5) DEFAULT '0.00000',
  `efi_maximo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Máximo de Eficiencia',
  `dis_minimo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Mínimo de Disponibilidad',
  `dis` decimal(10,5) DEFAULT '0.00000',
  `dis_maximo` decimal(10,5) DEFAULT '0.00000' COMMENT 'Máximo de Disponibilidad',
  PRIMARY KEY (`id`),
  UNIQUE KEY `NewIndex1` (`linea`,`desde`)
) ENGINE=MyISAM AUTO_INCREMENT=3 DEFAULT CHARSET=latin1 COMMENT='Estimado de cumplimiento por semana';

/*Data for the table `estimados` */

/*Table structure for table `figuras` */

DROP TABLE IF EXISTS `figuras`;

CREATE TABLE `figuras` (
  `id` varchar(50) NOT NULL,
  `mapa_id` int(11) NOT NULL,
  `objeto_id` varchar(50) NOT NULL,
  `x` float NOT NULL DEFAULT '0',
  `y` float NOT NULL DEFAULT '0',
  `ancho` float NOT NULL DEFAULT '0',
  `largo` float NOT NULL DEFAULT '0',
  `rotacion` float NOT NULL DEFAULT '0',
  `idx` float NOT NULL DEFAULT '0',
  `idy` float NOT NULL DEFAULT '0',
  `rotacion_texto` float NOT NULL DEFAULT '0',
  `tipo_id` int(11) NOT NULL,
  `color_borde` varchar(20) NOT NULL DEFAULT '#FFF',
  `alfa_borde` float DEFAULT '1',
  `color_fondo` varchar(20) NOT NULL DEFAULT '#FFF',
  `alfa_fondo` float DEFAULT '1',
  `color_texto` varchar(20) NOT NULL DEFAULT '#000',
  `fuente` varchar(50) NOT NULL DEFAULT 'Sans-serif',
  `tamano_fuente` float NOT NULL DEFAULT '0',
  `fuente_italica` tinyint(4) NOT NULL DEFAULT '0',
  `fuente_negrita` tinyint(4) NOT NULL DEFAULT '0',
  `archivo` varchar(300) NOT NULL DEFAULT '',
  `status_asociado` int(11) NOT NULL,
  `ultima_actualizacion` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  `mensaje` varchar(300) NOT NULL DEFAULT '',
  `efecto` varchar(2) NOT NULL DEFAULT '',
  `ancho_borde` float NOT NULL DEFAULT '0',
  PRIMARY KEY (`id`,`mapa_id`),
  KEY `mapa_id` (`mapa_id`,`objeto_id`),
  KEY `tipo_id` (`tipo_id`),
  KEY `status_asociado` (`status_asociado`),
  KEY `figuras_ibfk_3` (`status_asociado`,`mapa_id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

/*Data for the table `figuras` */

insert  into `figuras`(`id`,`mapa_id`,`objeto_id`,`x`,`y`,`ancho`,`largo`,`rotacion`,`idx`,`idy`,`rotacion_texto`,`tipo_id`,`color_borde`,`alfa_borde`,`color_fondo`,`alfa_fondo`,`color_texto`,`fuente`,`tamano_fuente`,`fuente_italica`,`fuente_negrita`,`archivo`,`status_asociado`,`ultima_actualizacion`,`mensaje`,`efecto`,`ancho_borde`) values ('shp1026',124,'44',958.259,648.797,68,67,0,971.612,602.903,0,125,'#FFFFFF',2147480000,'#FFFFFF',1,'#000000','Calibri',11,0,0,'img-124-1026.png',0,'2020-07-28 21:06:33','','',0),('shp375',124,'17',706.143,249.672,86,46,0,777.832,275.724,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-375.png',10,'2020-07-28 21:06:33','','P',0),('shp355',124,'37',642.303,92.3594,149,60,0,678.369,156.412,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-355.png',10,'2020-07-28 21:06:33','','P',0),('shp369',124,'18',706.948,291.646,85,56,0,779.939,324.557,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-369.png',10,'2020-07-28 21:06:33','','P',0),('shp377',124,'17',706.143,249.772,86,46,0,780.566,260.668,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-377.png',0,'2020-07-28 21:06:33','','P',0),('shp367',124,'18',706.503,291.698,85,56,0,777.832,309.794,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-367.png',0,'2020-07-28 21:06:33','','P',0),('shp357',124,'37',642.355,92.4608,149,60,0,683.333,140.275,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-357.png',0,'2020-07-28 21:06:33','','P',0),('shp346',124,'36',663.147,12.9229,120,58,0,669.539,77.6469,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-346.png',10,'2020-07-28 21:06:33','','P',0),('shp348',124,'36',662.633,12.8077,119,58,0,669.192,64.1502,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-348.png',0,'2020-07-28 21:06:33','','P',0),('shp338',124,'39',591.121,26.4686,78,59,0,603.188,90.0762,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-338.png',10,'2020-07-28 21:06:33','','P',0),('shp331',124,'47',581.722,4.34898,28,35,0,569.999,50.3975,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-331.png',10,'2020-07-28 21:06:33','','P',0),('shp329',124,'47',582.195,4.55223,29,35,0,570.192,37.5139,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-329.png',0,'2020-07-28 21:06:33','','P',0),('shp340',124,'39',590.773,26.9677,77,59,0,602.984,78.2165,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-340.png',0,'2020-07-28 21:06:33','','P',0),('shp321',124,'48',632.297,228.155,32,33,0,619.39,262.608,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-321.png',10,'2020-07-28 21:06:33','','P',0),('shp323',124,'48',632.483,228.18,33,33,0,620.7,250.368,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-323.png',0,'2020-07-28 21:06:33','','P',0),('shp315',124,'42',657.598,342.575,133,40,0,780.534,350.676,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-315.png',10,'2020-07-28 21:06:33','','P',0),('shp313',124,'42',657.62,342.941,133,41,0,780.066,335.024,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-313.png',0,'2020-07-28 21:06:33','','P',0),('shp305',124,'22',360.206,449.04,89,64,0,253.811,556.485,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-305.png',10,'2020-07-28 21:06:33','','P',0),('shp307',124,'22',360.25,449.255,89,65,0,254.77,543.753,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-307.png',0,'2020-07-28 21:06:33','','P',0),('shp297',124,'23',364.407,509.963,82,78,0,256.216,509.906,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-297.png',0,'2020-07-28 21:06:33','','P',0),('shp299',124,'23',364.58,509.086,81,78,0,258.102,524.044,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-299.png',10,'2020-07-28 21:06:33','','P',0),('shp289',124,'24',450.438,462.873,69,70,0,255.347,495.244,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-289.png',10,'2020-07-28 21:06:33','','P',0),('shp291',124,'24',450.917,463.058,68,70,0,253.402,481.121,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-291.png',0,'2020-07-28 21:06:33','','P',0),('shp283',124,'20',446.501,381.785,68,84,0,257.317,461.028,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-283.png',10,'2020-07-28 21:06:33','','P',0),('shp281',124,'20',445.969,382.051,68,84,0,257.598,444.579,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-281.png',0,'2020-07-28 21:06:33','','P',0),('shp272',124,'21',351.489,384.122,70,84,0,259.133,402.751,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-272.png',0,'2020-07-28 21:06:33','','P',0),('shp270',124,'21',351.259,383.599,70,83,0,259.563,417.113,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-270.png',10,'2020-07-28 21:06:33','','P',0),('shp258',124,'19',566.244,415.931,60,107,0,514.483,431.965,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-258.png',0,'2020-07-28 21:06:33','','P',0),('shp260',124,'19',566.244,415.55,60,107,0,514.908,445.207,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-260.png',10,'2020-07-28 21:06:33','','P',0),('shp249',124,'32',671.263,602.767,91,115,0,759.196,628.752,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-249.png',0,'2020-07-28 21:06:33','','P',0),('shp251',124,'32',668.771,601.528,91,114,0,762.971,645.203,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-251.png',10,'2020-07-28 21:06:33','','P',0),('shp238',124,'41',689.004,457.327,51,54,0,772.493,474.817,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-238.png',10,'2020-07-28 21:06:33','','P',0),('shp236',124,'41',688.761,457.785,51,53,0,772.949,487.893,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-236.png',0,'2020-07-28 21:06:33','','P',0),('shp222',124,'30',706.877,508.586,22,48,0,686.051,559.17,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-222.png',10,'2020-07-28 21:06:33','','P',0),('shp230',124,'29',615.44,491.406,85,80,0,625.155,556.485,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-230.png',0,'2020-07-28 21:06:33','','P',0),('shp228',124,'29',615.477,491.406,85,80,0,623.669,568.618,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-228.png',10,'2020-07-28 21:06:33','','P',0),('shp209',124,'40',735.698,463.231,40,21,0,777.832,458.366,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-209.png',0,'2020-07-28 21:06:33','','P',0),('shp220',124,'30',706.877,508.586,22,48,0,687.046,547.438,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-220.png',0,'2020-07-28 21:06:33','','P',0),('shp200',124,'31',736.021,389.679,57,77,0,731.627,368.895,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-200.png',10,'2020-07-28 21:06:33','','P',0),('shp207',124,'40',735.092,464.014,41,21,0,777.095,470.946,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-207.png',10,'2020-07-28 21:06:33','','P',0),('shp198',124,'31',736.178,389.611,57,77,0,731.309,379.003,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-198.png',0,'2020-07-28 21:06:33','','P',0),('shp186',124,'46',516.725,313.61,34,46,0,498.135,367.603,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-186.png',10,'2020-07-28 21:06:33','','P',0),('shp188',124,'46',517.193,313.609,34,46,0,499.654,353.211,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-188.png',0,'2020-07-28 21:06:33','','P',0),('shp178',124,'45',550.768,184.458,30,66,0,541.803,221.566,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-178.png',0,'2020-07-28 21:06:33','','P',0),('shp180',124,'45',550.488,184.694,31,65,0,542.953,234.043,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-180.png',10,'2020-07-28 21:06:33','','P',0),('shp171',124,'49',494.784,17.7723,45,48,0,486.675,58.7303,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-171.png',0,'2020-07-28 21:06:33','','P',0),('shp161',124,'34',545.746,-0.20378,34,37,0,540.612,27.9002,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-161.png',0,'2020-07-28 21:06:33','','P',0),('shp163',124,'34',545.368,-0.0570079,35,37,0,540.023,37.1461,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-163.png',10,'2020-07-28 21:06:33','','P',0),('shp153',124,'35',505.576,0.500053,35,31,0,497.342,44.8407,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-153.png',0,'2020-07-28 21:06:33','','P',0),('shp155',124,'35',505.576,0.380997,35,32,0,497.805,31.2873,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-155.png',10,'2020-07-28 21:06:33','','P',0),('shp116',124,'2',474.43,116.859,101,47,0,469.933,163.263,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-116.png',10,'2020-07-28 21:06:33','','P',0),('shp137',124,'1',345.064,107.382,50,39,0,343.312,139.912,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-137.png',0,'2020-07-28 21:06:33','','P',0),('shp145',124,'33',451.832,8.10079,59,38,0,448.639,54.7689,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-145.png',0,'2020-07-28 21:06:33','','P',0),('shp147',124,'33',451.396,8.21354,60,38,0,449.313,67.2738,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-147.png',10,'2020-07-28 21:06:33','','P',0),('shp124',124,'2',474.322,116.993,101,47,0,470.436,151.375,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-124.png',0,'2020-07-28 21:06:33','','P',0),('shp125',124,'27',613.99,449.798,77,52,0,567.506,492.281,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-125.png',0,'2020-07-28 21:06:33','','P',0),('shp117',124,'28',675.86,408.883,64,51,0,671.072,384.852,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-117.png',10,'2020-07-28 21:06:33','','P',0),('shp115',124,'28',676.083,408.897,64,51,0,672.294,370.877,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-115.png',0,'2020-07-28 21:06:33','','P',0),('shp109',124,'26',615.45,399.801,65,60,0,606.473,373.018,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-109.png',0,'2020-07-28 21:06:33','','P',0),('shp100',124,'26',614.798,400.004,64,60,0,608.052,385.827,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-100.png',10,'2020-07-28 21:06:33','','P',0),('shp65',124,'25',439.581,518.853,78,57,0,488.898,537.777,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-65.png',10,'2020-07-28 21:06:33','','P',0),('shp42',124,'25',439.581,519.24,78,57,0,490.319,524.044,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-42.png',0,'2020-07-28 21:06:33','','P',0),('shp101',124,'14',608.731,319.685,64,62,0,632.944,308.36,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-101.png',0,'2020-07-28 21:06:33','','P',0),('shp99',124,'14',609.378,319.685,65,62,0,632.659,320.342,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-99.png',10,'2020-07-28 21:06:33','','P',0),('shp91',124,'16',547.932,325.147,65,53,0,543.617,356.731,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-91.png',0,'2020-07-28 21:06:33','','P',0),('shp63',124,'13',609.137,271.116,57,54,0,657.967,288.961,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-63.png',0,'2020-07-28 21:06:33','','P',0),('shp71',124,'13',608.938,270.866,57,53,0,657.967,300.294,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-71.png',10,'2020-07-28 21:06:33','','P',0),('shp72',124,'9',661.838,215.25,76,44,0,620.977,212.384,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-72.png',0,'2020-07-28 21:06:33','','P',0),('shp80',124,'10',735.756,217.183,57,39,0,697.207,241.88,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-80.png',0,'2020-07-28 21:06:33','','P',0),('shp82',124,'10',735.776,217.16,57,39,0,697.815,254.983,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-82.png',10,'2020-07-28 21:06:33','','P',0),('shp34',124,'15',555.388,264.233,56,68,0,560.446,250.501,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-34.png',0,'2020-07-28 21:06:33','','P',0),('shp74',124,'9',661.838,215.25,76,44,0,620.977,232.499,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-74.png',10,'2020-07-28 21:06:33','','P',0),('shp64',124,'8',731.633,175.77,61,34,0,714.197,152.553,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-64.png',0,'2020-07-28 21:06:33','','P',0),('shp55',124,'7',665.089,174.13,48,38,0,584.727,155.71,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-55.png',0,'2020-07-28 21:06:33','','P',0),('shp57',124,'7',664.732,174.13,48,38,0,582.776,174.13,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-57.png',10,'2020-07-28 21:06:33','','P',0),('shp50',124,'12',467.97,242.162,91,68,0,238.688,217.637,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-50.png',0,'2020-07-28 21:06:33','','P',0),('shp36',124,'11',464.75,183.175,87,66,0,245.396,149.492,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-36.png',0,'2020-07-28 21:06:33','','P',0),('shp20',124,'11',464.75,183.5,87,65,0,245.396,162.765,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-20.png',10,'2020-07-28 21:06:33','','P',0),('shp35',124,'6',361.735,325.185,72,56,0,240.319,343.788,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-35.png',0,'2020-07-28 21:06:33','','P',0),('shp37',124,'6',361.215,325.613,72,56,0,240.319,363.606,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-37.png',10,'2020-07-28 21:06:33','','P',0),('shp21',124,'5',364.75,281.032,70,52,0,242.595,296.483,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-21.png',0,'2020-07-28 21:06:33','','P',0),('shp27',124,'4',366.801,232.955,68,53,0,242.761,249.001,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-27.png',0,'2020-07-28 21:06:33','','P',0),('shp25',124,'4',366.801,232.955,68,53,0,242.761,262.734,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-25.png',10,'2020-07-28 21:06:33','','P',0),('shp17',124,'3',369.2,186.676,66,51,0,242.761,179.064,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-17.png',0,'2020-07-28 21:06:33','','P',0),('shp6',124,'3',369.2,186.676,66,51,0,243.213,196.611,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-6.png',10,'2020-07-28 21:06:33','','P',0),('shp7',124,'43',380.033,0,45,93,0,213.213,69.7514,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-7.png',0,'2020-07-28 21:06:33','','P',0),('shp3',124,'43',380.033,-0.0394751,45,93,0,213.213,86.3606,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-3.png',10,'2020-07-28 21:06:33','','P',0),('shp8',124,'44',339.833,0,44,83,0,323.182,578.409,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#FF0000','Calibri',14,0,1,'img-124-8.png',0,'2020-07-28 21:06:33','NITREX 1','PP',0),('shp66',124,'8',731.875,175.52,61,34,0,713,163.263,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-66.png',10,'2020-07-28 21:06:33','','P',0),('shp48',124,'12',468.041,242.666,90,66,0,240.319,232.988,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-48.png',10,'2020-07-28 21:06:33','','P',0),('shp16',124,'5',364.402,280.679,71,52,0,241.609,313.677,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-16.png',10,'2020-07-28 21:06:33','','P',0),('shp10',124,'44',339.833,0,44,83,0,217.744,20.0712,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-10.png',10,'2020-07-28 21:06:33','','P',0),('shp47',124,'15',555.308,264.354,55,69,0,560.549,263.481,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-47.png',10,'2020-07-28 21:06:33','','P',0),('shp93',124,'16',548,325.136,65,53,0,544.137,370.006,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-93.png',10,'2020-07-28 21:06:33','','P',0),('shp123',124,'27',614.102,449.772,77,52,0,566.495,506.938,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-123.png',10,'2020-07-28 21:06:33','','P',0),('shp139',124,'1',345.064,107.29,50,39,0,344.331,153.645,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-139.png',10,'2020-07-28 21:06:33','','P',0),('shp169',124,'49',495.085,17.5951,46,48,0,483.823,79.2877,0,125,'#FFFFFF',2147480000,'#FFFFFF',0,'#000000','Calibri',11,0,0,'img-124-169.png',10,'2020-07-28 21:06:33','','P',0);

/*Table structure for table `horarios` */

DROP TABLE IF EXISTS `horarios`;

CREATE TABLE `horarios` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `calendario` bigint(20) DEFAULT NULL COMMENT 'ID del calendario',
  `clase` int(1) DEFAULT '0' COMMENT 'Clase de horario (0=Disponibilidad de producción, 1=Mantenimiento)',
  `tipo` char(1) DEFAULT 'S' COMMENT 'Tipo de horario (S/N)',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la maquina',
  `dia` int(1) DEFAULT '0' COMMENT 'Dia de semana (9=por fecha)',
  `fecha` date DEFAULT NULL COMMENT 'Fecha a revisar',
  `desde` time DEFAULT NULL COMMENT 'Hora desde',
  `hasta` time DEFAULT NULL COMMENT 'Hora hasta',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=52 DEFAULT CHARSET=latin1 COMMENT='Horarios';

/*Data for the table `horarios` */

/*Table structure for table `int_eventos` */

DROP TABLE IF EXISTS `int_eventos`;

CREATE TABLE `int_eventos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `nombre` varchar(100) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `monitor` char(1) DEFAULT NULL COMMENT 'Monitorear (S/N)',
  `alerta` bigint(20) DEFAULT '0' COMMENT 'ID de la alerta asociado',
  `revision` int(6) DEFAULT '0' COMMENT 'Tiempo de revisión',
  `revisado` datetime DEFAULT NULL COMMENT 'Fecha y hora la última revisión',
  `prioridad` int(3) DEFAULT NULL COMMENT 'Prioridad de la revisión',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=36 DEFAULT CHARSET=latin1 COMMENT='Catálogo de eventos';

/*Data for the table `int_eventos` */

insert  into `int_eventos`(`id`,`planta`,`nombre`,`monitor`,`alerta`,`revision`,`revisado`,`prioridad`,`estatus`,`creacion`,`modificacion`,`creado`,`modificado`) values (1,1,'Tiempo de soporte excedido','S',101,3,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,1,1),(2,1,'Tiempo de reparación excedido','S',102,3,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,1,1),(3,1,'Tiempo de documentación excedido','N',103,20,'2020-01-28 16:50:22',NULL,'A',NULL,NULL,1,1),(18,1,'OAE - Bajo rate','S',201,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,1,1),(19,1,'OAE - Sobre rate','S',202,10,'2020-08-01 09:20:10',NULL,'A',NULL,NULL,1,1),(20,1,'OAE - No se detectan piezas','S',203,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,1,1),(23,1,'WIP - Tiempo de stock excedido','N',302,10,'2020-01-28 17:45:42',NULL,'A',NULL,NULL,0,0),(24,1,'WIP - Tiempo de proceso excedido','N',303,10,'2020-01-28 17:47:49',NULL,'A',NULL,NULL,0,0),(25,1,'WIP - Tiempo de entrega excedido','N',304,10,'2020-01-28 09:30:21',NULL,'A',NULL,NULL,0,0),(26,1,'WIP - Antic/Tiempo de stock','N',305,10,'2020-01-28 17:57:47',NULL,'A',NULL,NULL,0,0),(27,1,'WIP - Antic/Tiempo de proceso','N',306,10,'2020-01-28 18:02:04',NULL,'A',NULL,NULL,0,0),(28,1,'WIP - Antic/Tiempo de entrega','S',307,10,'2020-08-01 09:20:10',NULL,'A',NULL,NULL,0,0),(29,1,'WIP - Salto de operación','N',301,10,'2020-01-28 17:38:44',NULL,'A',NULL,NULL,0,0),(30,0,'OAE - Baja en FTQ','S',204,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,0,0),(31,0,'OAE - Baja en Disponibilidad','S',205,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,0,0),(32,0,'OAE - Bajo en Desempeño','S',206,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,0,0),(33,0,'OAE - Baja en OAE','S',207,5,'2020-08-01 09:20:16',NULL,'A',NULL,NULL,0,0),(35,1,'CL - Checklist ausencia de llenado','S',401,10,NULL,NULL,'A',NULL,NULL,0,0);

/*Table structure for table `int_listados` */

DROP TABLE IF EXISTS `int_listados`;

CREATE TABLE `int_listados` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del correo',
  `nombre` varchar(50) DEFAULT NULL COMMENT 'Reporte',
  `datos` char(1) DEFAULT 'S' COMMENT 'Permite datos en csv',
  `grafica` char(1) DEFAULT 'S' COMMENT 'Permite gráficas',
  `orden` int(3) DEFAULT '0' COMMENT 'Orden en la vista',
  `file_name` varchar(200) DEFAULT NULL COMMENT 'Nombre del archivo',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Activo?',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=45 DEFAULT CHARSET=latin1 COMMENT='Reportes para el negocio';

/*Data for the table `int_listados` */

insert  into `int_listados`(`id`,`nombre`,`datos`,`grafica`,`orden`,`file_name`,`estatus`) values (1,'MTTR por línea','S','S',10,'mttr_linea','A'),(2,'MTTR por máquina','S','S',11,'mttr_maquina','A'),(3,'MTTR por área','S','S',12,'mttr_area','m'),(4,'MTTR por falla','S','S',13,'mttr_falla','A'),(5,'MTTR por tipo de máquina','S','S',14,'mttr_tipo_maquina','A'),(6,'MTTR por grupo de máquinas (1)','S','S',15,'mttr_grupo1_maquina','A'),(7,'MTTR por grupo de máquinas (2)','S','S',16,'mttr_grupo2_maquina','A'),(8,'MTTR por grupo de fallas (1)','S','S',17,'mttr_grupo1_falla','A'),(9,'MTTR por grupo de fallas (2)','S','S',18,'mttr_grupo2_falla','A'),(10,'MTTR por dia','S','S',19,'mttr_dia','A'),(11,'MTTR por semana','S','S',20,'mttr_semana','A'),(12,'MTTR por mes','S','S',21,'mttr_mes','A'),(13,'MTTR por técnico','S','S',22,'mttr_tecnico','A'),(14,'MTBF por línea','S','S',23,'mtbf_linea','A'),(15,'MTBF por maquina','S','S',24,'mtbf_maquina','A'),(16,'MTBF por area','S','S',25,'mtbf_area','A'),(17,'MTBF por falla','S','S',26,'mtbf_falla','A'),(18,'MTBF por tipo de máquina','S','S',27,'mtbf_tipo_maquina','A'),(19,'MTBF por grupo de máquinas (1)','S','S',28,'mtbf_grupo1_maquina','A'),(20,'MTBF por grupo de máquinas (2)','S','S',29,'mtbf_grupo2_maquina','A'),(21,'MTBF por grupo de fallas (1)','S','S',30,'mtbf_grupo1_falla','A'),(22,'MTBF por grupo de fallas (2)','S','S',31,'mtbf_grupo2_falla','A'),(23,'MTBF por dia','S','S',32,'mtbf_dia','A'),(24,'MTBF por semana','S','S',33,'mtbf_semana','A'),(25,'MTBF por mes','S','S',34,'mtbf_mes','A'),(26,'MTBF por técnico','S','S',35,'mtbf_tecnico','A'),(27,'Pareto por linea','S','S',36,'pareto_linea','A'),(28,'Pareto por maquina','S','S',37,'pareto_mquina','A'),(29,'Pareto por area','S','S',38,'pareto_area','A'),(30,'Pareto por falla','S','S',39,'pareto_falla','A'),(31,'Pareto por tipo de máquina','S','S',40,'pareto_tipo_maquina','A'),(32,'Pareto por grupo de máquinas (1)','S','S',41,'pareto_grupo1_maquina','A'),(33,'Pareto por grupo de máquinas (2)','S','S',42,'pareto_grupo2_maquina','A'),(34,'Pareto por grupo de fallas (1)','S','S',43,'pareto_grupo1_falla','A'),(35,'Pareto por grupo de fallas (2)','S','S',44,'pareto_grupo2_falla','A'),(36,'Pareto por dia','S','S',45,'pareto_dia','A'),(37,'Pareto por semana','S','S',46,'pareto_semana','A'),(38,'Pareto por mes','S','S',47,'pareto_mes','A'),(39,'Pareto por técnico','S','S',48,'pareto_tecnico','A'),(42,'Reportes de mantenimiento TODOS LOS ESTATUS','S','N',1,'reportes_todos','A');

/*Table structure for table `int_opciones` */

DROP TABLE IF EXISTS `int_opciones`;

CREATE TABLE `int_opciones` (
  `id` int(6) DEFAULT NULL COMMENT 'ID de la opción',
  `rol` char(1) DEFAULT NULL COMMENT 'ID del rol',
  `nombre` varchar(60) DEFAULT NULL COMMENT 'Descripción de la opción',
  `orden` int(4) DEFAULT NULL COMMENT 'Orden en la pantalla de usuarios',
  `visualizar` char(1) DEFAULT 'S' COMMENT 'Visualizar en el sistema',
  `acciones` varchar(10) DEFAULT 'SSSSS' COMMENT 'Acciones (Visualizar, Crear, Editar, Inactivar/Reactivar, Eliminar)',
  `opcion_app` int(4) DEFAULT '0' COMMENT 'Opción de la app',
  `url` varchar(40) DEFAULT NULL COMMENT 'URL a ejecutar',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus de la opción'
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Interna - Opciones';

/*Data for the table `int_opciones` */

insert  into `int_opciones`(`id`,`rol`,`nombre`,`orden`,`visualizar`,`acciones`,`opcion_app`,`url`,`estatus`) values (10,'O','Gestión de kanban',10,'N','SSSSS',12,'/exportar','I'),(20,'A','Visualizar alarmas',20,'S','SSSSS',23,'/exportar','A'),(30,'A','Terminar alertas',30,'S','SSSSS',9001,'" & traduccion(145) & "','A'),(40,'S','Confirmación de reparaciones',60,'S','SSSSS',9002,'" & traduccion(145) & "','A'),(50,'O','Generación de llamadas ANDON',40,'S','SSSSS',10,'/andon','A'),(60,'T','Atención a llamadas de ANDON',50,'S','SSSSS',11,'/andon','A'),(70,'G','Mantenimiento de Líneas/Células',60,'S','SSSSS',30,'/catalogos','A'),(80,'G','Mantenimiento de Máquinas',70,'S','SSSSS',31,'/catalogos','A'),(90,'G','Mantenimiento de Áreas',80,'S','SSSSS',32,'/catalogos','A'),(100,'G','Mantenimiento de Fallas',90,'S','SSSSS',33,'/catalogos','A'),(110,'G','Mantenimiento de Tablas generales',100,'S','SSSSS',34,'/catalogos','A'),(120,'G','Mantenimiento de Turnos',110,'S','SSSSS',38,'/catalogos','A'),(130,'G','Mantenimiento de Traductor',120,'S','SSSSS',39,'/catalogos','A'),(140,'G','Mantenimiento de Correos/Reportes',130,'S','SSSSS',36,'/catalogos','A'),(150,'G','Mantenimiento de Recipientes',140,'S','SSSSS',35,'/catalogos','A'),(160,'G','Mantenimiento de Alertas',150,'S','SSSSS',37,'/catalogos','A'),(170,'A','Gestión de usuarios',160,'S','SSSSS',41,'/catalogos','A'),(180,'A','Gestión de parámetros',170,'S','SSSSS',42,'/parametros','A'),(190,'*','MULTI-VISOR',180,'S','SSSSS',20,'/visor','A'),(200,'*','Consulta de gráficos',190,'S','SSSSS',21,'/graficas','A'),(210,'G','Descargar datos del sistema',200,'S','SSSSS',22,'/exportar','A'),(220,'A','Gestión de políticas',165,'S','SSSSS',43,'/catalogos','A'),(230,'*','Panel de control OEE y ANDON',175,'S','SSSSS',52,'/panel','A'),(240,'A','Gestión de reportes',170,'S','SSSSS',45,'Comp','A'),(250,'A','Equipamiento',180,'S','SSSSS',44,'/parametros','A'),(15,'O','Producción (OEE)',15,'S','SSSSS',13,'/produccion','A'),(260,'S','OEE - Reiniciar el conteo',63,'S','SSSSS',0,'" & traduccion(145) & "','A'),(270,'S','OEE - Iniciar un paro manual',61,'S','SSSSS',0,'" & traduccion(145) & "','A'),(280,'S','OEE - Terminar un paro',62,'S','SSSSS',0,'" & traduccion(145) & "','A'),(290,'S','OEE - Cambiar parámetros',64,'S','SSSSS',0,'" & traduccion(145) & "','A'),(300,'G','OEE - Mantenimiento de Rates de producción',131,'S','SSSSS',46,'/catalogos','A'),(310,'G','OEE - Mantenimiento de Objetivos',132,'S','SSSSS',47,'/catalogos','A'),(320,'G','OEE - Mantenimiento de Estimados de producción',133,'S','SSSSS',48,'/catalogos','A'),(330,'G','OEE - Gestión de sensores',124,'S','SSSSS',49,'/catalogos','A'),(340,'G','OEE - Gestión de paros',126,'S','SSSSS',50,'/catalogos','A'),(350,'*','Solicitud de SMED/Changeover',127,'S','SSSSS',51,'Comp','A'),(103,'G','Mantenimiento de Numeros de parte',93,'S','SSSSS',2003,'/catalogoswip','A'),(101,'G','Mantenimiento de Procesos',91,'S','SSSSS',2001,'/catalogoswip','A'),(102,'G','Mantenimiento de Rutas de fabricación',92,'S','SSSSS',2002,'/catalogoswip','A'),(241,'A','Gestión de lotes',175,'S','SSSSS',53,'Comp','A'),(295,'O','WIP - Permitir reverso de lotes',65,'S','SSSSS',0,'" & traduccion(145) & "','A'),(111,'O','Flujo de materiales (WIP)',111,'S','SSSSS',3001,'/operaciones','A'),(104,'G','Mantenimiento de Horarios de trabajo',94,'S','SSSSS',2007,'/catalogoswip','A'),(105,'G','Mantenimiento de Situaciones de calidad (WIP)',95,'S','SSSSS',2006,'/catalogoswip','A'),(106,'G','Mantenimiento de Prioridades de entrega (WIP)',97,'S','SSSSS',2011,'/flujo','A'),(107,'G','Programación de equipo (carga) (WIP)',96,'S','SSSSS',2010,'/flujo','A'),(108,'S','Inspección de lotes en cuarentena (WIP)',98,'S','SSSSS',2012,'/flujo','A'),(109,'S','Inspección de lotes rechazados (WIP)',99,'S','SSSSS',2013,'/flujo','A'),(110,'O','Consulta de inventarios (WIP)',100,'S','SSSSS',2014,'/flujo','A'),(296,'O','WIP - Enviar un lote a inspección',66,'S','SSSSS',3001,NULL,'A'),(297,'O','WIP - Rechazar un lore',67,'S','SSSSS',3002,NULL,'A'),(298,'O','WIP - Liberar un lote',68,'S','SSSSS',3003,NULL,'A'),(242,'O','Mantenimiento de Números de parte',176,'S','SSSSS',54,'/catalogos','A'),(243,'O','Documentación de reportes',177,'S','SSSSS',55,'Comp','A'),(245,'G','Gestión de rechazos',178,'S','SSSSS',56,'/catalogos','A'),(185,'G','Mantenimiento de Variables de checklist',129,'S','SSSSS',57,'/catalogos','A'),(187,'G','Mantenimiento de Checklists',130,'S','SSSSS',58,'/catalogos','A'),(189,'G','Planeación de Checklists',131,'S','SSSSS',59,'/catalogos','A'),(360,'O','OEE - Sincronizar sensores Buffer-Producción',134,'S','SSSSS',0,NULL,'A'),(370,'O','OEE - Ajuste de sensores',135,'S','SSSSS',0,NULL,'A');

/*Table structure for table `kanban` */

DROP TABLE IF EXISTS `kanban`;

CREATE TABLE `kanban` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `linea` bigint(20) DEFAULT NULL COMMENT 'ID de la línea',
  `equipo` bigint(20) DEFAULT NULL COMMENT 'ID del equipo',
  `ruta` bigint(20) DEFAULT NULL COMMENT 'ID de la ruta',
  `operacion` bigint(20) DEFAULT NULL COMMENT 'ID de la operacion',
  `orden` varchar(50) DEFAULT NULL COMMENT 'ID de la O/P',
  `parte` varchar(100) DEFAULT NULL COMMENT 'Número de parte',
  `existencia` decimal(25,6) DEFAULT NULL COMMENT 'Cantidad en Stock',
  `desde` datetime DEFAULT NULL COMMENT 'Fecha y hora de entrada',
  `estatus` int(2) DEFAULT NULL COMMENT 'Estatus del stock',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Detalle de existencias en kanban';

/*Data for the table `kanban` */

/*Table structure for table `lecturas` */

DROP TABLE IF EXISTS `lecturas`;

CREATE TABLE `lecturas` (
  `id` int(20) NOT NULL AUTO_INCREMENT COMMENT 'Número automático único',
  `fecha` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha y hora automática en que se genera el registro',
  `sensor` bigint(10) DEFAULT NULL COMMENT 'ID del sensor que genera la señal',
  `valor` varchar(255) COLLATE utf16_bin DEFAULT '' COMMENT 'Valor recibido del sensor',
  `estatus` int(2) DEFAULT '0' COMMENT 'Estatus 0=Creado, 1=En proceso, 2=Procesado',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`estatus`),
  KEY `NewIndex2` (`fecha`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=3469063 DEFAULT CHARSET=utf16 COLLATE=utf16_bin COMMENT='Tabla histórica de lectura de los sensores';

/*Data for the table `lecturas` */

/*Table structure for table `lecturas_cortes` */

DROP TABLE IF EXISTS `lecturas_cortes`;

CREATE TABLE `lecturas_cortes` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `dia` date DEFAULT NULL COMMENT 'Dia del corte',
  `orden` bigint(20) DEFAULT '0' COMMENT 'O/P, Lote',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `tc` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Tiempo ciclo',
  `tripulacion` bigint(20) DEFAULT '0' COMMENT 'ID de la tripulación',
  `produccion` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad de eventos producidos',
  `calidad` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad de eventos rechazados',
  `buffer` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad en buffer',
  `produccion_tc` decimal(25,10) DEFAULT '0.0000000000',
  `calidad_tc` decimal(25,10) DEFAULT '0.0000000000',
  `paro` bigint(20) DEFAULT '0' COMMENT 'ID del paro actual',
  `paro_actual` bigint(20) DEFAULT '0',
  `tiempo_disponible` bigint(10) DEFAULT '0' COMMENT 'Tiempo disponible',
  `bloque_inicia` datetime DEFAULT NULL COMMENT 'Fecha de inicio',
  `bloque_finaliza` datetime DEFAULT NULL COMMENT 'Fecha de finalización',
  `calidad_clasificada` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Piezas clasificadas',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`equipo`),
  KEY `NewIndex2` (`equipo`,`id`),
  KEY `NewIndex3` (`dia`,`equipo`,`orden`,`parte`,`turno`,`tc`)
) ENGINE=MyISAM AUTO_INCREMENT=12187 DEFAULT CHARSET=latin1 COMMENT='Acumulado de piezas contadas';

/*Data for the table `lecturas_cortes` */

/*Table structure for table `lecturas_resumen` */

DROP TABLE IF EXISTS `lecturas_resumen`;

CREATE TABLE `lecturas_resumen` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `tipo` int(1) DEFAULT '0' COMMENT 'Tipo de color (0: sin plan, 1: Produccion, 2: bajo rate, 3: en paro)',
  `paro` int(4) DEFAULT '0',
  `bajorate` int(4) DEFAULT '0',
  `sinplan` int(4) DEFAULT '0',
  `produccion` int(4) DEFAULT '0',
  `desde` datetime DEFAULT NULL,
  `hasta` datetime DEFAULT NULL,
  `hora` int(2) DEFAULT '0',
  PRIMARY KEY (`id`),
  KEY `NewIndex2` (`equipo`,`desde`,`hasta`),
  KEY `NewIndex1` (`equipo`,`desde`)
) ENGINE=MyISAM AUTO_INCREMENT=274316 DEFAULT CHARSET=latin1 COMMENT='Histórico últimas 24 horas';

/*Data for the table `lecturas_resumen` */

/*Table structure for table `licencias` */

DROP TABLE IF EXISTS `licencias`;

CREATE TABLE `licencias` (
  `tipo` char(1) DEFAULT NULL COMMENT 'Tipo de dispositivo (B=Botonera, R=Reloj)',
  `mmcall` varchar(50) DEFAULT NULL COMMENT 'Valor MMCall',
  `cronos` varchar(50) DEFAULT NULL COMMENT 'Clave de Cronos',
  `inicio` datetime DEFAULT NULL COMMENT 'Fecha de inicio de la licencia',
  `licenciado` date DEFAULT NULL COMMENT 'Fecha de licencia',
  `vencimiento` datetime DEFAULT NULL COMMENT 'Fecha de vencimiento',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus de la licencia',
  KEY `NewIndex1` (`tipo`,`mmcall`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Licenciamiento de equipos';

/*Data for the table `licencias` */

/*Table structure for table `log` */

DROP TABLE IF EXISTS `log`;

CREATE TABLE `log` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `fecha` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha y hora del registro',
  `aplicacion` int(6) DEFAULT NULL COMMENT 'ID de la aplicación',
  `tipo` int(1) DEFAULT '0' COMMENT 'Tipo de mensaje',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'Número del proceso',
  `texto` varchar(250) DEFAULT NULL COMMENT 'Mensaje descriptivo (hasta 250 caracteres)',
  `visto` char(1) DEFAULT 'N' COMMENT 'Ya se vió en el visor?',
  `visto_pc` char(1) DEFAULT 'N' COMMENT 'Ya se vió en el log del PC?',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`fecha`),
  KEY `NewIndex2` (`visto_pc`),
  KEY `NewIndex3` (`visto`)
) ENGINE=MyISAM AUTO_INCREMENT=7630222 DEFAULT CHARSET=latin1;

/*Data for the table `log` */

/*Table structure for table `lotes` */

DROP TABLE IF EXISTS `lotes`;

CREATE TABLE `lotes` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `numero` varchar(100) DEFAULT NULL COMMENT 'Número la orden de producción/proceso',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del Número de parte',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha de entrada del lote',
  `hasta` datetime DEFAULT NULL COMMENT 'Fecha estimada de salida del proceso',
  `calcular_hasta` char(1) DEFAULT 'N' COMMENT 'Calcular el tiempo hasta del lote (N=No calcular, 1=Calcular tiempo de stock, 2=calcular tiempo de proceso)',
  `rechazo` datetime DEFAULT NULL COMMENT 'Fecha de rechazo',
  `inspeccion_id` bigint(20) DEFAULT '0' COMMENT 'Número de inspección',
  `rechazo_id` bigint(20) DEFAULT '0' COMMENT 'Número de rechazo',
  `inspeccion` datetime DEFAULT NULL COMMENT 'Fecha de inspeccion',
  `inspecciones` int(4) DEFAULT '0' COMMENT 'Veces que se inspecciona',
  `rechazos` int(4) DEFAULT '0' COMMENT 'Veces que se rechaza',
  `alarmas` int(4) DEFAULT '0' COMMENT 'Cantidas de alarmas en el lote',
  `reversos` int(4) DEFAULT '0' COMMENT 'Veces que se reversó',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `ruta` bigint(20) DEFAULT '0' COMMENT 'ID de la ruta',
  `ruta_detalle` bigint(20) DEFAULT '0' COMMENT 'ID del detalle de la ruta',
  `ruta_secuencia` int(6) DEFAULT '0' COMMENT 'Secuencia de la operación',
  `inicia` datetime DEFAULT NULL COMMENT 'Fecha de inicio en el sistema',
  `finaliza` datetime DEFAULT NULL COMMENT 'Fecha de fin en el sistema',
  `estimada` datetime DEFAULT NULL COMMENT 'Fecha estimada de completación',
  `tiempo_estimado` bigint(12) DEFAULT '0' COMMENT 'Tiempo estimado del lote',
  `tiempo` bigint(12) DEFAULT '0' COMMENT 'Tiempo total del lote',
  `estado` int(2) DEFAULT '0' COMMENT 'Estado del lote',
  `inspeccionado_por` bigint(20) DEFAULT '0' COMMENT 'Usuario que inspecciono la última vez',
  `rechazado_por` bigint(20) DEFAULT '0' COMMENT 'Usuario que rechazó la última vez',
  `alarma_tse` char(1) DEFAULT 'N' COMMENT 'Esta alarmada por tiempo de stock excedido',
  `alarma_tse_p` char(1) NOT NULL DEFAULT 'N',
  `alarma_tse_paso` char(1) DEFAULT 'N',
  `alarma_tpe` char(1) DEFAULT 'N' COMMENT 'Esta alarmada por tiempo de proceso excedido',
  `alarma_tpe_p` char(1) NOT NULL DEFAULT 'N',
  `alarma_tpe_paso` char(1) DEFAULT 'N',
  `alarma_plan` char(1) DEFAULT 'N' COMMENT 'Esta alarmada por programación no alcanzada',
  `carga` bigint(20) DEFAULT '0' COMMENT 'ID de la carga (temporal)',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`proceso`,`estado`,`estatus`,`creacion`),
  KEY `NewIndex2` (`parte`,`carga`),
  KEY `NewIndex3` (`carga`),
  KEY `NewIndex4` (`equipo`,`estado`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=2855 DEFAULT CHARSET=latin1 COMMENT='Lotes';

/*Data for the table `lotes` */

/*Table structure for table `lotes_cambiados` */

DROP TABLE IF EXISTS `lotes_cambiados`;

CREATE TABLE `lotes_cambiados` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `lote` bigint(20) DEFAULT NULL COMMENT 'ID del lote',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha de la actualización',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`lote`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='TRIGGER - Lotes cambiados';

/*Data for the table `lotes_cambiados` */

/*Table structure for table `lotes_historia` */

DROP TABLE IF EXISTS `lotes_historia`;

CREATE TABLE `lotes_historia` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `lote` bigint(20) DEFAULT '0' COMMENT 'ID de lote',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `ruta` bigint(20) DEFAULT '0' COMMENT 'ID de la ruta',
  `ruta_detalle` bigint(20) DEFAULT '0' COMMENT 'ID del detalle de la ruta',
  `ruta_detalle_anterior` bigint(20) DEFAULT '0' COMMENT 'ID del detalle de la ruta anterior',
  `ruta_secuencia` int(6) DEFAULT '0' COMMENT 'Secuecia de la operación',
  `ruta_secuencia_antes` int(6) DEFAULT '0' COMMENT 'Secuecia de la anterior',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `proceso_anterior` int(20) DEFAULT '0' COMMENT 'ID del proceso anterior',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `fecha_entrada` datetime DEFAULT NULL COMMENT 'Fecha de entrada del lote',
  `fecha_stock` datetime DEFAULT NULL COMMENT 'Fecha de enrada al stock',
  `fecha_proceso` datetime DEFAULT NULL COMMENT 'Fecha de entrada al proceso',
  `fecha_salida` datetime DEFAULT NULL COMMENT 'Fecha de salida del proceso',
  `fecha_estimada` datetime DEFAULT NULL COMMENT 'Fecha estimada del proceso (stock + proceso + setup)',
  `tiempo_estimado` bigint(12) DEFAULT '0' COMMENT 'Tiempo estimado en segundos',
  `tiempo_total` bigint(12) DEFAULT '0' COMMENT 'Tiempo total del lote en el proceso',
  `tiempo_espera` bigint(12) DEFAULT '0' COMMENT 'Tiempo del lote en la situación de espera',
  `tiempo_stock` bigint(12) DEFAULT '0' COMMENT 'Tiempo del lote en la situación de stock',
  `tiempo_proceso` bigint(12) DEFAULT '0' COMMENT 'Tiempo del lote en la situación de proceso',
  `alarma_so` char(1) DEFAULT 'N' COMMENT 'Alarmada por Salto de operación?',
  `alarma_so_rep` datetime DEFAULT NULL COMMENT 'Fecha de alarma por salto de operación',
  `inspecciones` int(4) DEFAULT '0' COMMENT 'Inspecciones en el proceso',
  `rechazos` int(4) DEFAULT '0' COMMENT 'Rechazos en el proceso',
  `reversado` char(1) DEFAULT NULL COMMENT 'El proceso fue reversado',
  `reversos` int(4) DEFAULT '0' COMMENT 'reversos del proceso',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`lote`,`proceso`),
  KEY `NewIndex2` (`lote`,`ruta_secuencia`)
) ENGINE=MyISAM AUTO_INCREMENT=56587 DEFAULT CHARSET=latin1 COMMENT='Histórico de lotes';

/*Data for the table `lotes_historia` */

/*Table structure for table `mapas` */

DROP TABLE IF EXISTS `mapas`;

CREATE TABLE `mapas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `planta` bigint(20) DEFAULT NULL COMMENT 'ID de la planta',
  `descripcion` varchar(200) NOT NULL,
  `ancho` float NOT NULL DEFAULT '0',
  `alto` float NOT NULL DEFAULT '0',
  `tasa_actualizacion` int(11) NOT NULL DEFAULT '5',
  `tasa_refresco` int(11) NOT NULL DEFAULT '600',
  `activo` tinyint(4) NOT NULL DEFAULT '0',
  `nombre` varchar(50) NOT NULL DEFAULT '',
  PRIMARY KEY (`id`),
  UNIQUE KEY `descripcion` (`descripcion`)
) ENGINE=MyISAM AUTO_INCREMENT=126 DEFAULT CHARSET=utf8;

/*Data for the table `mapas` */

insert  into `mapas`(`id`,`planta`,`descripcion`,`ancho`,`alto`,`tasa_actualizacion`,`tasa_refresco`,`activo`,`nombre`) values (124,NULL,'C:\\Users\\Dell E7490\\Documents\\mapas\\tenneco2\\mapa-256',1280,720,5,600,1,'');

/*Table structure for table `maquinas` */

DROP TABLE IF EXISTS `maquinas`;

CREATE TABLE `maquinas` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `id_wip` bigint(11) DEFAULT NULL,
  `celula` bigint(11) DEFAULT '0' COMMENT 'ID de la célula',
  `referencia` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Referencia',
  `nombre` varchar(100) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Nombre de la máquina',
  `tipo` bigint(11) DEFAULT '0' COMMENT 'ID del tipo de máquina',
  `clasificacion1` bigint(11) DEFAULT '0' COMMENT 'ID de la clasificación 1',
  `clasificacion2` bigint(11) DEFAULT '0' COMMENT 'ID de la clasificación 2',
  `imagen` varchar(250) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Imágen a utilizar',
  `colorhexa` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Color hexadecimal',
  `prefijo` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Nombre corto',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en pantalla',
  `estatus` char(1) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT '0' COMMENT 'Usuario que creó el registro',
  `modificado` bigint(11) DEFAULT '0' COMMENT 'Usuario que modificó el registro',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=1000 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de máquinas';

/*Data for the table `maquinas` */

/*Table structure for table `mensajes` */

DROP TABLE IF EXISTS `mensajes`;

CREATE TABLE `mensajes` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `alerta` bigint(20) DEFAULT '0' COMMENT 'ID de la alerta',
  `canal` int(2) DEFAULT '0' COMMENT 'Canal de la alerta (0=Llamada, 1=SMS, 2=Correo, 3=MMcall)',
  `tipo` int(2) DEFAULT '0' COMMENT '0=Inicio, 1-5=Escalación1, 9=Repetición, 11-15=Repetición de escalamiento',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `alarma` bigint(20) DEFAULT '0' COMMENT 'ID de la alarma',
  `prioridad` int(1) DEFAULT '0' COMMENT 'Prioridad del mensaje',
  `estatus` varchar(20) DEFAULT 'A' COMMENT 'Estatus del mensaje',
  `lista` bigint(20) DEFAULT '0' COMMENT 'ID del Recipiente',
  `creada` timestamp NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Creación del mensaje',
  `enviada` datetime DEFAULT NULL COMMENT 'Fecha de envío del mensaje',
  `texto` varchar(20) DEFAULT '' COMMENT 'Mensaje',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`prioridad`,`estatus`),
  KEY `NewIndex2` (`estatus`),
  KEY `NewIndex3` (`canal`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=673766 DEFAULT CHARSET=latin1;

/*Data for the table `mensajes` */

/*Table structure for table `mensajes_procesados` */

DROP TABLE IF EXISTS `mensajes_procesados`;

CREATE TABLE `mensajes_procesados` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del mensaje',
  `texto` varchar(400) DEFAULT NULL COMMENT 'Texto a enviar o dictar',
  `canal` int(2) DEFAULT '0' COMMENT 'Canal de difusión',
  `titulo` varchar(100) DEFAULT '0' COMMENT '(Para correo)',
  `prioridad` int(2) DEFAULT '0' COMMENT 'Proridad del mensaje',
  `fecha` date DEFAULT NULL COMMENT 'Fecha del mensaje (solo dura un día)',
  `mensaje` decimal(20,0) DEFAULT '0' COMMENT 'ID del mensaje',
  `estatus` char(20) DEFAULT 'A' COMMENT 'Estatus del envío',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`id`,`canal`,`prioridad`),
  KEY `NewIndex2` (`fecha`),
  KEY `NewIndex3` (`mensaje`)
) ENGINE=MyISAM AUTO_INCREMENT=592984 DEFAULT CHARSET=latin1 COMMENT='Mensajes a enviar o llamar';

/*Data for the table `mensajes_procesados` */

/*Table structure for table `objetos` */

DROP TABLE IF EXISTS `objetos`;

CREATE TABLE `objetos` (
  `mapa_id` int(11) NOT NULL,
  `id` varchar(50) NOT NULL,
  `descripcion` varchar(300) NOT NULL,
  `ultima_actualizacion` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`mapa_id`,`id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

/*Data for the table `objetos` */

insert  into `objetos`(`mapa_id`,`id`,`descripcion`,`ultima_actualizacion`) values (83,'1','Objeto 1','2019-09-22 00:57:45'),(83,'2','Objeto 2','2019-09-22 00:57:45'),(83,'3','Objeto 3','2019-09-22 00:57:45'),(84,'2','Objeto 2','2019-09-22 00:57:46'),(85,'2','Objeto 2','2019-09-22 01:57:42'),(85,'3','Objeto 3','2019-09-22 01:57:42'),(85,'1','Objeto 1','2019-09-22 01:57:42'),(86,'2','Objeto 2','2019-09-22 01:57:43'),(87,'1','Objeto 1','2019-11-04 07:05:24'),(87,'2','Objeto 2','2019-11-04 07:05:24'),(87,'3','Objeto 3','2019-11-04 07:05:24'),(87,'4','Objeto 4','2019-11-04 07:05:24'),(87,'5','Objeto 5','2019-11-04 07:05:24'),(87,'6','Objeto 6','2019-11-04 07:05:24'),(89,'1','Objeto 1','2019-11-04 07:05:27'),(89,'2','Objeto 2','2019-11-04 07:05:27'),(89,'4','Objeto 4','2019-11-04 07:05:27'),(89,'5','Objeto 5','2019-11-04 07:05:27'),(89,'3','Objeto 3','2019-11-04 07:05:27'),(89,'6','Objeto 6','2019-11-04 07:05:27'),(90,'1','Objeto 1','2019-11-04 07:07:36'),(90,'2','Objeto 2','2019-11-04 07:07:36'),(90,'4','Objeto 4','2019-11-04 07:07:36'),(90,'5','Objeto 5','2019-11-04 07:07:36'),(90,'3','Objeto 3','2019-11-04 07:07:36'),(90,'6','Objeto 6','2019-11-04 07:07:36'),(91,'1','Objeto 1','2019-11-04 07:07:40'),(91,'2','Objeto 2','2019-11-04 07:07:40'),(91,'4','Objeto 4','2019-11-04 07:07:40'),(91,'5','Objeto 5','2019-11-04 07:07:40'),(91,'3','Objeto 3','2019-11-04 07:07:40'),(91,'6','Objeto 6','2019-11-04 07:07:40'),(92,'1','Objeto 1','2019-11-04 10:59:41'),(92,'2','Objeto 2','2019-11-04 10:59:41'),(92,'4','Objeto 4','2019-11-04 10:59:41'),(92,'5','Objeto 5','2019-11-04 10:59:41'),(92,'3','Objeto 3','2019-11-04 10:59:41'),(92,'6','Objeto 6','2019-11-04 10:59:41'),(93,'1','Objeto 1','2019-11-04 10:59:45'),(93,'2','Objeto 2','2019-11-04 10:59:45'),(93,'4','Objeto 4','2019-11-04 10:59:45'),(93,'5','Objeto 5','2019-11-04 10:59:45'),(93,'3','Objeto 3','2019-11-04 10:59:45'),(93,'6','Objeto 6','2019-11-04 10:59:45'),(94,'1','Objeto 1','2019-11-04 10:54:05'),(94,'2','Objeto 2','2019-11-04 10:54:05'),(94,'4','Objeto 4','2019-11-04 10:54:05'),(94,'5','Objeto 5','2019-11-04 10:54:05'),(94,'3','Objeto 3','2019-11-04 10:54:05'),(94,'6','Objeto 6','2019-11-04 10:54:05'),(95,'1','Objeto 1','2019-11-04 10:54:07'),(95,'2','Objeto 2','2019-11-04 10:54:07'),(95,'4','Objeto 4','2019-11-04 10:54:07'),(95,'5','Objeto 5','2019-11-04 10:54:07'),(95,'3','Objeto 3','2019-11-04 10:54:07'),(95,'6','Objeto 6','2019-11-04 10:54:07'),(96,'1','Objeto 1','2019-11-08 03:01:02'),(96,'2','Objeto 2','2019-11-08 03:01:02'),(96,'4','Objeto 4','2019-11-08 03:01:02'),(96,'5','Objeto 5','2019-11-08 03:01:02'),(96,'3','Objeto 3','2019-11-08 03:01:02'),(96,'6','Objeto 6','2019-11-08 03:01:02'),(97,'1','Objeto 1','2019-11-04 20:45:19'),(97,'2','Objeto 2','2019-11-04 20:45:19'),(97,'4','Objeto 4','2019-11-04 20:45:19'),(97,'5','Objeto 5','2019-11-04 20:45:19'),(97,'3','Objeto 3','2019-11-04 20:45:19'),(97,'6','Objeto 6','2019-11-04 20:45:19'),(98,'1','Objeto 1','2019-11-16 22:17:03'),(98,'2','Objeto 2','2019-11-16 22:17:03'),(98,'3','Objeto 3','2019-11-16 22:17:03'),(98,'4','Objeto 4','2019-11-16 22:17:03'),(98,'5','Objeto 5','2019-11-16 22:17:03'),(98,'6','Objeto 6','2019-11-16 22:17:03'),(99,'1','Objeto 1','2019-11-22 16:59:47'),(99,'2','Objeto 2','2019-11-22 16:59:47'),(99,'3','Objeto 3','2019-11-22 16:59:47'),(99,'4','Objeto 4','2019-11-22 16:59:47'),(99,'5','Objeto 5','2019-11-22 16:59:47'),(99,'6','Objeto 6','2019-11-22 16:59:47'),(100,'7','Objeto 7','2019-11-22 16:59:48'),(100,'8','Objeto 8','2019-11-22 16:59:48'),(100,'9','Objeto 9','2019-11-22 16:59:48'),(100,'11','Objeto 11','2019-11-22 16:59:48'),(100,'10','Objeto 10','2019-11-22 16:59:48'),(100,'12','Objeto 12','2019-11-22 16:59:48'),(103,'2','Objeto 2','2019-11-28 06:45:41'),(103,'3','Objeto 3','2019-11-28 06:45:41'),(103,'1','Objeto 1','2019-12-02 13:24:49'),(104,'1','Objeto 1','2019-11-30 21:53:28'),(104,'2','Objeto 2','2019-11-30 21:53:28'),(104,'3','Objeto 3','2019-11-30 21:53:28'),(104,'4','Objeto 4','2019-11-30 21:53:28'),(104,'5','Objeto 5','2019-11-30 21:53:28'),(104,'6','Objeto 6','2019-11-30 21:53:28'),(105,'7','Objeto 7','2019-11-30 21:53:30'),(105,'8','Objeto 8','2019-11-30 21:53:30'),(105,'9','Objeto 9','2019-11-30 21:53:30'),(105,'11','Objeto 11','2019-11-30 21:53:30'),(105,'10','Objeto 10','2019-11-30 21:53:30'),(105,'12','Objeto 12','2019-11-30 21:53:30'),(106,'1','Objeto 1','2019-12-02 13:25:12'),(107,'1','Objeto 1','2019-12-13 13:25:56'),(107,'2','Objeto 2','2019-12-13 13:25:56'),(107,'3','Objeto 3','2019-12-13 13:25:56'),(107,'4','Objeto 4','2019-12-13 13:25:56'),(107,'5','Objeto 5','2019-12-13 13:25:56'),(107,'6','Objeto 6','2019-12-13 13:25:56'),(108,'7','Objeto 7','2019-12-13 13:25:57'),(108,'8','Objeto 8','2019-12-13 13:25:57'),(108,'9','Objeto 9','2019-12-13 13:25:57'),(108,'11','Objeto 11','2019-12-13 13:25:57'),(108,'10','Objeto 10','2019-12-13 13:25:57'),(108,'12','Objeto 12','2019-12-13 13:25:57'),(109,'1','Objeto 1','2019-12-13 13:39:20'),(109,'2','Objeto 2','2019-12-13 13:39:20'),(109,'3','Objeto 3','2019-12-13 13:39:20'),(109,'4','Objeto 4','2019-12-13 13:39:20'),(109,'5','Objeto 5','2019-12-13 13:39:20'),(109,'6','Objeto 6','2019-12-13 13:39:20'),(110,'7','Objeto 7','2019-12-13 13:39:22'),(110,'8','Objeto 8','2019-12-13 13:39:22'),(110,'9','Objeto 9','2019-12-13 13:39:22'),(110,'11','Objeto 11','2019-12-13 13:39:22'),(110,'10','Objeto 10','2019-12-13 13:39:22'),(110,'12','Objeto 12','2019-12-13 13:39:22'),(111,'1','Objeto 1','2019-12-13 13:45:58'),(113,'1','Objeto 1','2020-07-22 20:51:22'),(114,'1','Objeto 1','2020-07-22 20:55:47'),(115,'1','Objeto 1','2020-07-22 23:57:47'),(115,'2','Objeto 2','2020-07-22 23:57:47'),(116,'1','Objeto 1','2020-07-23 00:02:56'),(116,'2','Objeto 2','2020-07-23 00:02:56'),(117,'1','Objeto 1','2020-07-23 00:42:32'),(117,'2','Objeto 2','2020-07-23 00:42:32'),(118,'1','Objeto 1','2020-07-23 00:43:51'),(118,'2','Objeto 2','2020-07-23 00:43:51'),(119,'1','Objeto 1','2020-07-23 00:44:56'),(119,'2','Objeto 2','2020-07-23 00:44:56'),(120,'1','Objeto 1','2020-07-23 00:54:21'),(120,'2','Objeto 2','2020-07-23 00:54:21'),(121,'1','Objeto 1','2020-07-23 00:56:11'),(121,'2','Objeto 2','2020-07-23 00:56:11'),(122,'1','Objeto 1','2020-07-23 01:22:34'),(122,'2','Objeto 2','2020-07-23 01:22:34'),(123,'1','Objeto 1','2020-07-23 01:23:09'),(123,'2','Objeto 2','2020-07-23 01:23:09'),(124,'3','Objeto 3','2020-07-28 21:06:17'),(124,'44','Objeto 44','2020-07-28 21:06:17'),(124,'43','Objeto 43','2020-07-28 21:06:17'),(124,'4','Objeto 4','2020-07-28 21:06:17'),(124,'5','Objeto 5','2020-07-28 21:06:17'),(124,'6','Objeto 6','2020-07-28 21:06:17'),(124,'12','Objeto 12','2020-07-28 21:06:17'),(124,'11','Objeto 11','2020-07-28 21:06:17'),(124,'8','Objeto 8','2020-07-28 21:06:17'),(124,'7','Objeto 7','2020-07-28 21:06:17'),(124,'9','Objeto 9','2020-07-28 21:06:17'),(124,'10','Objeto 10','2020-07-28 21:06:17'),(124,'16','Objeto 16','2020-07-28 21:06:17'),(124,'15','Objeto 15','2020-07-28 21:06:17'),(124,'13','Objeto 13','2020-07-28 21:06:17'),(124,'14','Objeto 14','2020-07-28 21:06:17'),(124,'25','Objeto 25','2020-07-28 21:06:17'),(124,'26','Objeto 26','2020-07-28 21:06:17'),(124,'28','Objeto 28','2020-07-28 21:06:17'),(124,'27','Objeto 27','2020-07-28 21:06:17'),(124,'49','Objeto 49','2020-07-28 21:06:17'),(124,'1','Objeto 1','2020-07-28 21:06:17'),(124,'2','Objeto 2','2020-07-28 21:06:17'),(124,'33','Objeto 33','2020-07-28 21:06:17'),(124,'35','Objeto 35','2020-07-28 21:06:17'),(124,'34','Objeto 34','2020-07-28 21:06:17'),(124,'45','Objeto 45','2020-07-28 21:06:17'),(124,'46','Objeto 46','2020-07-28 21:06:17'),(124,'31','Objeto 31','2020-07-28 21:06:17'),(124,'40','Objeto 40','2020-07-28 21:06:17'),(124,'30','Objeto 30','2020-07-28 21:06:17'),(124,'29','Objeto 29','2020-07-28 21:06:17'),(124,'41','Objeto 41','2020-07-28 21:06:17'),(124,'32','Objeto 32','2020-07-28 21:06:17'),(124,'19','Objeto 19','2020-07-28 21:06:17'),(124,'21','Objeto 21','2020-07-28 21:06:17'),(124,'20','Objeto 20','2020-07-28 21:06:17'),(124,'24','Objeto 24','2020-07-28 21:06:17'),(124,'23','Objeto 23','2020-07-28 21:06:17'),(124,'22','Objeto 22','2020-07-28 21:06:17'),(124,'42','Objeto 42','2020-07-28 21:06:17'),(124,'48','Objeto 48','2020-07-28 21:06:17'),(124,'47','Objeto 47','2020-07-28 21:06:17'),(124,'39','Objeto 39','2020-07-28 21:06:17'),(124,'36','Objeto 36','2020-07-28 21:06:17'),(124,'37','Objeto 37','2020-07-28 21:06:17'),(124,'18','Objeto 18','2020-07-28 21:06:17'),(124,'17','Objeto 17','2020-07-28 21:06:17');

/*Table structure for table `paros` */

DROP TABLE IF EXISTS `paros`;

CREATE TABLE `paros` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `paro` bigint(20) DEFAULT '0' COMMENT 'ID del paro',
  `automatico` char(1) DEFAULT NULL COMMENT 'El paro fue automático',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `tripulacion` bigint(20) DEFAULT '0' COMMENT 'ID de la tripulación',
  `orden` bigint(20) DEFAULT NULL COMMENT 'ID de la orden de producción',
  `parte` bigint(20) DEFAULT NULL COMMENT 'ID del material',
  `inicia` datetime DEFAULT NULL COMMENT 'Fecha de inicio',
  `finaliza` datetime DEFAULT NULL COMMENT 'Fecha de fin',
  `tiempo_plan` bigint(20) DEFAULT NULL COMMENT 'Tiempo total en segundos (plan)',
  `tiempo_real` bigint(20) DEFAULT NULL COMMENT 'Tiempo total en segundos (real)',
  `adelantado` char(1) DEFAULT NULL COMMENT 'El paro se adelantó',
  `cancelado` char(1) DEFAULT NULL COMMENT 'El paro se canceló',
  `terminado_por` bigint(11) DEFAULT NULL COMMENT 'Usuario que terminó el paro',
  `terminado_causa` bigint(11) DEFAULT NULL COMMENT 'Causa de la cancelación',
  `terminado_fecha` datetime DEFAULT NULL COMMENT 'Fecha de la cancelación',
  `genero_otro` char(2) DEFAULT NULL COMMENT 'Generó otro paro (N/P)',
  `comentarios` varchar(1000) DEFAULT NULL COMMENT 'Comentarios del sistema',
  `atendido` char(1) DEFAULT NULL COMMENT 'El paro fue atendido?',
  `clasificado_por` bigint(20) DEFAULT NULL COMMENT 'Usuario que clasificó el paro',
  `clasificado_fecha` datetime DEFAULT NULL COMMENT 'Fecha de la clasificación',
  `cambiado` char(1) DEFAULT NULL COMMENT 'El paro fue cambiado de N/P a P',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Paros sucedidos';

/*Data for the table `paros` */

/*Table structure for table `piezasxminuto` */

DROP TABLE IF EXISTS `piezasxminuto`;

CREATE TABLE `piezasxminuto` (
  `equipo` bigint(20) DEFAULT NULL,
  `fecha` datetime DEFAULT NULL,
  `produccion` decimal(20,5) DEFAULT NULL,
  `paro` bigint(4) DEFAULT NULL,
  `hhmm` char(4) DEFAULT NULL,
  KEY `NewIndex1` (`equipo`,`fecha`),
  KEY `NewIndex2` (`equipo`,`hhmm`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Tabla de registros por segundo';

/*Data for the table `piezasxminuto` */

/*Table structure for table `plan_checklists` */

DROP TABLE IF EXISTS `plan_checklists`;

CREATE TABLE `plan_checklists` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `secuencia` int(6) DEFAULT NULL COMMENT 'Secuencia en la pantalla',
  `nombre` varchar(60) DEFAULT NULL COMMENT 'Nombre/Descripción del registro',
  `referencia` varchar(30) DEFAULT NULL COMMENT 'ID para otros sistemas',
  `calendario` bigint(20) DEFAULT '0' COMMENT 'ID del calendario',
  `imagen` varchar(255) DEFAULT NULL COMMENT 'Imagen a mostrar',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha y hora (fijo)',
  `maximo` int(6) DEFAULT '0',
  `reiniciar_cada` int(2) DEFAULT '0',
  `frecuencia` varchar(2) DEFAULT NULL COMMENT 'Frecuencia del checklist',
  `hora` time DEFAULT NULL COMMENT 'Horario del checklist',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas del registro',
  `checklists` char(1) DEFAULT 'S',
  `anticipacion` char(1) DEFAULT 'N',
  `anticipado` date DEFAULT NULL COMMENT 'Fecha de ejecucion',
  `tiempo` int(6) DEFAULT '0',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;

/*Data for the table `plan_checklists` */

/*Table structure for table `politicas` */

DROP TABLE IF EXISTS `politicas`;

CREATE TABLE `politicas` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `nombre` varchar(30) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Nombre corto',
  `deunsolouso` char(1) COLLATE utf8_bin DEFAULT NULL COMMENT 'De un sólo uso',
  `obligatoria` char(1) COLLATE utf8_bin DEFAULT NULL COMMENT 'Se requiere contraseña?',
  `vence` char(1) CHARACTER SET latin1 DEFAULT NULL COMMENT 'La contraseña vence?',
  `diasvencimiento` bigint(7) DEFAULT NULL COMMENT 'Días para vencerse',
  `aviso` bigint(7) DEFAULT NULL COMMENT 'Días de anticipación para avisar',
  `complejidad` char(10) COLLATE utf8_bin DEFAULT NULL COMMENT 'Complejidad de la contraseña (1-2=largo,3=especiales,4=numeros,5=mayus/minus)',
  `usadas` int(2) DEFAULT NULL COMMENT 'Últimas contraseñas usadas',
  `caducidad` int(3) DEFAULT NULL COMMENT 'Días de gracia',
  `estatus` char(1) CHARACTER SET latin1 DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT NULL COMMENT 'Usuario que creó el registro',
  `modificado` bigint(20) DEFAULT NULL COMMENT 'Usuario que modificó el registro',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=6 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de políticas';

/*Data for the table `politicas` */

insert  into `politicas`(`id`,`nombre`,`deunsolouso`,`obligatoria`,`vence`,`diasvencimiento`,`aviso`,`complejidad`,`usadas`,`caducidad`,`estatus`,`creacion`,`modificacion`,`creado`,`modificado`) values (1,'POLITICA GENERAL','N','S','S',365,30,'4;N;N;N',5,0,'A','2020-02-05 15:00:07','2020-02-05 13:00:07',1,1),(5,'OPERADORES','N','N','N',0,0,'0;N;N;N',0,NULL,'A','2020-01-21 21:11:33','2020-01-21 13:11:33',1,1);

/*Table structure for table `prioridades` */

DROP TABLE IF EXISTS `prioridades`;

CREATE TABLE `prioridades` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `fecha` datetime DEFAULT NULL COMMENT 'Fecha y hora de programación',
  `parte` bigint(20) DEFAULT NULL COMMENT 'Número de parte',
  `orden` int(3) DEFAULT NULL COMMENT 'Prioridad',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID del proceso',
  `notas` varchar(300) DEFAULT NULL COMMENT 'Notas varias',
  `estatus` char(1) DEFAULT NULL COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`parte`,`fecha`,`estatus`)
) ENGINE=MyISAM AUTO_INCREMENT=94 DEFAULT CHARSET=latin1;

/*Data for the table `prioridades` */

/*Table structure for table `programacion` */

DROP TABLE IF EXISTS `programacion`;

CREATE TABLE `programacion` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `carga` bigint(20) DEFAULT NULL COMMENT 'ID de la carga',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `cantidad` bigint(12) DEFAULT '0' COMMENT 'Cantidad',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(11) DEFAULT NULL COMMENT 'Creado por',
  `modificado` bigint(11) DEFAULT NULL COMMENT 'Modificado por',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=106 DEFAULT CHARSET=latin1 COMMENT='Progranación de la producción';

/*Data for the table `programacion` */

/*Table structure for table `pu_colores` */

DROP TABLE IF EXISTS `pu_colores`;

CREATE TABLE `pu_colores` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del tema',
  `planta` bigint(20) DEFAULT '0' COMMENT 'ID de la planta',
  `usuario` bigint(20) DEFAULT '0' COMMENT 'ID del usuario',
  `nombre` varchar(50) DEFAULT NULL COMMENT 'Nomre del tema',
  `personalizada` char(1) DEFAULT 'N' COMMENT 'Skin personalizada',
  `obligatorio` char(1) DEFAULT 'N' COMMENT 'Skin obligatorio',
  `seleccionado` char(1) DEFAULT 'N' COMMENT 'Skin seleccionado',
  `fondo_total` varchar(20) DEFAULT NULL COMMENT 'Color de fondo total',
  `fondo_barra_superior` varchar(20) DEFAULT NULL COMMENT 'Color de la barra superior',
  `fondo_barra_inferior` varchar(20) DEFAULT NULL COMMENT 'Color de la barra inferior',
  `fondo_aplicacion` varchar(20) DEFAULT NULL COMMENT 'Color de fondo',
  `fondo_seleccion` varchar(20) DEFAULT NULL COMMENT 'Color de fondo selección',
  `fondo_boton` varchar(20) DEFAULT NULL COMMENT 'Color de fondo botón ',
  `fondo_slider` varchar(20) DEFAULT NULL COMMENT 'Color de fondo slider',
  `fondo_tarjeta` varchar(20) DEFAULT NULL COMMENT 'Color de fondo de la tarjeta',
  `fondo_boton_inactivo` varchar(20) DEFAULT NULL COMMENT 'Color de fondo botón inactivo',
  `fondo_boton_positivo` varchar(20) DEFAULT NULL COMMENT 'Color de fondo botón positivo',
  `fondo_boton_negativo` varchar(20) DEFAULT NULL COMMENT 'Color de fondo botón negativo',
  `fondo_boton_barra` varchar(20) DEFAULT NULL COMMENT 'Color de fondo botón en barra',
  `fondo_tiptool` varchar(20) DEFAULT NULL COMMENT 'Color de fondo tiptool',
  `fondo_logo` varchar(20) DEFAULT NULL COMMENT 'Color de fondo logo',
  `fondo_snack_normal` varchar(20) DEFAULT NULL COMMENT 'Color de fondo snack normal',
  `fondo_snack_error` varchar(20) DEFAULT NULL COMMENT 'Color de fondo snack error',
  `borde_total` varchar(20) DEFAULT NULL COMMENT 'Color de borde',
  `borde_seleccion` varchar(20) DEFAULT NULL COMMENT 'Color de borde en selección',
  `borde_hover` varchar(20) DEFAULT NULL COMMENT 'Color de borde en Hover',
  `borde_boton` varchar(20) DEFAULT NULL COMMENT 'Color de borde de botones',
  `borde_boton_inactivo` varchar(20) DEFAULT NULL COMMENT 'Color de borde de inactivo',
  `borde_tarjeta` varchar(20) DEFAULT NULL COMMENT 'Color de borde de la tarjeta',
  `borde_tiptool` varchar(20) DEFAULT NULL COMMENT 'Color de borde del tiptool',
  `color_impar` varchar(20) DEFAULT NULL COMMENT 'Color impar',
  `color_par` varchar(20) DEFAULT NULL COMMENT 'Color par',
  `texto_tarjeta` varchar(20) DEFAULT NULL COMMENT 'Texto tarjeta',
  `texto_tarjeta_resalte` varchar(20) DEFAULT NULL COMMENT 'Texto tarjeta resaltado',
  `texto_barra_superior` varchar(20) DEFAULT NULL COMMENT 'Texto barra superior',
  `texto_barra_inferior` varchar(20) DEFAULT NULL COMMENT 'Texto barra inferior',
  `texto_boton` varchar(20) DEFAULT NULL COMMENT 'Texto botón',
  `texto_boton_inactivo` varchar(20) DEFAULT NULL COMMENT 'Texto botón inactivo',
  `texto_boton_positivo` varchar(20) DEFAULT NULL COMMENT 'Texto botón positivo',
  `texto_boton_negativo` varchar(20) DEFAULT NULL COMMENT 'Texto botón negativo',
  `texto_boton_barra` varchar(20) DEFAULT NULL COMMENT 'Texto botón barra',
  `texto_seleccion` varchar(20) DEFAULT NULL COMMENT 'Texto selección',
  `texto_tiptool` varchar(20) DEFAULT NULL COMMENT 'Texto tiptool',
  `texto_snack_normal` varchar(20) DEFAULT NULL COMMENT 'Texto snack normal',
  `texto_snack_error` varchar(20) DEFAULT NULL COMMENT 'Texto snack error',
  `texto_solo_texto` varchar(20) DEFAULT NULL COMMENT 'Texto solo lectura',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=latin1 COMMENT='Colores';

/*Data for the table `pu_colores` */

insert  into `pu_colores`(`id`,`planta`,`usuario`,`nombre`,`personalizada`,`obligatorio`,`seleccionado`,`fondo_total`,`fondo_barra_superior`,`fondo_barra_inferior`,`fondo_aplicacion`,`fondo_seleccion`,`fondo_boton`,`fondo_slider`,`fondo_tarjeta`,`fondo_boton_inactivo`,`fondo_boton_positivo`,`fondo_boton_negativo`,`fondo_boton_barra`,`fondo_tiptool`,`fondo_logo`,`fondo_snack_normal`,`fondo_snack_error`,`borde_total`,`borde_seleccion`,`borde_hover`,`borde_boton`,`borde_boton_inactivo`,`borde_tarjeta`,`borde_tiptool`,`color_impar`,`color_par`,`texto_tarjeta`,`texto_tarjeta_resalte`,`texto_barra_superior`,`texto_barra_inferior`,`texto_boton`,`texto_boton_inactivo`,`texto_boton_positivo`,`texto_boton_negativo`,`texto_boton_barra`,`texto_seleccion`,`texto_tiptool`,`texto_snack_normal`,`texto_snack_error`,`texto_solo_texto`,`estatus`) values (1,0,1,'DARK','N','N','N','000000','404040','404040','303030','101010','353535','151515','252525','t','9ACD32','FF4500','ffffff','FFFFFF','ffffff','303030','FF6347','FFFFFF','DDDDDD','707070','BBBBBB','404040','606060','FFFFFF','t','303030','CCCCCC','FFFFFF','FFFFFF','FFFFFF','CCCCCC','404040','000000','FFFFFF','151515','FFFFFF','000000','FFFFFF','FFFFFF','909090','A'),(2,0,0,'AZUL CORPORATIVO','N','S','N','002369','f6f6f6','c9c7c7','FFFFFF','73ccf0','f0eded','cfd0d1','FFFFFF','FFFFFF','58D68D','ff0000','002369','FFFFFF','FFFFFF','d7eefc','FF6347','64718c','babdbf','73ccf0','4bc8fa','d7dadb','64718c','000000','EDFBFF','ffffff','002369','002369','002369','002369','002369','d7dadb','000000','ffffff','FFFFFF','002369','002369','002369','FFFFFF','909090','A'),(3,0,0,'AMARILLO M/L','N','N','N','2D3277','ffe600','ffe600','FFFFFF','ffe600','fff9c7','fcfbeb','fffef5','FFFFFF','ffe600','ff0000','ffffff','fff9c7','FFFFFF','fff9c7','FF0000','64718c','babdbf','73ccf0','4bc8fa','c9c17b','64718c','000000','EDFBFF','fffef5','2D3277','002369','2D3277','2D3277','2D3277','c9c17b','000000','ffffff','002369','002369','002369','002369','FFFFFF','a8a694','A'),(4,0,0,'HEB','N','N','N','a33334','de2b27','de2b27','FFFFFF','ffc4c5','f0eded','fa6467','FFFFFF','FFFFFF','58D68D','ff0000','FFFFFF','FFFFFF','FFFFFF','d7eefc','FF6347','64718c','babdbf','f56264','fa6e70','d7dadb','64718c','000000','fff2f2','ffffff','7a0606','7a0606','ffffff','ffffff','a33334','d7dadb','000000','ffffff','de2b27','a33334','7a0606','7a0606','FFFFFF','909090','A');

/*Table structure for table `pu_graficos` */

DROP TABLE IF EXISTS `pu_graficos`;

CREATE TABLE `pu_graficos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `usuario` bigint(11) DEFAULT NULL COMMENT 'ID del usuario',
  `visualizar` char(1) DEFAULT 'S' COMMENT 'Visualizar gráfico',
  `grafico` int(4) DEFAULT NULL COMMENT 'Número del gráfico',
  `titulo` varchar(100) DEFAULT NULL COMMENT 'Título del gráfico',
  `titulo_fuente` int(2) DEFAULT NULL COMMENT 'Fuente del título',
  `sub_titulo` varchar(100) DEFAULT NULL COMMENT 'Subtítulo del gráfico',
  `subtitulo_fuente` int(2) DEFAULT NULL COMMENT 'Fuente del subtítulo',
  `texto_x` varchar(100) DEFAULT NULL COMMENT 'Texto eje X',
  `texto_x_fuente` int(2) DEFAULT NULL COMMENT 'Fuente del eje X',
  `texto_y` varchar(100) DEFAULT '0' COMMENT 'Texto eje Y',
  `texto_y_fuente` int(2) DEFAULT NULL COMMENT 'Fuente del eje Y',
  `texto_z` varchar(100) DEFAULT '0' COMMENT 'Texto eje Z',
  `texto_z_fuente` int(2) DEFAULT NULL COMMENT 'Fuente del eje Z',
  `etiqueta_mostrar` char(1) DEFAULT NULL COMMENT 'Mostrar etiquetas',
  `etiqueta_fuente` int(2) DEFAULT NULL COMMENT 'Fuente de la etiqueta',
  `etiqueta_leyenda` varchar(30) DEFAULT NULL COMMENT 'Título de la leyenda',
  `etiqueta_color` varchar(20) DEFAULT NULL COMMENT 'Color de la etiqueta',
  `etiqueta_fondo` varchar(20) DEFAULT NULL COMMENT 'Color del fondo',
  `ancho` int(6) DEFAULT NULL COMMENT 'Ancho de la pantalla',
  `alto` int(6) DEFAULT '0' COMMENT 'Alto de la pantalla',
  `margen_arriba` int(4) DEFAULT NULL COMMENT 'Margen arriba',
  `margen_abajo` int(4) DEFAULT NULL COMMENT 'Margen abajo',
  `margen_izquierda` int(4) DEFAULT NULL COMMENT 'Margen izquierda',
  `margen_derecha` int(4) DEFAULT NULL COMMENT 'Margen derecha',
  `maximo_barras` int(2) DEFAULT '0' COMMENT 'Máximo de barras',
  `maximo_barraspct` int(3) DEFAULT '0' COMMENT 'PCT de máximo de barras',
  `agrupar` char(1) DEFAULT NULL COMMENT 'Agrupar el resto',
  `agrupar_posicion` char(1) DEFAULT NULL COMMENT 'P=Principio, F=Final, N=Ordenado',
  `agrupar_texto` varchar(30) DEFAULT NULL COMMENT 'Texto de la barra agrupada',
  `fecha` datetime DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha de actualización',
  `periodo_tipo` int(2) DEFAULT NULL COMMENT 'Tipo de período (0: segundos, 1: minutos, 2: horas, 3: días, 4: semanas, 5: meses, 6: años, 10: DTD, 11: WTD, 12: MTD, 13: YTD)',
  `periodo_atras` bigint(8) DEFAULT NULL COMMENT 'Tiempo a recorrer hacía atrás',
  `mostrar_tabla` char(1) DEFAULT NULL COMMENT 'Mostrar tabla',
  `orden` char(1) DEFAULT NULL COMMENT 'Orden de la gráfica',
  `color_fondo_barras` varchar(20) DEFAULT NULL COMMENT 'Color de fondo de las barras',
  `color_letras` varchar(20) DEFAULT NULL COMMENT 'Color de las letras',
  `color_fondo` varchar(20) DEFAULT NULL COMMENT 'Color del fondo',
  `color_leyenda_fondo` varchar(20) DEFAULT NULL COMMENT 'Color del fondo de la leyenda',
  `color_leyenda` varchar(20) DEFAULT NULL COMMENT 'Color del texto de la leyenda',
  `ver_esperado` char(1) DEFAULT 'N' COMMENT 'Ver esperado',
  `grueso_esperado` int(2) DEFAULT NULL COMMENT 'Pixls del esperado',
  `color_esperado` varchar(20) DEFAULT NULL COMMENT 'Color del esperado',
  `texto_esperado` varchar(30) DEFAULT NULL COMMENT 'Texto para el valor esperado',
  `valor_esperado` decimal(20,3) DEFAULT '0.000' COMMENT 'Valor esperado',
  `incluir_ceros` char(1) DEFAULT 'N' COMMENT 'Incluir valores cero',
  `orden_grafica` char(1) DEFAULT 'N' COMMENT 'Orden de la gráfica (M=Mayor a menor, N= Menor a mayor, A=Alfabético)',
  `mostrarpct` char(1) DEFAULT 'N' COMMENT 'Mostrar frecuencia',
  `color_barra` varchar(20) DEFAULT NULL COMMENT 'Color de la barra',
  `color_barra_borde` varchar(20) DEFAULT NULL COMMENT 'Color de la barra borde',
  `color_barra_o` varchar(20) DEFAULT NULL COMMENT 'Color de la barra opcional',
  `color_barra_borde_o` varchar(20) DEFAULT NULL COMMENT 'Color de la barra borde opcional',
  `ver_leyenda` char(1) DEFAULT 'S' COMMENT 'Ver leyenda',
  `overlap` char(1) DEFAULT 'S' COMMENT 'Tipo de overlap',
  `adicionales` char(13) DEFAULT '0;0;0;0;0;0;0' COMMENT 'YTD, MTD, Año anterior, Mes anterior, Mismo mes Año anterior, Aplicar Filtro a los acumulados',
  `adicionales_colores` varchar(130) DEFAULT NULL COMMENT 'Colores',
  `adicionales_titulos` varchar(130) DEFAULT NULL,
  `oee` char(6) DEFAULT 'NNN',
  `oee_colores` varchar(80) DEFAULT NULL,
  `oee_tipo` char(3) DEFAULT 'BBB',
  `oee_nombre` varchar(80) DEFAULT ';;',
  `tipo_principal` char(1) DEFAULT 'B',
  `colores_multiples` char(1) DEFAULT NULL COMMENT 'Usar colores diferentes',
  `color_spiline` varchar(20) DEFAULT NULL COMMENT 'Color del spline',
  `grueso_spiline` int(2) DEFAULT '0' COMMENT 'Pixls del spline',
  `mostrar_argumentos` char(1) DEFAULT 'S' COMMENT 'Mostrar argumentos',
  `activo` char(1) DEFAULT 'S' COMMENT 'Gráfico activo para el sistema?',
  `esperado_esquema` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`grafico`)
) ENGINE=MyISAM AUTO_INCREMENT=208 DEFAULT CHARSET=latin1 COMMENT='Preferencias de usuario (Gráficos)';

/*Data for the table `pu_graficos` */

insert  into `pu_graficos`(`id`,`usuario`,`visualizar`,`grafico`,`titulo`,`titulo_fuente`,`sub_titulo`,`subtitulo_fuente`,`texto_x`,`texto_x_fuente`,`texto_y`,`texto_y_fuente`,`texto_z`,`texto_z_fuente`,`etiqueta_mostrar`,`etiqueta_fuente`,`etiqueta_leyenda`,`etiqueta_color`,`etiqueta_fondo`,`ancho`,`alto`,`margen_arriba`,`margen_abajo`,`margen_izquierda`,`margen_derecha`,`maximo_barras`,`maximo_barraspct`,`agrupar`,`agrupar_posicion`,`agrupar_texto`,`fecha`,`periodo_tipo`,`periodo_atras`,`mostrar_tabla`,`orden`,`color_fondo_barras`,`color_letras`,`color_fondo`,`color_leyenda_fondo`,`color_leyenda`,`ver_esperado`,`grueso_esperado`,`color_esperado`,`texto_esperado`,`valor_esperado`,`incluir_ceros`,`orden_grafica`,`mostrarpct`,`color_barra`,`color_barra_borde`,`color_barra_o`,`color_barra_borde_o`,`ver_leyenda`,`overlap`,`adicionales`,`adicionales_colores`,`adicionales_titulos`,`oee`,`oee_colores`,`oee_tipo`,`oee_nombre`,`tipo_principal`,`colores_multiples`,`color_spiline`,`grueso_spiline`,`mostrar_argumentos`,`activo`,`esperado_esquema`) values (63,0,'S',1,'ENTREGA DE LOTES POR SEMANA',25,NULL,20,'SEMANAS DEL AÑO',17,'NUMERO DE LOTES',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-13 01:12:46',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','N',NULL),(64,0,'S',1,'ENTREGA DE LOTES POR SEMANA',25,'Ver filtro',20,'Semanas',17,'NUMERO DE LOTES',17,'1;2;3;',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-13 01:45:44',NULL,NULL,'S',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','N',NULL),(66,0,'S',2,'ENTREGA DE LOTES POR OPERACIÓN',25,NULL,20,'OPERACIONES',17,'NUMERO DE LOTES',17,'1;2;3;',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-14 00:13:50',NULL,NULL,'S',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','N',NULL),(67,0,'S',2,'ENTREGA DE LOTES POR OPERACIÓN',25,NULL,20,NULL,17,'0',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-14 01:46:35',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','N',NULL),(68,0,'S',3,'PRODUCTIVIDAD POR OPERACIÓN',25,NULL,20,NULL,17,'0',17,'1;2;3;',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-14 01:46:44',NULL,NULL,'S',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','N',NULL),(71,0,'S',3,'PRODUCTIVIDAD POR OPERACIÓN',25,'operaciones',20,'OPERACIONES',17,'Lotes producidos',17,'1;2;3;Productividad (%)',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-09-14 02:35:16',NULL,NULL,'N',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','N',NULL),(72,0,'S',101,'MTTR POR CELULA',25,'EJEMPLO DEL SISTEMA',20,'CELULAS DE PRODUCCION',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,'0',NULL,NULL,0,0,0,0,0,0,10,0,'S','F','LINEA AGRUPADA','2019-10-20 09:44:45',NULL,NULL,NULL,NULL,'',NULL,'','','',NULL,NULL,NULL,NULL,'0.000','N','M','N',NULL,NULL,NULL,NULL,'N','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S','FF0000',0,'S','S',NULL),(73,0,'S',102,'MTTR POR MAQUINA',25,'EJEMPLO DEL SISTEMA',20,'MAQUINAS',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','MAQUINA AGRUPADA','2019-10-20 16:35:25',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','N',NULL,0,'S','S',NULL),(74,0,'S',103,'MTTR POR ÁREA',25,'EJEMPLO DEL SISTEMA',20,'AREAS',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','ÁREA AGRUPADA','2019-10-20 16:36:29',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','N',NULL,0,'S','S',NULL),(75,0,'S',104,'MTTR POR FALLA',25,'EJEMPLO DEL SISTEMA',20,'FALLAS',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,75,'S','O','FALLAS AGRUPADA','2019-10-20 16:36:30',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','M','N',NULL,NULL,NULL,NULL,'S','R','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(77,0,'S',105,'MTTR POR TIPO DE MAQUINA',25,'EJEMPLO DEL SISTEMA',20,'TIPOS DE MAQUINA',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','TIPOS DE MAQUINA','2019-10-21 19:23:52',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','N',NULL,0,'S','S',NULL),(78,0,'S',106,'MTTR POR AGRUPADOR (1) DE MAQUINA',25,'EJEMPLO DEL SISTEMA',20,'AGRUPADORES 1',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','AGRUPADOR 1','2019-10-22 00:43:42',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','N',NULL,0,'S','S',NULL),(79,0,'S',107,'MTTR POR AGRUPADOR (2) DE MAQUINA',25,'EJEMPLO DEL SISTEMA',20,'AGRUPADORES 2',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','AGRUPADOR 2','2019-10-22 00:43:43',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(80,0,'S',108,'MTTR POR AGRUPADOR (1) DE FALLAS',25,'EJEMPLO DEL SISTEMA',20,'AGRUPADORES 1',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,10,0,'S','F','AGRUPADOR 1','2019-10-22 00:43:45',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(81,0,'S',109,'MTTR POR AGRUPADOR (2) DE FALLAS',25,'EJEMPLO DEL SISTEMA',20,'AGRUPADORES 2',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'S','F','AGRUPADOR 2','2019-10-22 00:43:46',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(82,0,'N',110,'MTTR POR DIA',25,'EJEMPLO DEL SISTEMA',20,'DIA',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'S','F','DIA','2019-10-22 01:08:56',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(83,0,'N',111,'MTTR POR SEMANA',25,'EJEMPLO DEL SISTEMA',20,'SEMANA',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'S','F','SEMANA','2019-10-22 01:08:57',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(84,0,'N',112,'MTTR POR MES',25,'EJEMPLO DEL SISTEMA',20,'MES',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'S','F','MES','2019-10-22 01:08:58',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(85,0,'S',113,'MTTR POR TECNICO',25,'EJEMPLO DEL SISTEMA',20,'TECNICO',17,'MTTR Tiempo promedio de reparación/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'S','F','TECNICO','2019-10-22 01:08:59',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(86,0,'S',201,'MTBF POR LINEA',25,NULL,20,'CELULAS DE PRODUCCION',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:09',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(87,0,'S',202,'MTBF POR MAQUINA',25,NULL,20,'MAQUINAS',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(88,0,'S',203,'MTBF POR AREA',25,NULL,20,'AREAS',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:11',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(89,0,'S',204,'MTBF POR FALLA',25,NULL,20,'FALLAS',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:12',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(90,0,'S',205,'MTBF POR TIPO DE MAQUINA',25,NULL,20,'TIPOS DE MAQUINA',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:14',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(91,0,'S',206,'MTBF POR AGRUPADOR (1) DE MAQUINA',25,NULL,20,'AGRUPADORES 1',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:14',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(92,0,'S',207,'MTBF POR AGRUPADOR (2) DE MAQUINA',25,NULL,20,'AGRUPADORES 2',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:17',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(93,0,'S',208,'MTBF POR AGRUPADOR (1) DE FALLAS',25,NULL,20,'AGRUPADORES 1',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:18',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(94,0,'S',209,'MTBF POR AGRUPADOR (2) DE FALLAS',25,NULL,20,'AGRUPADORES 2',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:19',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(95,0,'S',210,'MTBF POR DIA',25,NULL,20,'DIA',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:20',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(96,0,'S',211,'MTBF POR SEMANA',25,NULL,20,'SEMANA',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:22',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(97,0,'S',212,'MTBF POR MES',25,NULL,20,'MES',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:24',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(98,0,'S',213,'MTBF POR TECNICO',25,NULL,20,'TECNICO',17,'MTBF (Tiempo promedio entre fallas/Horas)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 21:30:26',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(99,0,'S',301,'PARETO POR LINEA',25,NULL,20,'LINEA',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:56:41',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B','S','00ff00',5,'S','S',NULL),(100,0,'S',302,'PARETO POR MAQUINA',25,NULL,20,'MAQUINAS',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:46',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,'00ff00',2,'S','S',NULL),(101,0,'S',303,'PARETO POR AREA',25,NULL,20,'AREAS',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:47',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(102,0,'S',304,'PARETO POR FALLA',25,NULL,20,'FALLAS',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:48',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','R','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(103,0,'S',305,'PARETO POR TIPO DE MAQUINA',25,NULL,20,'TIPOS DE MAQUINA',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:49',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(104,0,'S',306,'PARETO POR AGRUPADOR (1) DE MAQUINA',25,NULL,20,'AGRUPADORES 1',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:50',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(105,0,'S',307,'PARETO POR AGRUPADOR (2) DE MAQUINA',25,NULL,20,'AGRUPADORES 2',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-21 23:59:52',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(106,0,'S',308,'PARETO POR AGRUPADOR (1) DE FALLAS',25,NULL,20,'AGRUPADORES 1',17,'NUMERO DE FALLAS',17,'PCT acumulado',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:00:31',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(107,0,'S',309,'PARETO POR AGRUPADOR (2) DE FALLAS',25,NULL,20,'AGRUPADORES 2',17,'NUMERO DE FALLAS',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:00:33',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(108,0,'S',310,'PARETO POR DIA',25,NULL,NULL,'DIAS',17,'NUMERO DE FALLAS',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:03:04',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(109,0,'S',311,'PARETO POR SEMANA',25,NULL,NULL,'SEMANAS',17,'NUMERO DE FALLAS',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:03:05',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(110,0,'S',312,'PARETO POR MES',25,NULL,NULL,'MESES',17,'NUMERO DE FALLAS',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:03:06',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(111,0,'S',313,'PARETO POR TECNICO',25,NULL,NULL,'TECNICOS',17,'NUMERO DE FALLAS',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2019-10-22 00:03:07',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,2,'S','S',NULL),(129,0,'S',401,'OAE POR LINEA',25,NULL,NULL,'LINEAS',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:52',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NSN','FFFFFF;FF0000;00FF00','BLL',';;','B',NULL,NULL,0,'S','S',NULL),(130,0,'S',402,'OAE POR MAQUINA/EQUIPO',25,NULL,NULL,'MAQUINAS',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:53',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(131,0,'S',403,'OAE POR TIPO DE MAQUINA',25,NULL,NULL,'TIPOS DE MAQUINAS',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:53',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(132,0,'S',404,'OAE POR AGRUPADOR (1) DE MAQUINA',25,NULL,NULL,'AGRUPADOR (1)',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:53',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(133,0,'S',405,'OAE POR AGRUPADOR (2) DE MAQUINA',25,NULL,NULL,'AGRUPADOR (2)',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:53',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(134,0,'S',406,'OAE POR DIA',25,NULL,NULL,'DIAS',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:54',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(135,0,'S',407,'OAE POR SEMANA',25,NULL,NULL,'SEMANAS',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:54',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(136,0,'S',408,'OAE POR MES',25,NULL,NULL,'MESES',17,'OEE (Overall Equipment Effectiveness)',17,'0',NULL,NULL,17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,NULL,NULL,NULL,'2020-01-30 06:37:54',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,'N',NULL,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','NNNNNN',NULL,NULL,'NNN',NULL,'BBB',';;','B',NULL,NULL,0,'S','S',NULL),(185,0,'S',507,'Pareto/paros por MES',25,NULL,20,'DIAS DEL AÑO',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:26:07',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(184,0,'S',506,'Pareto/paros por SEMANA',25,NULL,20,'SEMANAS DEL AÑO',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:25:10',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(183,0,'S',505,'Pareto/paros por DIA',25,NULL,20,'MESES DEL AÑO',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:25:03',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(182,0,'S',504,'Pareto/paros por CONCEPTO',25,NULL,20,'CONCEPTOS DE MANTENIMIENTO',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:24:47',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(181,0,'S',503,'Pareto/paros por AREA',25,NULL,20,'AREAS',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:24:46',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(180,0,'S',502,'Pareto/paros por MAQUINA',25,NULL,20,'MAQUINAS',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'N',17,NULL,NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-18 00:24:29',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(179,0,'S',501,'Pareto/paros por LINEA',25,NULL,20,'LINEAS/CELULAS',17,'NÚMERO DE FALLAS',17,'PCT acumulado',10,'S',17,'Fallas',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N',NULL,'2020-03-17 20:29:47',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,NULL,'0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(196,0,'S',605,'Rechazos por Dïa',25,NULL,20,'DIAS DEL AÑO',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:52',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(197,0,'S',606,'Rechazos por Semana',25,NULL,20,'SEMANAS DEL AÑO',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:54',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(198,0,'S',607,'Rechazos por Mes',25,NULL,20,'MESES DEL AÑO',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:56',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(199,0,'S',601,'Rechazos por Linea',25,NULL,20,'LINEAS/CELULAS',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:42',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(200,0,'S',602,'Rechazos por Máquina',25,NULL,20,'MAQUINA',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:44',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(201,0,'S',603,'Rechazos por Área',25,NULL,20,'AREAS',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:49',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL),(202,0,'S',604,'Rechazos por Concepto',25,NULL,20,'CONCEPTOS DE RECHAZO',NULL,'NÚMERO DE RECHAZOS',17,'PCT acumulado',10,'S',17,'Leyenda',NULL,NULL,NULL,0,NULL,NULL,NULL,NULL,0,0,'N','N','AGRUPADO','2020-05-03 22:42:50',NULL,NULL,NULL,'1',NULL,NULL,NULL,NULL,NULL,'N',1,NULL,'ESPERADO','0.000','N','N','N',NULL,NULL,NULL,NULL,'S','S','0;0;0;0;0;0;0',NULL,NULL,'NNN',NULL,'BBB',';;','B','S',NULL,0,'S','S',NULL);

/*Table structure for table `relacion_alertas_operaciones` */

DROP TABLE IF EXISTS `relacion_alertas_operaciones`;

CREATE TABLE `relacion_alertas_operaciones` (
  `alerta` bigint(20) DEFAULT NULL COMMENT 'ID de la alerta',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID del registro del catálogo',
  `tipo` int(2) DEFAULT NULL COMMENT 'Catálogo asociado',
  UNIQUE KEY `NewIndex1` (`alerta`,`proceso`,`tipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_alertas_operaciones` */

/*Table structure for table `relacion_fallas_operaciones` */

DROP TABLE IF EXISTS `relacion_fallas_operaciones`;

CREATE TABLE `relacion_fallas_operaciones` (
  `falla` bigint(20) DEFAULT NULL COMMENT 'ID de la falla',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID del registro del catálogo',
  `tipo` int(2) DEFAULT NULL COMMENT 'Catálogo asociado',
  UNIQUE KEY `NewIndex1` (`falla`,`proceso`,`tipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_fallas_operaciones` */

/*Table structure for table `relacion_maquinas_lecturas` */

DROP TABLE IF EXISTS `relacion_maquinas_lecturas`;

CREATE TABLE `relacion_maquinas_lecturas` (
  `equipo` bigint(20) NOT NULL DEFAULT '0' COMMENT 'ID del equipo',
  `orden` bigint(20) DEFAULT '0' COMMENT 'ID de la O/P',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `tripulacion` bigint(20) DEFAULT '0' COMMENT 'ID de la tripulación',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `nparte` varchar(100) DEFAULT NULL,
  `norden` varchar(30) DEFAULT NULL,
  `nequipo` varchar(100) DEFAULT NULL,
  `ntripulacion` varchar(100) DEFAULT NULL COMMENT 'Descipción de la tripulación',
  `nturno` varchar(100) DEFAULT NULL COMMENT 'Descipción del turno',
  `referencia` varchar(30) DEFAULT NULL,
  `rate_teorico` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Rate teórico',
  `objetivo` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Objetivo del corte',
  `produccion` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad de eventos producidos',
  `calidad` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad de eventos rechazados',
  `buffer` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Cantidad en buffer',
  `produccion_tc` decimal(25,10) DEFAULT '0.0000000000',
  `calidad_tc` decimal(25,10) DEFAULT '0.0000000000',
  `paros` bigint(20) DEFAULT '0' COMMENT 'Segundos de paro acumulado',
  `parosmostrar` bigint(20) DEFAULT '0' COMMENT 'Segundos de paro + Actual',
  `rate` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Rate actual',
  `rate_efecto` char(1) DEFAULT 'N' COMMENT 'Efecto del rate (B=Bajo, A=Alto, N=Normal)',
  `ftq_efecto` char(1) DEFAULT 'N' COMMENT 'Efecto del FTQ (B=Bajo, A=Alto, N=Normal)',
  `efi_efecto` char(1) DEFAULT 'N' COMMENT 'Efecto del EFI (B=Bajo, A=Alto, N=Normal)',
  `dis_efecto` char(1) DEFAULT 'N' COMMENT 'Efecto del DIS (B=Bajo, A=Alto, N=Normal)',
  `oee_efecto` char(1) DEFAULT 'N' COMMENT 'Efecto del OEE (B=Bajo, A=Alto, N=Normal)',
  `estatus` char(1) DEFAULT 'F' COMMENT 'Estatus del equipo (F=Funcionando, D=Paro)',
  `fecha_desde` datetime DEFAULT NULL COMMENT 'Fecha de inicio',
  `rate_min` decimal(25,10) DEFAULT '0.0000000000',
  `rate_max` decimal(25,10) DEFAULT '0.0000000000',
  `ratemed` varchar(15) DEFAULT NULL COMMENT 'Unidad de medida',
  `efi` decimal(10,3) DEFAULT '0.000',
  `dis` decimal(10,3) DEFAULT '0.000',
  `ftq` decimal(10,3) DEFAULT '0.000',
  `oee` decimal(10,3) DEFAULT '0.000',
  `esperadoefi` decimal(10,3) DEFAULT '0.000',
  `esperadodis` decimal(10,3) DEFAULT '0.000',
  `esperadooee` decimal(10,3) DEFAULT '0.000',
  `esperadoftq` decimal(10,3) DEFAULT '0.000',
  `rate_mal_desde` datetime DEFAULT NULL COMMENT 'Fecha en que inició el bajo/alto rate',
  `rate_tendencia_baja` datetime DEFAULT NULL COMMENT 'Fecha de tendencia a la baja',
  `rate_tendencia_alta` datetime DEFAULT NULL COMMENT 'Fecha de tendencia a la alta',
  `ftq_tendencia_baja` datetime DEFAULT NULL,
  `oee_tendencia_baja` datetime DEFAULT NULL,
  `dis_tendencia_baja` datetime DEFAULT NULL,
  `efi_tendencia_baja` datetime DEFAULT NULL,
  `proximo_paro` bigint(20) DEFAULT '0' COMMENT 'ID del próximo paro planeado',
  `ultima_lectura` datetime DEFAULT NULL COMMENT 'Fecha de la última lectura',
  `parada_desde` datetime DEFAULT NULL COMMENT 'En paro desde',
  `paro_actual` bigint(20) DEFAULT '0' COMMENT 'ID del paro actual',
  `ultima_produccion` datetime DEFAULT NULL COMMENT 'Fecha última producción (o buffer)',
  `ultima_buffer` datetime DEFAULT NULL,
  `oee_imagen` int(11) DEFAULT '0' COMMENT '0=Normal, 1=Subiendo, 2=Bajando',
  `hoyos` char(1) DEFAULT 'N' COMMENT 'Incluye hoyos en la disponibilidad',
  `iniciar_1` char(1) DEFAULT 'N' COMMENT 'Iniciar el contador',
  `iniciar_2` char(1) DEFAULT 'N',
  `iniciar_3` char(1) DEFAULT 'N',
  `iniciar_4` char(1) DEFAULT 'N',
  `iniciar_5` char(1) DEFAULT 'N',
  `iniciar_6` char(1) DEFAULT 'N',
  `iniciar_7` char(1) DEFAULT 'N',
  `iniciar_8` char(1) DEFAULT 'N',
  `iniciar` char(1) DEFAULT 'N',
  `factor_iniciar` varchar(6) DEFAULT NULL,
  `detener` bigint(20) DEFAULT '0' COMMENT 'Usuario que solicitó detener',
  `reanudar` bigint(20) DEFAULT '0' COMMENT 'Usuario que solicitó la cancelación',
  `transcurrido` int(6) DEFAULT '0' COMMENT 'Ultimo tiempo transcurrido',
  `tiempo_corte` int(2) DEFAULT '-1' COMMENT 'Tiempo de corte',
  `tiempo_reinicio` int(2) DEFAULT '-1' COMMENT 'Tiempo de reinicio',
  `fuera_programa` char(1) DEFAULT 'N' COMMENT 'Esta fuera de programa',
  `alarmado_bajo` char(1) DEFAULT 'N' COMMENT 'Equipo ya alarmado bajo rate',
  `alarmado_alto` char(1) DEFAULT 'N' COMMENT 'Equipo ya alarmado alto rate',
  `alarmado_manual` char(1) DEFAULT 'N' COMMENT 'Equipo ya alarmado por paro manual',
  `transcurrido_pasar` int(1) DEFAULT '0',
  `ultimo_rate` decimal(25,10) DEFAULT '0.0000000000',
  `desde_rate` datetime DEFAULT NULL,
  `efi_imagen` int(1) DEFAULT '0',
  `ftq_imagen` int(1) DEFAULT '0',
  `dis_imagen` int(1) DEFAULT '0',
  `planeado` char(1) DEFAULT 'N' COMMENT 'Si hay un paro planeado no cuenta en el consolidado',
  `ultima_reparacion` datetime DEFAULT NULL COMMENT 'Fecha para controlar el umbral',
  `estado_desde` datetime DEFAULT NULL COMMENT 'Fecha que se muestra en la app',
  `detener_tipo` bigint(20) DEFAULT '0',
  `detener_area` bigint(20) DEFAULT '0',
  `detener_notas` varchar(100) DEFAULT NULL,
  `detener_estimado` bigint(10) DEFAULT '0',
  `detener_piezas` char(1) DEFAULT 'N',
  `detener_paro` varchar(50) DEFAULT NULL,
  `detener_resultados` varchar(100) DEFAULT NULL,
  `alarmado_ftq` char(1) DEFAULT 'N',
  `alarmado_dis` char(1) DEFAULT 'N',
  `alarmado_oee` char(1) DEFAULT 'N',
  `alarmado_efi` char(1) DEFAULT 'N',
  `velocidad` varchar(30) DEFAULT NULL,
  `corte` decimal(20,0) DEFAULT NULL,
  `oee_minutos_rate` decimal(4,0) DEFAULT '0',
  `wip_paro` char(1) DEFAULT 'N',
  `wip_contador` decimal(25,10) DEFAULT '0.0000000000',
  `wip_tiempo` int(6) DEFAULT '0',
  PRIMARY KEY (`equipo`),
  KEY `NewIndex1` (`equipo`,`factor_iniciar`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_maquinas_lecturas` */

/*Table structure for table `relacion_partes_equipos` */

DROP TABLE IF EXISTS `relacion_partes_equipos`;

CREATE TABLE `relacion_partes_equipos` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `parte` bigint(20) DEFAULT '0' COMMENT 'ID del número de parte',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID de la ruta',
  `piezas` decimal(25,10) DEFAULT '0.0000000000' COMMENT 'Piezas por segundo',
  `sesion` char(1) DEFAULT 'N' COMMENT 'Sesión',
  `unidad` varchar(20) DEFAULT NULL COMMENT 'Unidad de medida',
  `tiempo` int(1) DEFAULT '0' COMMENT 'Unidad de tiempo (0=Hora, 1=Minuto, 2=Segundo, 3=Dia)',
  `bajo` decimal(6,2) DEFAULT '0.00' COMMENT 'PCT de bajo rate',
  `alto` decimal(7,2) DEFAULT '0.00' COMMENT 'PCT de sobre rate',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`parte`,`equipo`),
  KEY `NewIndex2` (`parte`,`equipo`,`piezas`)
) ENGINE=MyISAM AUTO_INCREMENT=300 DEFAULT CHARSET=latin1 COMMENT='Relación Equipo/Parte (rate)';

/*Data for the table `relacion_partes_equipos` */

/*Table structure for table `relacion_partes_maquinas` */

DROP TABLE IF EXISTS `relacion_partes_maquinas`;

CREATE TABLE `relacion_partes_maquinas` (
  `parte` bigint(20) DEFAULT NULL COMMENT 'ID del número de parte',
  `maquina` bigint(20) DEFAULT NULL COMMENT 'ID de la ruta',
  KEY `NewIndex1` (`parte`,`maquina`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_partes_maquinas` */

/*Table structure for table `relacion_partes_rutas` */

DROP TABLE IF EXISTS `relacion_partes_rutas`;

CREATE TABLE `relacion_partes_rutas` (
  `parte` bigint(20) DEFAULT NULL COMMENT 'ID del número de parte',
  `ruta` bigint(20) DEFAULT NULL COMMENT 'ID de la ruta',
  `defecto` char(1) DEFAULT NULL COMMENT 'Ruta por defecto',
  KEY `NewIndex1` (`parte`,`ruta`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Relación entre números de parte y rutas';

/*Data for the table `relacion_partes_rutas` */

/*Table structure for table `relacion_plan_checklists` */

DROP TABLE IF EXISTS `relacion_plan_checklists`;

CREATE TABLE `relacion_plan_checklists` (
  `plan` bigint(20) DEFAULT '0' COMMENT 'ID del plan',
  `checklist` bigint(20) DEFAULT '0' COMMENT 'ID del checklist',
  KEY `Index01` (`plan`,`checklist`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_plan_checklists` */

/*Table structure for table `relacion_procesos_sensores` */

DROP TABLE IF EXISTS `relacion_procesos_sensores`;

CREATE TABLE `relacion_procesos_sensores` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `equipo` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  `sensor` int(10) DEFAULT NULL COMMENT 'ID del sensor',
  `tipo` int(1) DEFAULT NULL COMMENT 'Tipo de sensor (0=Producción, 1=Calidad, 2=Otros, 3=Calidad+producción)',
  `ultima_lectura` datetime DEFAULT NULL COMMENT 'Fecha de la última lectura',
  `multiplicador` bigint(20) DEFAULT NULL COMMENT 'Factor de multiplicación',
  `base` bigint(20) DEFAULT NULL COMMENT 'Base del factor ',
  `area` bigint(20) DEFAULT '0' COMMENT 'ID del área por defecto',
  `clasificacion` bigint(20) DEFAULT '0' COMMENT 'ID de la clasificación del rechazo',
  `estatus` char(1) DEFAULT 'A' COMMENT 'Estatus del registro',
  `creacion` datetime DEFAULT NULL COMMENT 'Fecha de creación',
  `modificacion` datetime DEFAULT NULL COMMENT 'Fecha de modificación',
  `creado` bigint(20) DEFAULT '0' COMMENT 'Creado por',
  `modificado` bigint(20) DEFAULT '0' COMMENT 'Modificado por',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`equipo`,`sensor`)
) ENGINE=MyISAM AUTO_INCREMENT=22 DEFAULT CHARSET=latin1;

/*Data for the table `relacion_procesos_sensores` */

/*Table structure for table `relacion_usuarios_opciones` */

DROP TABLE IF EXISTS `relacion_usuarios_opciones`;

CREATE TABLE `relacion_usuarios_opciones` (
  `usuario` bigint(20) DEFAULT NULL COMMENT 'ID del número de parte',
  `opcion` bigint(6) DEFAULT NULL COMMENT 'ID de la ruta',
  KEY `NewIndex1` (`usuario`,`opcion`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Relación usuarios y opciones del sistema';

/*Data for the table `relacion_usuarios_opciones` */

/*Table structure for table `relacion_usuarios_operaciones` */

DROP TABLE IF EXISTS `relacion_usuarios_operaciones`;

CREATE TABLE `relacion_usuarios_operaciones` (
  `usuario` bigint(20) DEFAULT NULL COMMENT 'ID del número de parte',
  `proceso` bigint(20) DEFAULT NULL COMMENT 'ID de la ruta',
  `tipo` int(2) DEFAULT '0' COMMENT 'Catálogo asociado',
  KEY `NewIndex1` (`usuario`,`proceso`,`tipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Relación entre usuarios y operaciones';

/*Data for the table `relacion_usuarios_operaciones` */

/*Table structure for table `relacion_variables_equipos` */

DROP TABLE IF EXISTS `relacion_variables_equipos`;

CREATE TABLE `relacion_variables_equipos` (
  `variable` bigint(20) DEFAULT '0' COMMENT 'ID de la variable',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID del equipo',
  KEY `Index01` (`variable`,`maquina`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `relacion_variables_equipos` */

/*Table structure for table `relaciondefectos` */

DROP TABLE IF EXISTS `relaciondefectos`;

CREATE TABLE `relaciondefectos` (
  `defecto` bigint(11) NOT NULL COMMENT 'ID del defecto',
  `grupo` bigint(11) NOT NULL COMMENT 'ID del grupo',
  `equipo` bigint(11) NOT NULL COMMENT 'ID del equipo',
  PRIMARY KEY (`defecto`,`grupo`,`equipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Relación defectos verss equipos';

/*Data for the table `relaciondefectos` */

/*Table structure for table `relacionparos` */

DROP TABLE IF EXISTS `relacionparos`;

CREATE TABLE `relacionparos` (
  `paro` bigint(11) NOT NULL COMMENT 'ID del paro',
  `grupo` bigint(11) NOT NULL COMMENT 'ID del grupo',
  `equipo` bigint(11) NOT NULL COMMENT 'ID del equipo',
  PRIMARY KEY (`paro`,`grupo`,`equipo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Relacion paros versus equips';

/*Data for the table `relacionparos` */

/*Table structure for table `reportes` */

DROP TABLE IF EXISTS `reportes`;

CREATE TABLE `reportes` (
  `id` bigint(20) NOT NULL AUTO_INCREMENT COMMENT 'ID del reporte',
  `numero` bigint(20) DEFAULT '0' COMMENT 'Número de reporte',
  `fecha` datetime NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha y hora del reporte',
  `linea` bigint(20) DEFAULT '0' COMMENT 'ID de la célula',
  `maquina` bigint(20) DEFAULT '0' COMMENT 'ID de la máquina',
  `turno` bigint(20) DEFAULT '0' COMMENT 'ID del turno',
  `area` bigint(20) DEFAULT '0' COMMENT 'ID del área',
  `falla` bigint(20) DEFAULT '0' COMMENT 'ID de la falla',
  `detalle` varchar(300) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Detalle de la falla',
  `record_id` bigint(11) DEFAULT '0' COMMENT 'Registro en MMCall',
  `requester` varchar(20) COLLATE utf8_bin DEFAULT NULL COMMENT 'Código del requester en MMCall',
  `solicitante` bigint(20) DEFAULT '0' COMMENT 'Usuario solicitante',
  `tecnicoatend` bigint(20) DEFAULT '0' COMMENT 'Técnico que atendió',
  `tecnico` bigint(20) DEFAULT '0' COMMENT 'Técnico que reparó',
  `tipo` bigint(20) DEFAULT '0' COMMENT 'Tipo de mantenimiento',
  `estatus` int(4) DEFAULT '0' COMMENT 'Estatus del reporte',
  `falla_ajustada` bigint(20) DEFAULT '0' COMMENT 'ID de la falla ajustada',
  `inicio_atencion` datetime DEFAULT NULL COMMENT 'Fecha y hora de inicio de la atención',
  `cierre_atencion` datetime DEFAULT NULL COMMENT 'Fecha y hora del cierre de la atención',
  `inicio_reporte` datetime DEFAULT NULL COMMENT 'Fecha y hora de inicio del reporte',
  `alarmado` char(1) CHARACTER SET latin1 DEFAULT 'N' COMMENT 'El reporte fue alarmado?',
  `tiemporeporte` bigint(12) DEFAULT '0' COMMENT 'Tiempo en segundos del reporte',
  `comentarios` varchar(300) CHARACTER SET latin1 DEFAULT NULL COMMENT 'Comentarios del reporte',
  `cierre_reporte` datetime DEFAULT NULL COMMENT 'Fecha y hora del cierre del reporte',
  `contabilizar` char(1) CHARACTER SET latin1 DEFAULT 'S' COMMENT 'Contabilizar el reporte?',
  `confirmado` bigint(20) DEFAULT '0' COMMENT 'Usuario que confirmó la reparación',
  `fecha_reporte` date DEFAULT NULL COMMENT 'Fecha del reporte',
  `tiemporeparacion` bigint(12) DEFAULT '0' COMMENT 'Tiempo total de la falla',
  `tiempoatencion` bigint(12) DEFAULT '0' COMMENT 'Tiempo total de la atención',
  `tiempollegada` bigint(12) DEFAULT '0' COMMENT 'Tiempo de la llegada del tecnico',
  `modificador` bigint(20) DEFAULT '0' COMMENT 'Usuario que modificó',
  `modificado` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT 'Fecha y hora del úñtimo cambio',
  `alarmado_atender` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Alarmado en estatus 0',
  `alarmado_atendido` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Alarmado en estatus 10',
  `alarmado_atender_desde` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó',
  `alarmado_atendido_desde` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó',
  `alarmado_informe_desde` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó',
  `generar_audio` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Generar audio por área',
  `repeticiones` int(3) DEFAULT '0' COMMENT 'Número de repeticion',
  `escalado` int(1) DEFAULT '0' COMMENT 'Nivel de escalamiento',
  `mmcall` datetime DEFAULT NULL COMMENT 'Mensaje enviado a MMCall',
  `afecta_oee` char(1) COLLATE utf8_bin DEFAULT 'S',
  `herramental` bigint(20) DEFAULT '0',
  `primer_audio` char(1) COLLATE utf8_bin DEFAULT 'N',
  `audios` datetime DEFAULT NULL COMMENT 'N',
  `reproducir_audio_externo` int(2) DEFAULT '0',
  `origen` bigint(12) DEFAULT '0' COMMENT 'Origen del reporte (0=Digital, 1=Híbrido, 2=Manual)',
  `alarmado_atencion` char(1) COLLATE utf8_bin DEFAULT 'N' COMMENT 'Alarmado alarmado 0 o 10',
  `alarmado_atencion_desde` datetime DEFAULT NULL COMMENT 'Fecha en que se alarmó atención''',
  `contabilizar_fecha` datetime DEFAULT NULL,
  `contabilizar_usuario` bigint(20) DEFAULT '0',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`numero`),
  KEY `NewIndex2` (`fecha`),
  KEY `NewIndex3` (`maquina`,`estatus`),
  KEY `NewIndex4` (`linea`,`estatus`),
  KEY `NewIndex5` (`area`,`estatus`),
  KEY `NewIndex6` (`origen`)
) ENGINE=MyISAM AUTO_INCREMENT=22850 DEFAULT CHARSET=utf8 COLLATE=utf8_bin COMMENT='Tabla de reportes de Mantenimiento';

/*Data for the table `reportes` */

/*Table structure for table `ruta_congelada` */

DROP TABLE IF EXISTS `ruta_congelada`;

CREATE TABLE `ruta_congelada` (
  `lote` bigint(20) DEFAULT '0' COMMENT 'ID del lote',
  `id_detruta` bigint(20) DEFAULT '0' COMMENT 'ID de la ruta',
  `ruta` bigint(20) DEFAULT '0' COMMENT 'ID de la rutaa',
  `secuencia` int(6) DEFAULT '0' COMMENT 'Secuencia',
  `proceso` bigint(20) DEFAULT '0' COMMENT 'ID del proceso',
  KEY `NewIndex1` (`lote`,`secuencia`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Ruta congelada';

/*Data for the table `ruta_congelada` */

/*Table structure for table `sentencias` */

DROP TABLE IF EXISTS `sentencias`;

CREATE TABLE `sentencias` (
  `id` bigint(10) NOT NULL COMMENT 'ID de la consulta',
  `version` varchar(30) DEFAULT NULL COMMENT 'ID de la licencia',
  `orden` int(4) DEFAULT NULL COMMENT 'Orden de la ejecución',
  `sentencia` varchar(5000) DEFAULT NULL COMMENT 'Sentencia a aplicar',
  `estatus` char(1) DEFAULT '0' COMMENT 'Estatus de la licencia (0=Por aplicar, 1=Aplicado, 9=Error)',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Log de versiones';

/*Data for the table `sentencias` */

/*Table structure for table `status_objetos` */

DROP TABLE IF EXISTS `status_objetos`;

CREATE TABLE `status_objetos` (
  `id` int(11) NOT NULL,
  `descripcion` varchar(200) NOT NULL,
  `color` varchar(10) NOT NULL,
  `normal` tinyint(4) NOT NULL DEFAULT '0',
  `mapa_Id` int(11) NOT NULL,
  PRIMARY KEY (`id`,`mapa_Id`),
  KEY `fk_status_objetos_mapas_idx` (`mapa_Id`)
) ENGINE=MyISAM DEFAULT CHARSET=utf8;

/*Data for the table `status_objetos` */

insert  into `status_objetos`(`id`,`descripcion`,`color`,`normal`,`mapa_Id`) values (0,'critico','red',0,83),(0,'critico','red',0,84),(1,'critico','red',0,83),(1,'critico','red',0,84),(2,'atencion','yellow',0,83),(2,'atencion','yellow',0,84),(3,'normal','',1,83),(3,'atencion','yellow',1,84),(10,'atencion','yellow',0,83),(0,'critico','red',0,85),(10,'atencion','yellow',0,85),(0,'atencion','yellow',0,86),(0,'critico','red',0,87),(10,'atencion','yellow',0,87),(0,'critico','red',0,89),(10,'atencion','yellow',0,89),(0,'critico','red',0,90),(10,'atencion','yellow',0,90),(0,'critico','red',0,91),(10,'atencion','yellow',0,91),(0,'critico','red',0,92),(10,'atencion','yellow',0,92),(0,'critico','red',0,93),(10,'atencion','yellow',0,93),(0,'critico','red',0,94),(10,'atencion','yellow',0,94),(0,'critico','red',0,95),(10,'atencion','yellow',0,95),(0,'critico','red',0,96),(10,'atencion','yellow',0,96),(0,'critico','red',0,97),(10,'atencion','yellow',0,97),(0,'critico','red',0,98),(10,'atencion','yellow',0,98),(0,'critico','red',0,99),(10,'atencion','yellow',0,99),(0,'critico','red',0,100),(10,'atencion','yellow',0,100),(0,'critico','red',0,103),(10,'atencion','yellow',0,103),(0,'critico','red',0,104),(10,'atencion','yellow',0,104),(0,'critico','red',0,105),(10,'atencion','yellow',0,105),(0,'critico','red',0,106),(10,'atencion','yellow',0,106),(0,'critico','red',0,107),(10,'atencion','yellow',0,107),(0,'critico','red',0,108),(10,'atencion','yellow',0,108),(0,'critico','red',0,109),(10,'atencion','yellow',0,109),(0,'critico','red',0,110),(10,'atencion','yellow',0,110),(0,'critico','red',0,111),(10,'atencion','yellow',0,111),(0,'critico','red',0,113),(10,'atencion','yellow',0,113),(0,'critico','red',0,114),(10,'atencion','yellow',0,114),(0,'critico','red',0,115),(10,'atencion','yellow',0,115),(0,'critico','red',0,116),(10,'atencion','yellow',0,116),(0,'critico','red',0,117),(10,'atencion','yellow',0,117),(0,'critico','red',0,118),(10,'atencion','yellow',0,118),(0,'critico','red',0,119),(10,'atencion','yellow',0,119),(0,'critico','red',0,120),(10,'atencion','yellow',0,120),(0,'critico','red',0,121),(10,'atencion','yellow',0,121),(0,'critico','red',0,122),(10,'atencion','yellow',0,122),(0,'critico','red',0,123),(10,'atencion','yellow',0,123),(0,'critico','red',0,124),(10,'atencion','yellow',0,124);

/*Table structure for table `tablas` */

DROP TABLE IF EXISTS `tablas`;

CREATE TABLE `tablas` (
  `id` int(6) NOT NULL COMMENT 'ID de la tabla',
  `nombre` varchar(50) DEFAULT NULL COMMENT 'Nombre de la tabla',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COMMENT='Catálogo de tablas';

/*Data for the table `tablas` */

insert  into `tablas`(`id`,`nombre`) values (50,'TIPOS DE MÁQUINA'),(10,'LINEAS AGRUPADOR 1'),(15,'LINEAS AGRUPADOR 2'),(20,'MAQUINAS AGRUPADOR 1'),(25,'MAQUINAS AGRUPADOR 2'),(30,'AREAS AGRUPADOR 1'),(35,'AREAS AGRUPADOR 2'),(40,'FALLAS AGRUPADOR 1'),(45,'FALLAS AGRUPADOR 2'),(70,'DEPARTAMENTOS'),(80,'COMPANIAS'),(90,'PLANTAS'),(100,'PAGERS DE MMCALL'),(60,'TIPOS DE MANTENIMIENTO'),(105,'RECHAZOS DE OAE'),(110,'UNIDADES DE MEDIDA'),(115,'TIPOS DE VARIABLES'),(120,'TIPOS DE CHECKLISTS');

/*Table structure for table `temporal_mmcall` */

DROP TABLE IF EXISTS `temporal_mmcall`;

CREATE TABLE `temporal_mmcall` (
  `record` bigint(20) NOT NULL DEFAULT '0' COMMENT 'ID de la tabla records',
  `boton1` datetime DEFAULT NULL COMMENT 'Fecha botón 1',
  `boton2` datetime DEFAULT NULL COMMENT 'Fecha botón 2',
  `estatus` int(1) NOT NULL DEFAULT '0' COMMENT 'estatus del registro',
  PRIMARY KEY (`record`),
  UNIQUE KEY `Index01` (`record`,`estatus`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `temporal_mmcall` */

/*Table structure for table `tipos_figuras` */

DROP TABLE IF EXISTS `tipos_figuras`;

CREATE TABLE `tipos_figuras` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `descripcion` varchar(100) NOT NULL,
  PRIMARY KEY (`id`)
) ENGINE=MyISAM AUTO_INCREMENT=128 DEFAULT CHARSET=utf8;

/*Data for the table `tipos_figuras` */

insert  into `tipos_figuras`(`id`,`descripcion`) values (107,'Diamond'),(108,'Oval'),(109,'DownArrow'),(110,'Parallelogram'),(111,'RegularPentagon'),(112,'Hexagon'),(113,'Can'),(114,'Rectangle'),(115,'IsoscelesTriangle'),(116,'5pointStar'),(117,'RightArrow'),(118,'RoundedRectangle'),(119,'Cross'),(120,'6pointStar'),(121,'UpArrow'),(122,'LeftArrow'),(123,'BentUpArrow'),(124,'Cube'),(125,'Picture'),(126,'FlowchartData'),(127,'Trapezoid');

/*Table structure for table `traduccion` */

DROP TABLE IF EXISTS `traduccion`;

CREATE TABLE `traduccion` (
  `id` bigint(11) NOT NULL AUTO_INCREMENT COMMENT 'ID del registro',
  `literal` varchar(100) DEFAULT NULL COMMENT 'Literal a buscar',
  `idioma` varchar(5) DEFAULT NULL COMMENT 'Idioma',
  `traduccion` varchar(100) DEFAULT NULL COMMENT 'Traducción',
  PRIMARY KEY (`id`),
  KEY `NewIndex1` (`literal`,`idioma`)
) ENGINE=MyISAM AUTO_INCREMENT=12 DEFAULT CHARSET=latin1 COMMENT='Tabla de traducción';

/*Data for the table `traduccion` */

/*Table structure for table `traduccion_literales` */

DROP TABLE IF EXISTS `traduccion_literales`;

CREATE TABLE `traduccion_literales` (
  `planta` bigint(20) DEFAULT '0',
  `idioma` int(4) DEFAULT '0',
  `literal` bigint(20) DEFAULT NULL,
  `traduccion` varchar(300) DEFAULT NULL,
  UNIQUE KEY `Index01` (`planta`,`idioma`,`literal`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `traduccion_literales` */

/*Table structure for table `traduccion_when` */

DROP TABLE IF EXISTS `traduccion_when`;

CREATE TABLE `traduccion_when` (
  `campo` int(4) DEFAULT '0',
  `id` varchar(10) DEFAULT NULL COMMENT 'Literal a buscar',
  `idioma` bigint(4) DEFAULT '0' COMMENT 'Idioma',
  `nombre` varchar(50) DEFAULT NULL COMMENT 'Traducción',
  `orden` int(4) DEFAULT '0' COMMENT 'Orden del registro',
  KEY `NewIndex1` (`campo`,`id`,`idioma`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

/*Data for the table `traduccion_when` */

insert  into `traduccion_when`(`campo`,`id`,`idioma`,`nombre`,`orden`) values (10,'T',1,'Todos los días',1),(10,'LV',1,'De lunes a viernes',2),(10,'L',1,'Cada lunes',3),(10,'M',1,'Cada martes',4),(10,'MI',1,'Cada miércoles',5),(10,'J',1,'Cada jueves',6),(10,'V',1,'Cada viernes',7),(10,'S',1,'Cada sábado',8),(10,'D',1,'Cada domingo',9),(10,'1M',1,'El primer día del mes',10),(10,'UM',1,'El último día del mes',11),(20,'T',1,'A toda hora',1),(20,'0',1,'A media noche (12:00 am)',2),(20,'1',1,'1:00 am',3),(20,'2',1,'2:00 am',4),(20,'3',1,'3:00 am',5),(20,'4',1,'4:00 am',6),(20,'5',1,'5:00 am',7),(20,'6',1,'6:00 am',8),(20,'7',1,'7:00 am',9),(20,'8',1,'8:00 am',10),(20,'9',1,'9:00 am',11),(20,'10',1,'10:00 am',12),(20,'11',1,'11:00 am',13),(20,'12',1,'Medio día (12:00 m)',14),(20,'13',1,'1:00 pm',15),(20,'14',1,'2:00 pm',16),(20,'15',1,'3:00 pm',17),(20,'16',1,'4:00 pm',18),(20,'17',1,'5:00 pm',19),(20,'18',1,'6:00 pm',20),(20,'19',1,'7:00 pm',21),(20,'20',1,'8:00 pm',22),(20,'21',1,'9:00 pm',23),(20,'22',1,'10:00 pm',24),(20,'23',1,'11:00 pm',25),(30,'0',1,'Segundos (atrás)',1),(30,'1',1,'Minutos (atrás)',2),(30,'2',1,'Horas (atrás)',3),(30,'3',1,'Días (atrás)',4),(30,'4',1,'Semanas (atrás)',5),(30,'5',1,'Meses (atrás)',6),(30,'6',1,'Años (atrás)',7),(30,'10',1,'Lo que va del día',8),(30,'11',1,'Lo que va de la semana',9),(30,'12',1,'Lo que va del mes (MtD)',10),(30,'13',1,'Lo que va del año (YtD)',11),(30,'20',1,'Ayer',12),(30,'21',1,'Semana anterior',13),(30,'22',1,'Mes anterior',14),(30,'23',1,'Año anterior',15),(100,'A',1,'Activo',1),(100,'I',1,'Inactivo',2),(110,'N',1,'No',2),(110,'S',1,'Sí',1),(120,'S',1,'Sí',1),(120,'N',1,'No',2),(120,'Y',1,'Ya usado!',3),(130,'0',1,'Permitir el uso de la misma contraseña',1),(130,'1',1,'Después de 1 uso',2),(130,'2',1,'Después de 2 usos',3),(130,'3',1,'Después de 3 usos',4),(130,'4',1,'Después de 4 usos',5),(130,'5',1,'Después de 5 usos',6),(140,'0',1,'PRODUCCION (pieza buena)',1),(140,'1',1,'CALIDAD (rechazo antes de conteo)',2),(140,'2',1,'CALIDAD (rechazo después de conteo)',4),(140,'3',1,'BUFFER (pieza en proceso)',5),(140,'4',1,'PRODUCCION (RUN/OFF)',3),(150,'0',1,'Una fecha específica',0);

/*Table structure for table `variables_valores` */

DROP TABLE IF EXISTS `variables_valores`;

CREATE TABLE `variables_valores` (
  `variable` bigint(20) DEFAULT '0' COMMENT 'ID de variable',
  `orden` int(3) DEFAULT '0' COMMENT 'Indice del valor',
  `valor` varchar(50) DEFAULT NULL COMMENT 'Valor',
  `alarmar` char(1) DEFAULT 'N' COMMENT 'Alarmar',
  `defecto` char(1) DEFAULT 'N',
  KEY `NewIndex` (`variable`,`orden`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;
")
                XtraMessageBox.Show(traduccion(110), "SIGMA v1.0", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If


        cadSQL = "SELECT * FROM " & rutaBD & ".sentencias WHERE estatus = 'N'"
        Dim sentencias As DataSet = consultaSEL(cadSQL)

        If sentencias.Tables(0).Rows.Count > 0 Then
            huboProceso = True
            Dim tSentencias = 0, aSentencias = 0
            For Each sentencia In sentencias.Tables(0).Rows
                tSentencias = tSentencias + 1
                regsAfectados = consultaACT(sentencia!sentencia)
                If regsAfectados <> -1 Then
                    aSentencias = aSentencias + 1
                    regsAfectados = consultaACT("UPDATE " & rutaBD & ".sentencias SET estatus = 'S' WHERE id = " & sentencia!id)
                End If
            Next
            regsAfectados = consultaACT("DELETE FROM " & rutaBD & ".sentencias WHERE estatus = 'S'")
            XtraMessageBox.Show(traduccion(111) & tSentencias & traduccion(112) & aSentencias & traduccion(113), "SIGMA v1.0", MessageBoxButtons.OK, If(aSentencias = tSentencias, MessageBoxIcon.Information, MessageBoxIcon.Error))

        End If
        If huboProceso Then
            XtraMessageBox.Show(traduccion(114), "SIGMA v1.0", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Application.Restart()

        Else
            iniciarPantalla()
        End If

    End Sub

    Sub etiquetas()
        Dim general = consultaSEL("SELECT cadena FROM " & rutaBD & ".det_idiomas_back WHERE idioma = " & IIf(be_idioma = 0, 1, be_idioma) & " AND modulo = 0 ORDER BY linea")
        Dim cadenaTrad = ""
        If general.Tables(0).Rows.Count > 0 Then
            For Each cadena In general.Tables(0).Rows
                cadenaTrad = cadenaTrad & cadena!cadena
            Next
        End If
        traduccion = cadenaTrad.Split(New Char() {";"c})
        Me.Text = traduccion(127)
        'LabelControl2.Text = traduccion(128)
        LabelControl1.Text = traduccion(135)
        HyperlinkLabelControl1.Text = traduccion(129)
        LabelControl4.Text = traduccion(133)
        GroupControl1.Text = traduccion(134)

        SimpleButton1.Text = traduccion(130)
        SimpleButton2.ToolTip = traduccion(131)
        SimpleButton3.ToolTip = traduccion(132)
        SimpleButton3.ToolTipTitle = traduccion(137)
        NotifyIcon1.Text = traduccion(136)
        BarStaticItem1.Caption = traduccion(4) & Format(Now(), "dddd, dd-MMM-yyyy HH:mm:ss")
        If estadoPrograma Then
            agregarSolo(traduccion(5))
        End If

    End Sub

    Private Sub fallasPLC_Tick(sender As Object, e As EventArgs)
        If fallasPLCProcesando Or Not estadoPrograma Then Exit Sub
        fallasPLCProcesando = True


        Dim cadSQL = "SELECT * FROM " & rutaBD & ".fallas_plc WHERE estatus = 0"
        Dim cadAgregar = ""
        Dim registros As DataSet = consultaSEL(cadSQL, cadenaConexionFALLAS)
        Dim pendiente As DataSet
        Dim general As DataSet
        Dim idMaquina As Long = 0
        Dim idLinea As Long = 0
        Dim idArea As Long = 0
        Dim idUSuario As Long = 0
        Dim regsAfectados = 0

        Dim ultimoHora = Format(Now, "HH:mm:ss")
        Dim restar = 0
        Dim reportesCreados = 0

        If registros Is Nothing Then
            agregarSolo("No hubo conexión")
            Exit Sub
        End If
        If registros.Tables(0).Rows.Count > 0 Then
            For Each registro In registros.Tables(0).Rows
                cadSQL = "SELECT record, estatus FROM " & rutaBD & ".temporal_plc WHERE record = " & registro!id
                pendiente = consultaSEL(cadSQL)
                If pendiente.Tables(0).Rows.Count = 0 Then

                    Dim fechaReporte = registro!fecha
                    cadSQL = "SELECT * FROM " & rutaBD & ".cat_turnos WHERE id = " & be_turno_actual
                    Dim horarios As DataSet = consultaSEL(cadSQL)
                    If horarios.Tables(0).Rows.Count > 0 Then
                        If ValNull(horarios.Tables(0).Rows(0)!cambiodia, "A") = "S" Then

                            If ValNull(horarios.Tables(0).Rows(0)!mover, "N") = 1 And Format(fechaReporte, "HH") >= "00" And Format(fechaReporte, "HH:mm:ss") < horarios.Tables(0).Rows(0)!termina.ToString Then
                                fechaReporte = DateAdd(DateInterval.Day, -1, fechaReporte)
                            ElseIf ValNull(horarios.Tables(0).Rows(0)!mover, "N") = 2 And Format(fechaReporte, "HH:mm:ss") >= horarios.Tables(0).Rows(0)!inicia.ToString And Format(fechaReporte, "HHmmss") <= "235959" Then
                                fechaReporte = DateAdd(DateInterval.Day, 1, fechaReporte)
                            End If
                        End If
                        idUSuario = ValNull(horarios.Tables(0).Rows(0)!usuario, "N")
                    End If
                    idMaquina = 0
                    idLinea = 0
                    idArea = 0
                    idUSuario = 0
                    Dim tipo_andon = 0
                    'Se busca máquina/línea

                    cadSQL = "SELECT a.id, a.nombre, IFNULL((SELECT proceso FROM " & rutaBD & ".relacion_fallas_operaciones WHERE falla = a.id AND tipo = 1), 0) AS linea, IFNULL((SELECT proceso FROM  " & rutaBD & ".relacion_fallas_operaciones WHERE falla = a.id AND tipo = 2), 0) AS maquina, IFNULL((SELECT proceso FROM  " & rutaBD & ".relacion_fallas_operaciones WHERE falla = a.id AND tipo = 3), 0) AS area FROM  " & rutaBD & ".cat_fallas a WHERE a.plc = '" & registro!falla & "' AND estatus = 'A'"
                    general = consultaSEL(cadSQL)
                    If general.Tables(0).Rows.Count > 0 Then
                        idMaquina = general.Tables(0).Rows(0)!maquina
                        idLinea = general.Tables(0).Rows(0)!linea
                        idArea = general.Tables(0).Rows(0)!area
                        'Se agrega el registro en la tabla de reportes
                        If idMaquina > 0 And idLinea > 0 And idArea > 0 Then

                            cadSQL = "SELECT linea, usuario FROM " & rutaBD & ".cat_maquinas WHERE estatus = 'A' AND id = " & idMaquina
                            Dim general2 As DataSet = consultaSEL(cadSQL)
                            If general2.Tables(0).Rows.Count > 0 Then
                                idLinea = general2.Tables(0).Rows(0)!linea
                                If ValNull(general2.Tables(0).Rows(0)!usuario, "N") > 0 Then
                                    idUSuario = ValNull(general2.Tables(0).Rows(0)!usuario, "N")
                                End If
                            End If

                            reportesCreados = reportesCreados + 1
                            cadAgregar = cadAgregar & "INSERT INTO " & rutaBD & ".reportes (linea, maquina, area, falla, falla_ajustada, solicitante, turno, fecha_reporte, fecha, mmcall, numero) VALUES(" & idLinea & ", " & idMaquina & ", " & idArea & ", " & general.Tables(0).Rows(0)!id & ", " & general.Tables(0).Rows(0)!id & ", " & IIf(idUSuario = 0, 1, idUSuario) & ", " & be_turno_actual & ", '" & Format(fechaReporte, "yyyy/MM/dd") & "', '" & Format(fechaReporte, "yyyy/MM/dd HH:mm:ss") & "', NOW(), " & registro!id & ")"
                            cadAgregar = cadAgregar & ";INSERT INTO " & rutaBD & ".temporal_plc (record) VALUES(" & registro!id & ");"
                        End If
                    End If
                End If
            Next
            If cadAgregar.Length > 0 Then
                cadAgregar = Microsoft.VisualBasic.Strings.Mid(cadAgregar, 1, Len(cadAgregar) - 1)
                regsAfectados = consultaACT(cadAgregar)
                If reportesCreados > 0 Then
                    agregarLOG(traduccion(185).Replace("campo_0", reportesCreados), 0, 0)
                End If
            End If
        End If



        'Se revisa de reversa (cierre de tickets)
        cadSQL = "SELECT * FROM " & rutaBD & ".temporal_plc WHERE estatus = 0"
        registros = consultaSEL(cadSQL)
        If registros.Tables(0).Rows.Count > 0 Then
            reportesCreados = 0
            cadAgregar = ""
            For Each registro In registros.Tables(0).Rows
                cadSQL = "SELECT fecha_cierre FROM " & rutaBD & ".fallas_plc WHERE id = " & registro!record & " AND NOT ISNULL(fecha_cierre) AND estatus = 1"
                pendiente = consultaSEL(cadSQL, cadenaConexionFALLAS)
                cadSQL = "SELECT a.area, b.tecnico, b.cerrar_boton FROM " & rutaBD & ".reportes a INNER JOIN " & rutaBD & ".cat_areas b ON a.area = b.id WHERE a.numero = " & registro!record
                general = consultaSEL(cadSQL)
                Dim accion = "N"
                Dim tecnico = 0
                Dim fANDON = Format(DateAndTime.Now, "yyyy/MM/dd HH:mm:ss")
                If general.Tables(0).Rows.Count > 0 Then
                    accion = ValNull(general.Tables(0).Rows(0)!cerrar_boton, "A")
                    tecnico = ValNull(general.Tables(0).Rows(0)!tecnico, "N")
                End If
                If pendiente.Tables(0).Rows.Count > 0 Then
                    Dim fMMcall = Format(pendiente.Tables(0).Rows(0)!fecha_cierre, "yyyy/MM/dd HH:mm:ss")
                    reportesCreados = reportesCreados + 1
                    'Se valida si se cierra el reporte o no
                    If accion = "N" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 10, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)) WHERE numero = " & registro!record
                    ElseIf accion = "R" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 100, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), cierre_atencion = '" & fMMcall & "', tiemporeparacion = 0, inicio_reporte = NOW(), detalle = '" & traduccion(107) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE numero = " & registro!record
                    ElseIf accion = "C" Then
                        cadAgregar = cadAgregar & "UPDATE " & rutaBD & ".reportes SET estatus = 1000, inicio_atencion = '" & fMMcall & "', tiempollegada = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), cierre_atencion = '" & fMMcall & "', tiemporeparacion = 0, tiemporeporte = 0, inicio_reporte = '" & fMMcall & "', cierre_reporte = '" & fMMcall & "', detalle = '" & traduccion(108) & "', comentarios = '" & traduccion(108) & "', tiempoatencion = TIME_TO_SEC(TIMEDIFF('" & fMMcall & "', fecha)), tecnicoatend = " & tecnico & ", tecnico = " & tecnico & " WHERE numero = " & registro!record
                    End If
                    cadAgregar = cadAgregar & ";UPDATE " & rutaBD & ".temporal_plc SET estatus = 2 WHERE record = " & registro!record & ";"
                End If
            Next
            If cadAgregar.Length > 0 Then
                cadAgregar = Microsoft.VisualBasic.Strings.Mid(cadAgregar, 1, Len(cadAgregar) - 1)
                regsAfectados = consultaACT(cadAgregar)
                If reportesCreados > 0 Then
                    agregarLOG(traduccion(186).Replace("campo_0", reportesCreados), 0, 0)
                End If
            End If
        End If

        cadSQL = "SELECT MAX(id) AS maximo FROM " & rutaBD & ".lecturas"
        general = consultaSEL(cadSQL)
        If general.Tables(0).Rows.Count > 0 Then
            cadSQL = "SELECT * FROM " & rutaBD & ".lecturas WHERE id > " & ValNull(general.Tables(0).Rows(0)!maximo, "N") & " ORDER BY id"
        Else
            cadSQL = "SELECT * FROM " & rutaBD & ".lecturas ORDER BY id DESC LIMIT 1"
        End If
        registros = consultaSEL(cadSQL, cadenaConexionFALLAS)
        cadSQL = ""
        Dim agregadas = 0
        If registros.Tables(0).Rows.Count > 0 Then
            For Each registro In registros.Tables(0).Rows
                agregadas = agregadas + 1
                cadSQL = cadSQL & "INSERT INTO " & rutaBD & ".lecturas (fecha, sensor, valor) VALUES('" & Format(registro!fecha, "yyyy/MM/dd HH:mm:ss") & "', " & registro!sensor & ", " & registro!valor & ");"
            Next
        End If
        If cadSQL.Length > 0 Then
            cadSQL = Strings.Mid(cadSQL, 1, cadSQL.Length - 1)
            regsAfectados = consultaACT(cadSQL)
        End If
        fallasPLCProcesando = False
    End Sub

    Sub calcularTiempos(fechaD, fechaH, proceso, maquina)
        tDisponible = 3600
        tMantto = 0
        Dim diaSemana = 0
        Dim fechaEspecifica As Boolean = False
        Dim procesoEspecifico As Boolean = False
        Dim fechaEstimada = fechaD
        Dim horaDesde = Format(fechaD, "HH:mm:ss")
        Dim horaHasta = Format(fechaH, "HH:mm:ss")
        Dim primerDia = True
        Dim MaximoDias = 14
        Dim diasContados = 0
        Dim primerRegistro = True
        Dim combinacion = 0
        Dim sumando As Boolean = True
        Dim continuar = False
        Dim rangosPositivosD(0) As String
        Dim rangosPositivosH(0) As String
        Dim rangosNegativosD(0) As String
        Dim rangosNegativosH(0) As String
        Dim rangoPositivo = 0
        Dim rangoNegativo = 0
        Dim horarios As DataSet
        Dim cadSQL = ""
        'Recorrido por día
        diaSemana = DateAndTime.Weekday(fechaEstimada) - 1
        cadSQL = "SELECT desde, hasta, dia, proceso, tipo FROM " & rutaBD & ".horarios WHERE (dia = " & diaSemana & " OR (dia = 9 AND fecha = '" & Format(fechaEstimada, "yyyy/MM/dd") & "')) AND (proceso = 0 OR proceso = " & proceso & ") AND ('" & horaHasta & "' <= hasta OR '" & horaDesde & "' >= desde) ORDER BY tipo DESC, proceso DESC, dia DESC, desde, hasta"
        horarios = consultaSEL(cadSQL)
        procesoEspecifico = False
        fechaEspecifica = False

        If horarios.Tables(0).Rows.Count > 0 Then
            For Each rango In horarios.Tables(0).Rows
                If rango!desde.ToString < horaHasta And rango!hasta.ToString >= horaDesde Then
                    If primerRegistro Then

                        primerRegistro = False
                        If rango!tipo = "S" Then
                            rangoPositivo = rangoPositivo + 1
                            If rango!desde.ToString < horaDesde Then
                                rangosPositivosD(rangoPositivo - 1) = horaDesde
                            Else
                                rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                            End If
                            If rango!hasta.ToString > horaHasta Then
                                rangosPositivosH(rangoPositivo - 1) = horaHasta
                            Else
                                rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                            End If
                            procesoEspecifico = rango!proceso <> 0
                            fechaEspecifica = rango!dia = 9
                            If procesoEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf procesoEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        End If
                    Else
                        If sumando And rango!tipo = "N" Then
                            sumando = False
                            rangoNegativo = rangoNegativo + 1
                            If rango!desde.ToString < horaDesde Then
                                rangosNegativosD(rangoNegativo - 1) = horaDesde
                            Else
                                rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                            End If
                            If rango!hasta.ToString > horaHasta Then
                                rangosNegativosH(rangoNegativo - 1) = horaHasta
                            Else
                                rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString()
                            End If
                            procesoEspecifico = rango!proceso <> 0
                            fechaEspecifica = rango!dia = 9
                            If procesoEspecifico And fechaEspecifica Then
                                combinacion = 1
                            ElseIf procesoEspecifico Then
                                combinacion = 2
                            ElseIf fechaEspecifica Then
                                combinacion = 3
                            Else
                                combinacion = 4
                            End If
                        Else
                            If combinacion = 1 Then
                                continuar = rango!proceso <> 0 And rango!dia <> 9
                            ElseIf combinacion = 2 Then
                                continuar = rango!proceso <> 0
                            ElseIf combinacion = 3 Then
                                continuar = rango!dia = 9
                            Else
                                continuar = True
                            End If
                        End If
                        If continuar Then
                            If sumando Then
                                rangoPositivo = rangoPositivo + 1
                                ReDim Preserve rangosPositivosD(rangoPositivo)
                                ReDim Preserve rangosPositivosH(rangoPositivo)
                                If rango!desde.ToString < horaDesde Then
                                    rangosPositivosD(rangoPositivo - 1) = horaDesde
                                Else
                                    rangosPositivosD(rangoPositivo - 1) = rango!desde.ToString
                                End If
                                If rango!hasta.ToString > horaHasta Then
                                    rangosPositivosH(rangoPositivo - 1) = horaHasta
                                Else
                                    rangosPositivosH(rangoPositivo - 1) = rango!hasta.ToString
                                End If
                            Else
                                rangoNegativo = rangoNegativo + 1
                                ReDim Preserve rangosNegativosD(rangoNegativo)
                                ReDim Preserve rangosNegativosH(rangoNegativo)
                                If rango!desde.ToString < horaDesde Then
                                    rangosNegativosD(rangoNegativo - 1) = horaDesde
                                Else
                                    rangosNegativosD(rangoNegativo - 1) = rango!desde.ToString
                                End If
                                If rango!hasta.ToString > horaHasta Then
                                    rangosNegativosH(rangoNegativo - 1) = horaHasta
                                Else
                                    rangosNegativosH(rangoNegativo - 1) = rango!hasta.ToString()
                                End If

                            End If
                        End If
                    End If
                End If
            Next
        End If
        'Se crear un registro único por día
        If rangoPositivo > 0 Then
            Dim arreDefD(0) As String
            Dim arreDefH(0) As String
            Dim arreDefP(0) As String
            Dim totalItems = 1
            arreDefD(0) = rangosPositivosD(0)
            arreDefH(0) = rangosPositivosH(0)
            arreDefP(0) = "S"
            For i = 0 To rangoPositivo - 1
                If rangosPositivosD(i) > arreDefD(totalItems - 1) Then
                    totalItems = totalItems + 1
                    ReDim Preserve arreDefD(totalItems)
                    ReDim Preserve arreDefH(totalItems)
                    ReDim Preserve arreDefP(totalItems)
                    arreDefD(totalItems - 1) = rangosPositivosD(i)
                    arreDefH(totalItems - 1) = rangosPositivosH(i)
                    arreDefP(totalItems - 1) = "S"
                ElseIf rangosPositivosH(i) > arreDefH(totalItems - 1) Then
                    arreDefH(totalItems - 1) = rangosPositivosH(i)
                End If
            Next
            If rangoNegativo > 0 Then
                For i = 0 To rangoNegativo - 1
                    For j = 0 To totalItems - 1
                        If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefH(j) Then
                            arreDefP(j) = "N"
                        End If
                        If rangosNegativosD(i) <= arreDefD(j) And rangosNegativosH(i) >= arreDefD(j) Then
                            arreDefD(j) = rangosNegativosH(i)
                        End If
                        If rangosNegativosD(i) >= arreDefD(j) And rangosNegativosD(i) < arreDefH(j) And rangosNegativosH(i) >= arreDefH(j) Then
                            arreDefH(j) = rangosNegativosD(i)
                        End If
                        If rangosNegativosD(i) > arreDefD(j) And rangosNegativosH(i) < arreDefH(j) Then
                            totalItems = totalItems + 1
                            ReDim Preserve arreDefD(totalItems)
                            ReDim Preserve arreDefH(totalItems)
                            ReDim Preserve arreDefP(totalItems)
                            arreDefD(totalItems - 1) = rangosNegativosH(i)
                            arreDefH(totalItems - 1) = arreDefH(j)
                            arreDefP(totalItems - 1) = "S"
                            arreDefH(j) = rangosNegativosD(i)
                        End If
                    Next
                Next
            End If
            If totalItems > 0 Then
                Dim swap1 = 0
                Dim swap2 = 0
                Dim swap3 = 0
                For i = 0 To totalItems - 1
                    For j = 0 To totalItems - 2
                        If arreDefD(j) > arreDefD(j + 1) Then
                            swap1 = arreDefD(j)
                            swap2 = arreDefH(j)
                            swap3 = arreDefP(j)
                            arreDefD(j) = arreDefD(j + 1)
                            arreDefH(j) = arreDefH(j + 1)
                            arreDefP(j) = arreDefP(i + 1)
                            arreDefD(j + 1) = swap1
                            arreDefH(j + 1) = swap2
                            arreDefP(j + 1) = swap3
                        End If
                    Next
                Next
                tDisponible = 0
                For i = 0 To totalItems - 1
                    If arreDefP(i) = "S" Then
                        horaDesde = arreDefD(i)
                        tDisponible = tDisponible + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaD, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(fechaD, "yyyy/MM/dd") & " " & arreDefH(i)))
                    End If
                Next
            End If
        End If

        horaDesde = Format(fechaD, "HH:mm:ss")
        horaHasta = Format(fechaH, "HH:mm:ss")
        ReDim Preserve rangosPositivosD(0)
        ReDim Preserve rangosPositivosH(0)
        rangoPositivo = 0
        rangoNegativo = 0
        primerRegistro = False


        cadSQL = "SELECT desde, hasta FROM " & rutaBD & ".detalleparos WHERE maquina = " & maquina & " AND estado = 'L' AND estatus = 'A' AND ('" & Format(fechaH, "yyyy/MM/dd HH:mm:ss") & "' <= hasta OR '" & Format(fechaH, "yyyy/MM/dd HH:mm:ss") & "' >= desde) ORDER BY desde, hasta"
        horarios = consultaSEL(cadSQL)
        primerRegistro = True
        If horarios.Tables(0).Rows.Count > 0 Then
            For Each rango In horarios.Tables(0).Rows
                If rango!desde < fechaH And rango!hasta > fechaD Then
                    If primerRegistro Then
                        primerRegistro = False
                        rangoPositivo = rangoPositivo + 1
                        If Format(rango!desde, "HH:mm:ss") < horaDesde Then
                            rangosPositivosD(rangoPositivo - 1) = horaDesde
                        Else
                            rangosPositivosD(rangoPositivo - 1) = Format(rango!desde, "HH:mm:ss")
                        End If
                        If Format(rango!hasta, "HH:mm:ss") > horaHasta Then
                            rangosPositivosH(rangoPositivo - 1) = horaHasta
                        Else
                            rangosPositivosH(rangoPositivo - 1) = Format(rango!hasta, "HH:mm:ss")
                        End If
                    Else
                        rangoPositivo = rangoPositivo + 1
                        ReDim Preserve rangosPositivosD(rangoPositivo)
                        ReDim Preserve rangosPositivosH(rangoPositivo)
                        If Format(rango!desde, "HH:mm:ss") < horaDesde Then
                            rangosPositivosD(rangoPositivo - 1) = horaDesde
                        Else
                            rangosPositivosD(rangoPositivo - 1) = Format(rango!desde, "HH:mm:ss")
                        End If
                        If Format(rango!hasta, "HH:mm:ss") > horaHasta Then
                            rangosPositivosH(rangoPositivo - 1) = horaHasta
                        Else
                            rangosPositivosH(rangoPositivo - 1) = Format(rango!hasta, "HH:mm:ss")
                        End If
                    End If
                End If
            Next
        End If
        'Se crear un registro único por día
        If rangoPositivo > 0 Then
            Dim arreDefD(0) As String
            Dim arreDefH(0) As String
            Dim arreDefP(0) As String
            Dim totalItems = 1
            arreDefD(0) = rangosPositivosD(0)
            arreDefH(0) = rangosPositivosH(0)
            arreDefP(0) = "S"
            For i = 0 To rangoPositivo - 1
                If rangosPositivosD(i) > arreDefD(totalItems - 1) Then
                    totalItems = totalItems + 1
                    ReDim Preserve arreDefD(totalItems)
                    ReDim Preserve arreDefH(totalItems)
                    ReDim Preserve arreDefP(totalItems)
                    arreDefD(totalItems - 1) = rangosPositivosD(i)
                    arreDefH(totalItems - 1) = rangosPositivosH(i)
                    arreDefP(totalItems - 1) = "S"
                ElseIf rangosPositivosH(i) > arreDefH(totalItems - 1) Then
                    arreDefH(totalItems - 1) = rangosPositivosH(i)
                End If
            Next
            If totalItems > 0 Then
                Dim swap1 = 0
                Dim swap2 = 0
                Dim swap3 = 0
                For i = 0 To totalItems - 1
                    For j = 0 To totalItems - 2
                        If arreDefD(j) > arreDefD(j + 1) Then
                            swap1 = arreDefD(j)
                            swap2 = arreDefH(j)
                            swap3 = arreDefP(j)
                            arreDefD(j) = arreDefD(j + 1)
                            arreDefH(j) = arreDefH(j + 1)
                            arreDefP(j) = arreDefP(i + 1)
                            arreDefD(j + 1) = swap1
                            arreDefH(j + 1) = swap2
                            arreDefP(j + 1) = swap3
                        End If
                    Next
                Next
                For i = 0 To totalItems - 1
                    If arreDefP(i) = "S" Then
                        horaDesde = arreDefD(i)
                        tMantto = tMantto + DateDiff(DateInterval.Second, Convert.ToDateTime(Format(fechaD, "yyyy/MM/dd") & " " & arreDefD(i)), Convert.ToDateTime(Format(fechaD, "yyyy/MM/dd") & " " & arreDefH(i)))
                    End If
                Next
            End If
        End If

    End Sub

    Sub crearProgramacion(maquina As Long, lote As Long, parte As Long, fecha As DateTime, proceso As Long, cantidad As Double, van As Double)
        Dim regsAfectados = 0
        Dim cadSQL = ""
        If cantidad - van <= 0 Then Exit Sub

        Dim rateEquipo As Double = 1
        Dim medTiempo As Long = 0
        regsAfectados = consultaACT("UPDATE " & rutaBD & ".horaxhora SET estatus = 'C' WHERE (dia = '" & Format(fecha, "yyyy/MM/dd") & "' AND hora >= " & Format(fecha, "HH") & " OR dia > '" & Format(fecha, "yyyy/MM/dd") & "') AND equipo = " & maquina & " AND lote = " & lote & " AND estatus = 'A'")

        cadSQL = "SELECT plan_van FROM " & rutaBD & ".horaxhora WHERE (dia = '" & Format(fecha, "yyyy/MM/dd") & "' AND hora <= " & Format(fecha, "HH") & " OR dia < '" & Format(fecha, "yyyy/MM/dd") & "') AND equipo = " & maquina & " AND lote = " & lote & " AND estatus = 'A' OR estatus = 'Z' ORDER BY dia, desde DESC LIMIT 1"
        Dim vanPlan = 0
        Dim general As DataSet = consultaSEL(cadSQL)
        If general.Tables(0).Rows.Count > 0 Then
            vanPlan = general.Tables(0).Rows(0)!plan_van
        End If



        cadSQL = "SELECT piezas, tiempo, unidad FROM " & rutaBD & ".relacion_partes_equipos WHERE (equipo = " & maquina & " Or equipo = 0) AND (parte = " & parte & " Or parte = 0) ORDER BY parte DESC, equipo DESC LIMIT 1"
        general = consultaSEL(cadSQL)
        If general.Tables(0).Rows.Count > 0 Then
            rateEquipo = general.Tables(0).Rows(0)!piezas
            medTiempo = general.Tables(0).Rows(0)!tiempo
        End If

        If rateEquipo = 0 Then rateEquipo = 1
        Dim TC As Double = 0
        If rateEquipo = 0 Then
            TC = 1
        ElseIf medTiempo = 2 Then
            TC = 3600 / rateEquipo
        ElseIf medTiempo = 1 Then
            TC = 60 / rateEquipo
            ' TC = rateEquipo / 60
        ElseIf medTiempo = 0 Then
            TC = 1 / rateEquipo
        ElseIf medTiempo = 3 Then
            TC = 86400 / rateEquipo
        End If

        TC = Math.Round(TC, 10)

        Dim completado = False
        Dim fechaInicial = fecha
        Dim cantidadInicial = cantidad - van

        Dim primeraVez = True
        Dim secuenciaTurno
        Dim BTurno As Integer = 0
        Dim elTurno = 0
        Do While Not completado

            Dim termina = Format(fechaInicial, "HH") & ":59:59"
            Dim inicia = Format(fechaInicial, "HH:mm:ss")
            elTurno = -1
            'Saber en que turno está
            cadSQL = "SELECT id, inicia, termina, nombre FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A' AND (inicia <= '" & inicia & "' OR termina >= '" & inicia & "') ORDER BY inicia, termina"
            Dim horarios As DataSet = consultaSEL(cadSQL)
            If horarios.Tables(0).Rows.Count > 0 Then
                For Each horario In horarios.Tables(0).Rows
                    If inicia >= horario!inicia.ToString And inicia <= horario!termina.ToString Then
                        elTurno = horario!id
                        Exit For
                    End If
                Next
                If elTurno = -1 Then
                    For Each horario In horarios.Tables(0).Rows
                        If (inicia >= horario!inicia.ToString Or inicia <= horario!termina.ToString) And horario!termina.ToString < horario!inicia.ToString Then
                            elTurno = horario!id
                            Exit For
                        End If
                    Next

                End If

            End If

            Dim plan = 0
            cadSQL = "SELECT secuencia, termina FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A' AND termina > '" & inicia & "' AND termina < '" & termina & "'"
            Dim turnos As DataSet = consultaSEL(cadSQL)
            If turnos.Tables(0).Rows.Count > 0 Then
                termina = turnos.Tables(0).Rows(0)!termina.ToString
                secuenciaTurno = turnos.Tables(0).Rows(0)!secuencia
                BTurno = 1
            End If

            calcularTiempos(Convert.ToDateTime(Format(fechaInicial, "yyyy/MM/dd") & " " & inicia), Convert.ToDateTime(Format(fechaInicial, "yyyy/MM/dd") & " " & termina), proceso, maquina)
            If (tDisponible - tMantto) > TC Then
                plan = Math.Round((tDisponible - tMantto) / TC, 0)
                If cantidadInicial - plan < 0 Then
                    plan = cantidadInicial
                End If
            End If
            'If tDisponible Or tMantto > 0 Or BTurno > 0 Then
            vanPlan = vanPlan + plan
            regsAfectados = consultaACT("INSERT INTO " & rutaBD & ".horaxhora (equipo, dia, hora, desde, hasta, plan, lote, parte, tc, disponible, mantto, plan_van, ruptura, turno) VALUES (" & maquina & ", '" & Format(fechaInicial, "yyyy/MM/dd") & "', " & Strings.Left(inicia, 2) & ", '" & inicia & "', '" & termina & "', " & plan & ", " & lote & ", " & parte & ", " & TC & ", " & IIf(tDisponible = 3599, 3600, tDisponible) & ", " & IIf(tMantto = 3599, 3600, tMantto) & ", " & vanPlan & ", " & IIf(BTurno > 0, 1, 0) & ", " & elTurno & ")")
            cantidadInicial = cantidadInicial - plan
            ' End If
            If BTurno = 2 Then
                BTurno = 0
            End If

            If cantidadInicial <= 0 Then
                completado = True
            Else
                If BTurno = 1 Then
                    cadSQL = "SELECT inicia FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A' AND secuencia > " & secuenciaTurno & " ORDER BY secuencia LIMIT 1"
                    turnos = consultaSEL(cadSQL)
                    If turnos.Tables(0).Rows.Count > 0 Then
                        inicia = turnos.Tables(0).Rows(0)!inicia.ToString
                    Else
                        cadSQL = "SELECT inicia FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A'ORDER BY secuencia LIMIT 1"
                        turnos = consultaSEL(cadSQL)
                        If turnos.Tables(0).Rows.Count > 0 Then
                            inicia = turnos.Tables(0).Rows(0)!inicia.ToString
                        End If
                    End If
                    BTurno = 2
                    fechaInicial = Convert.ToDateTime(Format(fechaInicial, "yyyy/MM/dd") & " " & inicia)
                Else
                    fechaInicial = DateAndTime.DateAdd(DateInterval.Second, 1, Convert.ToDateTime(Format(fechaInicial, "yyyy/MM/dd") & " " & termina))
                End If
            End If
        Loop
        regsAfectados = consultaACT("UPDATE " & rutaBD & ".cat_maquinas SET replanear = 'Y' WHERE id = " & maquina)
        Dim hxh = consultaSEL(cadSQL)
    End Sub


    Sub compaginar(maquina As Long, lote As Long, fecha As String, hora As Integer)
        Dim iniPlan = 0
        Dim regsAfectados = 0
        Dim cadSQL = "SELECT plan_van FROM " & rutaBD & ".horaxhora WHERE (dia = '" & fecha & "' AND hora <= " & hora & " OR dia < '" & fecha & "') AND equipo = " & maquina & " AND lote = " & lote & " AND estatus = 'A' OR estatus = 'Z' ORDER BY dia, desde DESC LIMIT 1"
        Dim general As DataSet = consultaSEL(cadSQL)
        If general.Tables(0).Rows.Count > 0 Then
            iniPlan = general.Tables(0).Rows(0)!plan_van
        End If
        cadSQL = "SELECT id, plan, tocada, turno, desde FROM " & rutaBD & ".horaxhora WHERE (dia = '" & fecha & "' AND hora >= " & hora & " OR dia > '" & fecha & "') AND equipo = " & maquina & " AND lote = " & lote & " AND estatus = 'A' ORDER BY dia, hora"
        general = consultaSEL(cadSQL)
        Dim cadAdic = ""
        If general.Tables(0).Rows.Count > 0 Then
            For Each plan In general.Tables(0).Rows
                iniPlan = iniPlan + plan!plan
                Dim elTurno = -1
                If plan!tocada = 0 And plan!turno Then
                    'Saber en que turno está
                    cadSQL = "SELECT id, inicia, termina FROM " & rutaBD & ".cat_turnos WHERE estatus = 'A' AND (inicia <= '" & plan!desde & "' OR termina >= '" & plan!desde & "') ORDER BY inicia, termina"
                    Dim horarios As DataSet = consultaSEL(cadSQL)
                    If horarios.Tables(0).Rows.Count > 0 Then
                        For Each horario In horarios.Tables(0).Rows
                            If plan!desde >= horario!inicia.ToString And plan!desde <= horario!termina.ToString Then
                                elTurno = horario!id
                                Exit For
                            End If
                        Next
                        If elTurno = -1 Then
                            For Each horario In horarios.Tables(0).Rows
                                If (plan!desde >= horario!inicia.ToString Or plan!desde <= horario!termina.ToString) And horario!termina.ToString < horario!inicia.ToString Then
                                    elTurno = horario!id
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    If elTurno > -1 Then
                        cadAdic = cadAdic & "UPDATE " & rutaBD & ".horaxhora SET plan_van = " & iniPlan & ", turno = " & elTurno & " WHERE id = " & plan!id & ";"
                    End If
                Else
                    cadAdic = cadAdic & "UPDATE " & rutaBD & ".horaxhora SET plan_van = " & iniPlan & " WHERE id = " & plan!id & ";"
                End If
            Next
        End If
        regsAfectados = consultaACT(cadAdic & "UPDATE " & rutaBD & ".cat_maquinas SET compaginar = 'N' WHERE id = " & maquina)
    End Sub


    Sub mensajeEmergencia(accion As Boolean)
        Try
            agregarSolo(IIf(accion, traduccion(187), traduccion(188)))
            Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & IIf(accion, "P_EMERGENCIA", "R_EMERGENCIA") & Chr(34), AppWinStyle.MinimizedNoFocus)
        Catch ex As Exception
            agregarLOG(traduccion(28) & "mensajes.exe" & traduccion(29) & ex.Message, 7, 0)
        End Try
    End Sub

    Sub mensajeOPCompleta(orden As Long)
        Try
            agregarSolo(traduccion(190))
            Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "completada" & ";" & orden & Chr(34), AppWinStyle.MinimizedNoFocus)
        Catch ex As Exception
            agregarLOG(traduccion(28) & "mensajes.exe" & traduccion(29) & ex.Message, 7, 0)
        End Try
    End Sub
    Sub mensajeOPCambio(entrante As Long, saliente As Long)
        Try
            agregarSolo(traduccion(189))
            Shell(Application.StartupPath & "\mensajes.exe " & Chr(34) & cadenaConexion & Chr(34) & " " & Chr(34) & "cambio" & ";" & saliente & ";" & entrante & Chr(34), AppWinStyle.MinimizedNoFocus)
        Catch ex As Exception
            agregarLOG(traduccion(28) & "mensajes.exe" & traduccion(29) & ex.Message, 7, 0)
        End Try
    End Sub

End Class

