Imports MySql.Data.MySqlClient
Imports DevExpress.XtraEditors
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Module basico

    Public errorBD As String
    Public horaDesde As DateTime
    Public ultimaFalla
    Public autenticado As Boolean
    Public usuarioCerrar As String
    Public cadenaConexion As String
    Public cadenaConexionMMCALL As String
    Public cadenaConexionFALLAS As String
    Public rutaBD As String = "robot"
    Public rutaMMCALL As String = "mmcall"
    Public miTurno As Long
    Public miSecuencia As Long

    Public esaqui As Boolean = False
    Public traduccion As String()


    Public Function conexion(miCadenaConexion As String) As MySqlConnection
        conexion = Nothing
        errorBD = ""
        Try
            conexion = New MySqlConnection

            conexion.ConnectionString = miCadenaConexion
            conexion.Open()
        Catch ex As Exception
            errorBD = ex.Message
        End Try
    End Function


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
                consultaACT = -1
            End Try

        End If
        miConexion.Dispose()
        miConexion.Close()
        miConexion = Nothing

    End Function

    Public Function consultaSEL(cadena As String, Optional miCadenaConexion As String = "") As DataSet
        Dim miConexion = New MySqlConnection

        Try
            errorBD = ""
            If miCadenaConexion = "" Then
                miCadenaConexion = cadenaConexion
            End If
            miConexion.ConnectionString = miCadenaConexion

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
        Catch ex As Exception

            errorBD = ex.Message
        Finally
            miConexion.Dispose()
            miConexion.Close()
            miConexion = Nothing
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


    Sub Sombrear_Texto(ByVal sender As Object, ByVal e As EventArgs)
        TryCast(sender, TextEdit).SelectAll()
    End Sub


    Public Function Cifrado(ByVal modo As Byte, ByVal Algoritmo As Byte, ByVal cadena As String, ByVal key As String, ByVal VecI As String) As String

        Dim plaintext() As Byte

        If modo = 1 Then

            plaintext = Encoding.ASCII.GetBytes(cadena)

        ElseIf modo = 2 Then

            plaintext = Convert.FromBase64String(cadena)

        End If

        Dim keys() As Byte = Encoding.ASCII.GetBytes(key)

        Dim memdata As New MemoryStream

        Dim transforma As ICryptoTransform

        Select Case Algoritmo

            Case 1

                Dim des As New DESCryptoServiceProvider ' DES

                des.Mode = CipherMode.CBC
                If modo = 1 Then

                    transforma = des.CreateEncryptor(keys, Encoding.ASCII.GetBytes(VecI))

                ElseIf modo = 2 Then

                    transforma = des.CreateDecryptor(keys, Encoding.ASCII.GetBytes(VecI))

                End If

            Case 2

                Dim des3 As New TripleDESCryptoServiceProvider 'TripleDES

                des3.Mode = CipherMode.CBC
                If modo = 1 Then

                    transforma = des3.CreateEncryptor(keys, Encoding.ASCII.GetBytes(VecI))

                ElseIf modo = 2 Then

                    transforma = des3.CreateDecryptor(keys, Encoding.ASCII.GetBytes(VecI))

                End If

            Case 3

                Dim rc2 As New RC2CryptoServiceProvider 'RC2

                rc2.Mode = CipherMode.CBC
                If modo = 1 Then

                    transforma = rc2.CreateEncryptor(keys, Encoding.ASCII.GetBytes(VecI))

                ElseIf modo = 2 Then

                    transforma = rc2.CreateDecryptor(keys, Encoding.ASCII.GetBytes(VecI))

                End If

            Case 4

                Dim rj As New RijndaelManaged 'Rijndael

                rj.Mode = CipherMode.CBC
                If modo = 1 Then

                    transforma = rj.CreateEncryptor(keys, Encoding.ASCII.GetBytes(VecI))

                ElseIf modo = 2 Then

                    transforma = rj.CreateDecryptor(keys, Encoding.ASCII.GetBytes(VecI))

                End If

        End Select

        Dim encstream As New CryptoStream(memdata, transforma, CryptoStreamMode.Write)

        encstream.Write(plaintext, 0, plaintext.Length)

        encstream.FlushFinalBlock()

        encstream.Close()

        If modo = 1 Then

            cadena = Convert.ToBase64String(memdata.ToArray)

        ElseIf modo = 2 Then

            cadena = Encoding.ASCII.GetString(memdata.ToArray)

        End If

        Return cadena 'Aquí Devuelve los Datos Cifrados

    End Function

End Module


