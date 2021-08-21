Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports DevExpress.XtraEditors

Public Class XtraForm2
    Dim cronometro = 16
    Private Sub SimpleButton2_Click(sender As Object, e As EventArgs) Handles SimpleButton2.Click
        Me.Close()
    End Sub

    Private Sub SimpleButton1_Click(sender As Object, e As EventArgs) Handles SimpleButton1.Click
        'Dim cadSQL As String = "SELECT estatus, rol, referencia, nombre FROM " & rutabd & ".cat_usuarios WHERE referencia = '" & TextEdit1.Text & "' AND clave = '" & Cifrado(1, 1, TextEdit2.Text, "VW&Crono", "VW&Crono") & "'"
        Dim cadSQL As String = "SELECT estatus, rol, referencia, nombre FROM " & rutabd & ".cat_usuarios WHERE referencia = '" & TextEdit1.Text & "' AND AES_DECRYPT(clave, 'CronoEIntelraciVn201i') = '" & TextEdit2.Text & "'"
        Dim reader As DataSet = consultaSEL(cadSQL)

        If reader.Tables(0).Rows.Count > 0 Then
            If ValNull(reader.Tables(0).Rows(0)!estatus, "A") <> "A" Then
                XtraMessageBox.Show(traduccion(115), traduccion(116), MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextEdit1.Focus()
            ElseIf ValNull(reader.Tables(0).Rows(0)!rol, "A") <> "A" Then
                XtraMessageBox.Show(traduccion(117), traduccion(118), MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextEdit1.Focus()
            Else

                autenticado = True
            End If
            Me.Close()
        Else
            XtraMessageBox.Show(traduccion(119), traduccion(120), MessageBoxButtons.OK, MessageBoxIcon.Error)
            TextEdit1.Focus()
        End If
    End Sub

    Private Sub TextEdit1_EditValueChanged(sender As Object, e As EventArgs) Handles TextEdit1.EditValueChanged
        SimpleButton1.Enabled = TextEdit1.Text.Length > 0 And TextEdit2.Text.Length > 0
    End Sub

    Private Sub TextEdit2_EditValueChanged(sender As Object, e As EventArgs) Handles TextEdit2.EditValueChanged
        SimpleButton1.Enabled = TextEdit1.Text.Length > 0 And TextEdit2.Text.Length > 0
    End Sub

    Private Sub XtraForm2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = traduccion(121)
        SimpleButton1.Text = traduccion(122)
        SimpleButton2.Text = traduccion(123)
        GroupControl1.Text = traduccion(124)
        LabelControl4.Text = traduccion(125)
        LabelControl1.Text = traduccion(126)
    End Sub

    Private Sub TextEdit1_GotFocus(sender As Object, e As EventArgs) Handles TextEdit1.GotFocus
        TextEdit1.SelectAll()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        cronometro = cronometro - 1
        'LabelControl2.Text = traduccion(121).Replace("campo_0", cronometro)
        If cronometro <= 0 Then
            Me.Close()
        End If
    End Sub
End Class