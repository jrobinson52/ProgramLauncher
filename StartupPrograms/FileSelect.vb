Imports System.ComponentModel
Imports System.IO


Public Class FileSelect
    Private strfilename1 As String = String.Empty
    Private strfilename2 As String = String.Empty
    Private strfilename3 As String = String.Empty
    Private strfilename4 As String = String.Empty
    Private strfilename5 As String = String.Empty

    Private Sub FileSelect_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' restore filepath settings
        txtFile1.Text = My.Settings.File1
        txtFile2.Text = My.Settings.File2
        txtFile3.Text = My.Settings.File3
        txtFile4.Text = My.Settings.File4
        txtFile5.Text = My.Settings.File5

    End Sub

    Private Sub btnFile1_Click(sender As Object, e As EventArgs) Handles btnFile1.Click
        If ofdFile1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' retrieve selected filename
            strfilename1 = ofdFile1.FileName
            txtFile1.Text = strfilename1
        End If
    End Sub

    Private Sub btnFile2_Click(sender As Object, e As EventArgs) Handles btnFile2.Click
        If ofdFile1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' retrieve selected filename
            strfilename2 = ofdFile1.FileName
            txtFile2.Text = strfilename2
        End If
    End Sub

    Private Sub btnFile3_Click(sender As Object, e As EventArgs) Handles btnFile3.Click
        If ofdFile1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' retrieve selected filename
            strfilename3 = ofdFile1.FileName
            txtFile3.Text = strfilename3
        End If
    End Sub

    Private Sub btnFile4_Click(sender As Object, e As EventArgs) Handles btnFile4.Click
        If ofdFile1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' retrieve selected filename
            strfilename4 = ofdFile1.FileName
            txtFile4.Text = strfilename4
        End If
    End Sub

    Private Sub btnFile5_Click(sender As Object, e As EventArgs) Handles btnFile5.Click
        If ofdFile1.ShowDialog = System.Windows.Forms.DialogResult.OK Then
            ' retrieve selected filename
            strfilename5 = ofdFile1.FileName
            txtFile5.Text = strfilename5
        End If
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        ' store filepath as setting
        My.Settings.File1 = txtFile1.Text
        My.Settings.File2 = txtFile2.Text
        My.Settings.File3 = txtFile3.Text
        My.Settings.File4 = txtFile4.Text
        My.Settings.File5 = txtFile5.Text


        'unhide main form
        Form1.Visible = True

        ' close this window
        Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Form1.Visible = True
        ' close this form
        Close()
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As EventArgs) Handles btnClearAll.Click
        ' clear txt and setting
        txtFile1.Text = String.Empty
        My.Settings.File1 = String.Empty
        txtFile2.Text = String.Empty
        My.Settings.File2 = String.Empty
        txtFile3.Text = String.Empty
        My.Settings.File3 = String.Empty
        txtFile4.Text = String.Empty
        My.Settings.File4 = String.Empty
        txtFile5.Text = String.Empty
        My.Settings.File5 = String.Empty
    End Sub

    Private Sub btnClear1_Click(sender As Object, e As EventArgs) Handles btnClear1.Click
        ' clear txt, string, and setting
        txtFile1.Text = String.Empty
        My.Settings.File1 = String.Empty
    End Sub

    Private Sub btnClear2_Click(sender As Object, e As EventArgs) Handles btnClear2.Click
        ' clear txt, string, and setting
        txtFile2.Text = String.Empty
        My.Settings.File2 = String.Empty
    End Sub

    Private Sub btnClear3_Click(sender As Object, e As EventArgs) Handles btnClear3.Click
        ' clear txt, string, and setting
        txtFile3.Text = String.Empty
        My.Settings.File3 = String.Empty
    End Sub

    Private Sub btnClear4_Click(sender As Object, e As EventArgs) Handles btnClear4.Click
        ' clear txt, string, and setting
        txtFile4.Text = String.Empty
        My.Settings.File4 = String.Empty
    End Sub

    Private Sub btnClear5_Click(sender As Object, e As EventArgs) Handles btnClear5.Click
        ' clear txt, string, and setting
        txtFile5.Text = String.Empty
        My.Settings.File5 = String.Empty
    End Sub
End Class