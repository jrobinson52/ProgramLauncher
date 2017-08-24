Imports System.IO

Public Class URLs
    ' Clicking Close
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        ' Un-hide main form
        Form1.Visible = True

        ' Close this form
        Close()
    End Sub

    ' Clicking update
    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        ' store urls in settings
        My.Settings.URL1 = txtURL1.Text
        My.Settings.URl2 = txtURL2.Text
        My.Settings.URL3 = txtURL3.Text
        My.Settings.URL4 = txtURL4.Text
        My.Settings.URL5 = txtURL5.Text

        ' Indicate that settings have been changed
        My.Settings.ChangedURL = True

        ' Un-hide main form
        Form1.Visible = True

        ' Close this form
        Close()
    End Sub

    ' When the form closes
    Private Sub URLs_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If My.Settings.ChangedURL = False Then
            ' do nothing if nothing has changed

        ElseIf My.Settings.ChangedURL = True Then
            ' create or overwrite existing file
            Dim batch As StreamWriter
            batch = File.CreateText(".\Dependants\IEPages.bat")

            ' write to file
            batch.WriteLine("@echo off")
            If txtURL1.Text <> String.Empty Then
                Try
                    batch.WriteLine("start " & Chr(34) & Chr(34) & " " & Chr(34) & txtURL1.Text & Chr(34))
                    batch.WriteLine("ping -n 4 localhost > nul")
                Catch ex As Exception
                    MessageBox.Show("Something went wrong with URL 1.")
                End Try
            End If

            If txtURL2.Text <> String.Empty Then
                Try
                    batch.WriteLine("start " & Chr(34) & Chr(34) & " " & Chr(34) & txtURL2.Text & Chr(34))
                    batch.WriteLine("ping -n 4 localhost > nul")
                Catch ex As Exception
                    MessageBox.Show("Something went wrong with URL 2.")
                End Try
            End If

            If txtURL3.Text <> String.Empty Then
                Try
                    batch.WriteLine("start " & Chr(34) & Chr(34) & " " & Chr(34) & txtURL3.Text & Chr(34))
                    batch.WriteLine("ping -n 4 localhost > nul")
                Catch ex As Exception
                    MessageBox.Show("Something went wrong with URL 3.")
                End Try
            End If

            If txtURL4.Text <> String.Empty Then
                Try
                    batch.WriteLine("start " & Chr(34) & Chr(34) & " " & Chr(34) & txtURL4.Text & Chr(34))
                    batch.WriteLine("ping -n  localhost > nul")
                Catch ex As Exception
                    MessageBox.Show("Something went wrong with URL 4.")
                End Try
            End If

            If txtURL5.Text <> String.Empty Then
                Try
                    batch.WriteLine("start " & Chr(34) & Chr(34) & " " & Chr(34) & txtURL5.Text & Chr(34))
                    batch.WriteLine("ping -n 4 localhost > nul")
                Catch ex As Exception
                    MessageBox.Show("Something went wrong with URL 5.")
                End Try
            End If

            batch.Close()
        End If
    End Sub

    ' When the form loads
    Private Sub URLs_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' load any saved urls
        txtURL1.Text = My.Settings.URL1
        txtURL2.Text = My.Settings.URl2
        txtURL3.Text = My.Settings.URL3
        txtURL4.Text = My.Settings.URL4
        txtURL5.Text = My.Settings.URL5
    End Sub

    ' When text is entered into the text boxes
    Private Sub txtURL1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtURL1.KeyDown, txtURL2.KeyDown, txtURL3.KeyDown, txtURL4.KeyDown, txtURL5.KeyDown
        ' prevent spaces in url text boxes
        Select Case e.KeyValue
            Case Keys.Space : e.SuppressKeyPress = True
        End Select
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' clear text boxes and stored settings
        txtURL1.Text = String.Empty
        My.Settings.URL1 = String.Empty
        txtURL2.Text = String.Empty
        My.Settings.URl2 = String.Empty
        txtURL3.Text = String.Empty
        My.Settings.URL3 = String.Empty
        txtURL4.Text = String.Empty
        My.Settings.URL4 = String.Empty
        txtURL5.Text = String.Empty
        My.Settings.URL5 = String.Empty
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        ' set default tabs
        My.Settings.URL1 = "https://tmna.service-now.com/navpage.do"
        My.Settings.URl2 = "https://tmna.service-now.com/navpage.do"
        My.Settings.URL3 = "https://tmna.service-now.com/$knowledge.do"
        My.Settings.URL4 = "https://tmna.service-now.com/1ts/"
        My.Settings.URL5 = String.Empty

        txtURL1.Text = My.Settings.URL1
        txtURL2.Text = My.Settings.URl2
        txtURL3.Text = My.Settings.URL3
        txtURL4.Text = My.Settings.URL4
        txtURL5.Text = My.Settings.URL5
    End Sub
End Class