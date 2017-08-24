
Imports System.IO
Imports System.ComponentModel

Public Class Form1
    ' close button clicked
    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        ' Close the program
        Close()
    End Sub
    ' edit url button clicked
    Private Sub btnURLs_Click(sender As Object, e As EventArgs) Handles btnURLs.Click
        ' Show page to input URLs
        Dim frmURLs As New URLs
        frmURLs.Show()

        ' Hide this form
        Me.Visible = False
    End Sub
    ' launch button clicked
    Private Sub btnLaunch_Click(sender As Object, e As EventArgs) Handles btnLaunch.Click
        ' If IE box checked run IE batch file
        If cboxIE.Checked = True Then
            ' if there are no urls saved
            If My.Settings.URL1 & My.Settings.URl2 & My.Settings.URL3 &
                    My.Settings.URL4 & My.Settings.URL5 = String.Empty Then
                MessageBox.Show("Please select/update internet tabs to open.")
                Exit Sub ' prevent launching other programs and keep this program open
            Else
                Try
                    Process.Start(".\Dependants\IEPages.bat") ' launch generated batch file
                Catch ex As Exception
                    MessageBox.Show("Please update URLs")
                End Try
            End If
        End If
        ' if Interaction client/ desktop checked run it
        If cboxIC.Checked = True Then
            ' determine if desktop or client is installed and run accordingly
            If File.Exists("C:\Program Files (x86)\Interactive Intelligence\ICUserApps\InteractionDesktop.exe") = True Then
                ' -----run desktop-----
                ' runs program
                'Process.Start("C:\Program Files (x86)\Interactive Intelligence\ICUserApps\InteractionDesktop.exe")

                ' runs with autologin
                Dim process As Process = Nothing
                Dim processStartInfo As New ProcessStartInfo()
                processStartInfo.FileName = "C:\Program Files (x86)\Interactive Intelligence\ICUserApps\InteractionDesktop.exe"
                processStartInfo.Arguments = ("-silent")
                Try
                    process = Process.Start(processStartInfo)
                Catch ex As Exception

                Finally
                    If Not (process Is Nothing) Then
                        process.Dispose()
                    End If
                End Try


            ElseIf File.Exists("C:\Program Files\Interactive Intelligence\ICUserApps\InteractionClient.exe") = True Then
                ' -----run client-----
                ' runs normally
                ' Process.Start("C:\Program Files\Interactive Intelligence\ICUserApps\InteractionClient.exe")

                ' runs with autologin
                Dim process As Process = Nothing
                Dim processStartInfo As New ProcessStartInfo()
                processStartInfo.FileName = "C:\Program Files\Interactive Intelligence\ICUserApps\InteractionClient.exe"
                processStartInfo.Arguments = ("-silent")
                Try
                    process = Process.Start(processStartInfo)
                Catch ex As Exception

                Finally
                    If Not (process Is Nothing) Then
                        process.Dispose()
                    End If
                End Try
            End If
        End If
        ' if Outlook checked run outlook 2016
        If cboxOutlook.Checked = True Then
            Try
                Process.Start("C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE")
            Catch ex As Exception
                MessageBox.Show("Sorry, I was unable to open Outlook 2016")
            End Try

        Else
            ' do nothing
        End If
        ' if rodney's tool checked run it
        If cboxSDT.Checked = True Then
            Try
                Process.Start("\\T01fap01\isdata\Central Operations\SAIC\Toyota Service Desk tool\Toyota Service Desk Tool.application")
            Catch ex As Exception
                MessageBox.Show("Sorry, I was unable to open the Service Desk Tool")
            End Try
        Else
            ' do nothing
        End If
        ' if Software center checked run
        If cboxCCM.Checked = True Then
            Try
                Process.Start("softwarecenter:")
            Catch ex As Exception
                MessageBox.Show("Sorry, I was unable to open Software Center")
            End Try
        Else
            ' do notning
        End If
        ' if powershell checked run it
        If cboxPS.Checked = True Then
            Process.Start("powershell")
        Else
            ' do nothing
        End If
        ' If rumba checked run it (currently only web based)
        If cboxRumba.Checked = True Then
            Dim process As Process = Nothing
            Dim processStartInfo As New ProcessStartInfo()
            processStartInfo.FileName = "http://tmshost/hod/tmsmod2.html?Obplet=applet&Embedded=false&Launch=3270[0020]Display[0020]Mod[0020]2&JavaType=java2"
            processStartInfo.CreateNoWindow = True
            processStartInfo.WindowStyle = ProcessWindowStyle.Hidden
            Try
                process = Process.Start(processStartInfo)
                process.WaitForExit()
            Catch ex As Exception

            Finally
                If Not (process Is Nothing) Then
                    process.Dispose()
                End If
            End Try
        Else
            ' do nothing
        End If
        ' if notepad checked run it
        If cboxNotePad.Checked = True Then
            Process.Start("notepad")
        Else
            'do nothing
        End If
        ' if Paint checked run
        If cboxPaint.Checked = True Then
            Process.Start("mspaint")
        End If
        ' if Bomgar checked run
        If cboxBomgar.Checked = True Then
            Try
                Process.Start("C:\Program Files\Bomgar\Bomgar Representative Console\remote-toyota.naismc.com\bomgar-rep.exe")
            Catch ex As Exception
                MessageBox.Show("Sorry, I was unable to open Bomgar." & vbNewLine & "Is it installed on this computer?")
            End Try
        End If
        'if lotus notes checked run
        If cboxLotus.Checked = True Then
            Try
                Process.Start("C:\Program Files (x86)\IBM\Lotus\Notes\notes.exe")
            Catch ex As Exception
                MessageBox.Show("Sorry, I was unable to open Lotus Notes." & vbNewLine & "Is it installed on this computer?")
            End Try
        End If
        ' if specific files selected open them
        If cboxFile.Checked = True Then
            If My.Settings.File1 <> String.Empty Then
                Try
                    Process.Start(My.Settings.File1)
                Catch ex As Exception
                    MessageBox.Show("Sorry, an error occured trying to open:" & vbNewLine & My.Settings.File1 &
                                     vbNewLine & "Does the file exist?")
                End Try
            End If
            If My.Settings.File2 <> String.Empty Then
                Try
                    Process.Start(My.Settings.File2)
                Catch ex As Exception
                    MessageBox.Show("Sorry, an error occured trying to open:" & vbNewLine & My.Settings.File2 &
                                     vbNewLine & "Does the file exist?")
                End Try
            End If
            If My.Settings.File3 <> String.Empty Then
                Try
                    Process.Start(My.Settings.File3)
                Catch ex As Exception
                    MessageBox.Show("Sorry, an error occured trying to open:" & vbNewLine & My.Settings.File3 &
                                     vbNewLine & "Does the file exist?")
                End Try
            End If
            If My.Settings.File4 <> String.Empty Then
                Try
                    Process.Start(My.Settings.File4)
                Catch ex As Exception
                    MessageBox.Show("Sorry, an error occured trying to open:" & vbNewLine & My.Settings.File4 &
                                     vbNewLine & "Does the file exist?")
                End Try
            End If
            If My.Settings.File5 <> String.Empty Then
                Try
                    Process.Start(My.Settings.File5)
                Catch ex As Exception
                    MessageBox.Show("Sorry, an error occured trying to open:" & vbNewLine & My.Settings.File5 &
                                     vbNewLine & "Does the file exist?")
                End Try
            End If
        End If





        ' close form after launching
        Close()
    End Sub
    ' when form closes
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ' store checked values for next launch
        My.Settings.IE = cboxIE.Checked
        My.Settings.Interaction = cboxIC.Checked
        My.Settings.Outlook = cboxOutlook.Checked
        My.Settings.ServiceDeskTool = cboxSDT.Checked
        My.Settings.SoftwareCenter = cboxCCM.Checked
        My.Settings.PowerShell = cboxPS.Checked
        My.Settings.Rumba = cboxRumba.Checked
        My.Settings.NotePad = cboxNotePad.Checked
        My.Settings.Files = cboxFile.Checked
    End Sub
    ' when form loads
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' retrieve previously checked boxes
        cboxIE.Checked = My.Settings.IE
        cboxIC.Checked = My.Settings.Interaction
        cboxOutlook.Checked = My.Settings.Outlook
        cboxSDT.Checked = My.Settings.ServiceDeskTool
        cboxCCM.Checked = My.Settings.SoftwareCenter
        cboxPS.Checked = My.Settings.PowerShell
        cboxRumba.Checked = My.Settings.Rumba
        cboxNotePad.Checked = My.Settings.NotePad
        cboxFile.Checked = My.Settings.Files
        ' set welcome message
        If TimeOfDay.Hour <= 10 Then
            lblWelcome.Text = "Good Morning!"
        ElseIf TimeOfDay.Hour <= 16 AndAlso TimeOfDay.Hour > 10 Then
            lblWelcome.Text = "Good Afternoon!"
        ElseIf TimeOfDay.Hour >= 17 Then
            lblWelcome.Text = "Good Evening!"
        End If
    End Sub

    Private Sub btnAbout_Click(sender As Object, e As EventArgs) Handles btnAbout.Click
        ' open about box
        AboutBox1.Show()
    End Sub

    Private Sub btnFiles_Click(sender As Object, e As EventArgs) Handles btnFiles.Click
        ' open form to select files
        FileSelect.Show()
        Me.Visible = False
    End Sub
End Class
