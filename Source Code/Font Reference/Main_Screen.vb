Imports System.IO
Imports System.Net.Mail

Public Class Main_Screen

    Dim progresslabel As String = ""
    Dim filename As String = ""
    Dim shownminimizetip As Boolean = False
    Dim sentencelist As ArrayList
    Dim sentencelistCurrentIndex As Integer
    Dim pictureboxtop As Integer = 0


    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message()
                If FullErrors_Checkbox.Checked = True Then
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.ToString
                Else
                    Display_Message1.Message_Textbox.Text = "The Application encountered the following problem: " & vbCrLf & identifier_msg & ": " & ex.Message.ToString
                End If
                Display_Message1.Timer1.Interval = 1000
                Display_Message1.ShowDialog()
                Dim dir As System.IO.DirectoryInfo = New System.IO.DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                dir = Nothing
                Dim filewriter As System.IO.StreamWriter = New System.IO.StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & identifier_msg & ": " & ex.ToString)
                filewriter.WriteLine("")
                filewriter.Flush()
                filewriter.Close()
                filewriter = Nothing
                Label2.Text = "Error encountered in last action"
            End If
        Catch exc As Exception
            MsgBox("An error occurred in the application's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

 

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            loadSettings()
            Control.CheckForIllegalCrossThreadCalls = False
            shownminimizetip = False
            Me.Text = My.Application.Info.ProductName & " (" & Format(My.Application.Info.Version.Major, "0000") & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
            LoadSettings()
            sentencelist = New ArrayList
            sentencelist.Add("A large fawn jumped quickly over white zinc boxes.")
            sentencelist.Add("Crazy Fredrick bought many very exquisite opal jewels.")
            sentencelist.Add("Exquisite farm wench gives body jolt to prize stinker.")
            sentencelist.Add("How quickly daft jumping zebras vex.")
            sentencelist.Add("Jump by vow of quick, lazy strength in Oxford.")
            sentencelist.Add("Many-wived Jack laughs at probes of sex quiz.")
            sentencelist.Add("Pack my box with five dozen liquor jugs.")
            sentencelist.Add("Playing jazz vibe chords quickly excites my wife.")
            sentencelist.Add("Quick zephyrs blow, vexing daft Jim.")
            sentencelist.Add("Sphinx of black quartz: judge my vow.")
            sentencelist.Add("Sympathizing would fix Quaker objectives.")
            sentencelist.Add("The five boxing wizards jump quickly.")
            sentencelist.Add("The quick brown fox jumps over a lazy dog.")
            sentencelist.Add("Turgid saxophones blew over Mick's jazzy quaff.")

            sentencelistCurrentIndex = 0

            Dim i As Integer = Asc("A")
            For Each page As TabPage In TabControl1.TabPages
                page.Text = Chr(i)
                page.AutoScroll = True
                i = i + 1
            Next
            Label2.Text = "Application loaded"
            UpdateTabContents(0)

        Catch ex As Exception
            Error_Handler(ex, "Form Load")
        End Try
    End Sub

    
    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Try
            Label2.Text = "About displayed"
            AboutBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display About Screen")
        End Try
    End Sub

    Private Sub HelpToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HelpToolStripMenuItem.Click
        Try
            Label2.Text = "Help displayed"
            HelpBox1.ShowDialog()
        Catch ex As Exception
            Error_Handler(ex, "Display Help Screen")
        End Try
    End Sub

   
    Private Sub LoadSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")
            If My.Computer.FileSystem.FileExists(configfile) Then
                Dim reader As StreamReader = New StreamReader(configfile)
                Dim lineread As String
                Dim variablevalue As String
                While reader.Peek <> -1
                    lineread = reader.ReadLine
                    If lineread.IndexOf("=") <> -1 Then

                        variablevalue = lineread.Remove(0, lineread.IndexOf("=") + 1)

                        If lineread.StartsWith("FullErrors_Checkbox=") Then
                            FullErrors_Checkbox.Checked = variablevalue
                        End If

                    End If
                End While
                reader.Close()
                reader = Nothing
            End If
        Catch ex As Exception
            Error_Handler(ex, "Load Settings")
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim configfile As String = (Application.StartupPath & "\config.sav").Replace("\\", "\")

            Dim writer As StreamWriter = New StreamWriter(configfile, False)

            writer.WriteLine("FullErrors_Checkbox=" & FullErrors_Checkbox.Checked.ToString)

            writer.Flush()
            writer.Close()
            writer = Nothing

        Catch ex As Exception
            Error_Handler(ex, "Save Settings")
        End Try
    End Sub

    Private Sub Main_Screen_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            SaveSettings()
        Catch ex As Exception
            Error_Handler(ex, "Closing Application")
        End Try
    End Sub


    Private Sub AddString(ByVal ff As FontFamily, ByRef y As Int32, ByVal Style As FontStyle, ByVal panel1 As Control)
        Try
            Dim dfnt As Font = New Font("Arial", 8.25, FontStyle.Regular, GraphicsUnit.Point)
            Dim fnt As Font = New Font(ff, 20.25, Style, GraphicsUnit.Point)
            Dim LineSpace As Int32 = CInt((ff.GetLineSpacing(Style)) * fnt.Size / ff.GetEmHeight(Style)) + CInt((ff.GetLineSpacing(Style)) * dfnt.Size / ff.GetEmHeight(Style))


            Dim pictureBox1 As PictureBox = New PictureBox()
            pictureBox1.BorderStyle = BorderStyle.None
            pictureBox1.Height = 55
            pictureBox1.Width = panel1.Width - 22
            Dim B As Bitmap = New Bitmap(pictureBox1.Width, pictureBox1.Height)
            Dim G As Graphics = Graphics.FromImage(B)

            G.DrawString(ff.Name, dfnt, Brushes.SteelBlue, 0, 0)
            G.DrawString(sentencelist.Item(sentencelistCurrentIndex), fnt, Brushes.Black, 0, 12)
            sentencelistCurrentIndex = sentencelistCurrentIndex + 1
            If sentencelistCurrentIndex > (sentencelist.Count - 1) Then
                sentencelistCurrentIndex = 0
            End If
            pictureBox1.Image = B
            pictureBox1.Top = pictureboxtop
            pictureBox1.Left = 2
            pictureboxtop = pictureboxtop + pictureBox1.Height + 8
            panel1.Controls.Add(pictureBox1)
            'panel1.Controls(panel1.Controls.Count - 1).Location = New Point(2, pictureboxtop)


            If y < panel1.Height Then
                panel1.Refresh()
            End If
            dfnt.Dispose()
            fnt.Dispose()
            G.Dispose()
        Catch ex As Exception
            Error_Handler(ex, "Updating Fonts")
        End Try
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Try
            UpdateTabContents(TabControl1.SelectedIndex)
        Catch ex As Exception
            Error_Handler(ex, "Tab Selection Changed")
        End Try
    End Sub

    Private Sub UpdateTabContents(ByVal SelectedIndex As Integer)
        Try
            Label2.Text = ""
            Label2.Refresh()
            ProgressBar1.Visible = True
            pictureboxtop = 6
            Dim counter As Integer = 0

            Dim Style As FontStyle
            Dim y As Int32 = 0
            Dim ff As FontFamily

            Dim percentComplete, inputlines, inputlinesprecount As Integer
            inputlines = 0
            percentComplete = 0
            ProgressBar1.Value = percentComplete

            inputlinesprecount = FontFamily.Families.Length

            For Each ff In FontFamily.Families
                If ff.Name.ToUpper.StartsWith(Chr(SelectedIndex + 65)) Then
                    Style = FontStyle.Regular
                    If ff.IsStyleAvailable(Style) Then
                        AddString(ff, y, Style, TabControl1.TabPages(SelectedIndex))
                        counter = counter + 1
                    End If
                End If
                inputlines = inputlines + 1
                If inputlinesprecount > 0 Then
                    percentComplete = CSng(inputlines) / CSng(inputlinesprecount) * 100
                Else
                    percentComplete = 100
                End If
                If percentComplete > 100 Then
                    percentComplete = 100
                End If
                ProgressBar1.Value = percentComplete
            Next
            If counter <> 1 Then
                Label2.Text = "Font Listing Complete: " & Chr(SelectedIndex + 65) & " (" & counter & " fonts located)"
            Else
                Label2.Text = "Font Listing Complete: " & Chr(SelectedIndex + 65) & " (" & counter & " font located)"
            End If
            ProgressBar1.Visible = False
        Catch ex As Exception
            Error_Handler(ex, "Update Tab Contents")
        End Try
    End Sub
End Class
