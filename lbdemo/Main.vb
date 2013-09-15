Public Class Main
    Dim Script As lbscript.LBScript = New lbscript.LBScript
    Dim source
    Dim debug As Boolean = False

    

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If MsgBox("Exit application ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton1, "Exit") = MsgBoxResult.Yes Then
            End
        End If
    End Sub



    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        dlg.ShowDialog()
    End Sub


    Private Sub dlg_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles dlg.FileOk
        source = System.IO.File.ReadAllText(dlg.FileName)
        srcScript.Text = source
    End Sub

    Private Sub StartScriptToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StartScriptToolStripMenuItem.Click
        dlg.ShowDialog()
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim s() As String = System.Environment.GetCommandLineArgs()
        If (s.Count > 0) Then
            For i = 1 To s.Count - 1
                Select Case s(i)
                    Case "-load"
                        i = i + 1
                        dlg.FileName = s(i)
                        source = System.IO.File.ReadAllText(s(i))
                        srcScript.Text = source

                    Case "-run"
                        i = i + 1
                        dlg.FileName = s(i)
                        source = System.IO.File.ReadAllText(s(i))
                        srcScript.Text = source
                        StartScriptToolStripMenuItem1.PerformClick()
                    Case "-h", "-help", "-?"
                        'Case "-logon"
                        'Case "-u", "--user"
                        'Case "-p", "--password"
                        'Case "-c", "--client"
                End Select
            Next
        End If
    End Sub

    Private Sub Main_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        'If MsgBox("Exit application ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton1, "Exit") = MsgBoxResult.Yes Then End
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        sdlg.ShowDialog()
    End Sub


    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        System.IO.File.Copy(dlg.FileName, dlg.FileName & ".bak")
        System.IO.File.Delete(dlg.FileName)
        System.IO.File.WriteAllText(dlg.FileName, srcScript.Text)
    End Sub

    Private Sub sdlg_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles sdlg.FileOk
        System.IO.File.WriteAllText(sdlg.FileName, srcScript.Text)
        dlg.FileName = sdlg.FileName
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click

        If dlg.FileName.Length > 1 Then
            System.IO.File.Copy(dlg.FileName, dlg.FileName & ".bak", True)
            System.IO.File.Delete(dlg.FileName)
            System.IO.File.WriteAllText(dlg.FileName, srcScript.Text)
        Else
            sdlg.ShowDialog()
        End If

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        sdlg.ShowDialog()
    End Sub

    Private Sub StopScriptToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StopScriptToolStripMenuItem.Click
        Script.load(srcScript.Text)
        Script.run()
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Script.load(srcScript.Text)
        Script.run()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If MsgBox("Exit application ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton1, "Exit") = MsgBoxResult.No Then
            e.Cancel = True
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        AboutBox.ShowDialog()
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        LoadScriptToolStripMenuItem.PerformClick()
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        AboutBox.ShowDialog()
    End Sub

    Private Sub LoadScriptToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadScriptToolStripMenuItem.Click
        dlg.ShowDialog()
    End Sub

    Private Sub StartScriptToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles StartScriptToolStripMenuItem1.Click
        'start script
        'Me.Hide()
        Script.load(srcScript.Text)
        Script.run()
        'Me.Show()
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        'save
        If dlg.FileName.Length > 1 Then
            System.IO.File.Copy(dlg.FileName, dlg.FileName & ".bak", True)
            System.IO.File.Delete(dlg.FileName)
            System.IO.File.WriteAllText(dlg.FileName, srcScript.Text)
        Else
            sdlg.ShowDialog()
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem1.Click
        sdlg.ShowDialog()
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        'exit
        If MsgBox("Exit application ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information + MsgBoxStyle.DefaultButton1, "Exit") = MsgBoxResult.Yes Then
            End
        End If
    End Sub

    Private Sub ToolStripButton8_Click(sender As Object, e As EventArgs) Handles ToolStripButton8.Click
        StartScriptToolStripMenuItem1.PerformClick()
    End Sub

    Private Sub ToolStripButton7_Click(sender As Object, e As EventArgs) Handles ToolStripButton7.Click
        SaveAsToolStripMenuItem1.PerformClick()
    End Sub

    Private Sub ToolStripButton6_Click(sender As Object, e As EventArgs) Handles ToolStripButton6.Click
        SaveToolStripMenuItem1.PerformClick()
    End Sub

    Private Sub ContentsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ContentsToolStripMenuItem1.Click
        Process.Start(System.Environment.CurrentDirectory + "\LBScript.chm")
    End Sub
End Class