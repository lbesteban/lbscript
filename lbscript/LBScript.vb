Public Class LBScript
    'Dim sapguiAuto As Object 'SAPFEWSELib.GuiAuto
    Dim app As Object ' SAPFEWSELib.GuiApplication
    Dim conn As Object 'SAPFEWSELib.GuiConnection
    Dim sess As Object 'SAPFEWSELib.GuiSession
    Dim objWrapper As Object ' SapROTWr.SapROTWrapper")
    Dim objRotSAPGUI As Object '  objWrapper.GetROTEntry("SAPGUI")
    Dim endScript As Boolean
    Dim myCode As String()
    Function logon(host As String, system As String, client As String, user As String, password As String, Optional lang As String = "EN") As Boolean
        Return Me.logon("/H/" & host & "/S/32" & system, client, user, password, lang)
    End Function
    Function logon(cmdstring As String, client As String, user As String, Optional password As String = "", Optional lang As String = "EN") As Boolean
        objWrapper = CreateObject("SapROTWr.SapROTWrapper")
        objRotSAPGUI = objWrapper.GetROTEntry("SAPGUI")
        If IsNothing(objRotSAPGUI) Then
            'saplogon was not launched and loadscript(....) will not work
            app = CreateObject("Sapgui.ScriptingCtrl.1")
        Else
            ' sapgui script will have possibility to run more VBS scripts inside
            app = objRotSAPGUI.GetScriptingEngine()
        End If
        'app = New SAPFEWSELib.GuiApplication
        'app = CreateObject("Sapgui.ScriptingCtrl.1")
        conn = app.OpenConnectionByConnectionString(cmdstring)
        sess = conn.Children(0)
        sess.ActiveWindow.FindByName("RSYST-MANDT", "GuiTextField").Text = client
        sess.ActiveWindow.FindByName("RSYST-BNAME", "GuiTextField").Text = user
        sess.ActiveWindow.FindByName("RSYST-BCODE", "GuiPasswordField").Text = password
        sess.ActiveWindow.FindByName("RSYST-LANGU", "GuiTextField").Text = lang
        sess.SendCommand("/0")
        While app.Children(0).Sessions(0).Info.Transaction = "S000"
            ' Decir al sistema que quieres hacer logones multiples
            If InStr(UCase(app.Children(0).Sessions(0).ActiveWindow.Text), "LOGON") > 0 Then
                sess.ActiveWindow.FindByName("MULTI_LOGON_OPT2", "GuiRadioButton").Selected = True
                sess.SendCommand("OK")
            End If
            ' Decir al sistema que ignore los mensajes de bienvenida
            If InStr(UCase(app.Children(0).Sessions(0).ActiveWindow.Text), "SYSTEM MESSAGES") > 0 Then
                sess.SendCommand("&F12")
            End If
            ' Pasar la pantalla SAPMSSY0-0120 screen con okcode WEIT
            If UCase(app.Children(0).Sessions(0).ActiveWindow.Text) = "COPYRIGHT" Then
                app.Children(0).Sessions(0).SendCommand("WEIT")
            End If
            If app.Children(0).Sessions(0).Info.Transaction = "S000" Then
                If (MsgBox("Logon not complete" + vbCrLf + "Finish logon and press ok", vbExclamation + vbOKCancel + vbSystemModal)) = vbCancel Then
                    Return False
                End If
            End If
        End While
        Return True
    End Function
    Function getStatus() As String
        If IsNothing(sess) Then
            Return "session is not connected"
        Else
            Return Me.sess.ActiveWindow.FindByName("sbar", "GuiStatusbar").text
        End If
    End Function
    Function load(txt As String) As Boolean
        myCode = txt.Split(ControlChars.CrLf.ToCharArray(), System.StringSplitOptions.RemoveEmptyEntries)
        Return True
    End Function
    Function loadFile(path As String) As Boolean
        Dim txt As String
        txt = System.IO.File.ReadAllText(path)
        myCode = txt.Split(vbCrLf)
        Return True
    End Function
    Function run(text As String) As Boolean
        Me.load(text)
        Return Me.run()
    End Function
    Function run() As Boolean
        Dim line As String
        Dim response
        response = 0
        endScript = False
        For Each line In myCode
            If Me.endScript Then
                Me.disconnect()
                Exit For
            Else
                If Not line.TrimStart.StartsWith("'") Then
                    response = vbRetry
                    While response = vbRetry
                        Try
                            Me.Eval(line)
                            response = vbIgnore
                        Catch ex As Exception
                            Dim message As String
                            message = "Command:" & line & vbCrLf & ex.ToString & vbCrLf & vbCrLf
                            message = message & "Abort script, Retry instruction or skip and continue"
                            response = MsgBox(message, MsgBoxStyle.Exclamation + vbAbortRetryIgnore, "Eval Error")
                            If response = vbAbort Then
                                run = False
                                Exit Function
                            End If
                        End Try
                    End While
                End If
            End If
        Next
        Return True
    End Function
    Sub disconnect()
        If Not IsNothing(conn) Then
            conn.CloseConnection()
        End If
    End Sub

    Sub recursivedump(control As Object, file As Integer)
        Dim strCadena As String
        Dim contador_control
        Dim hijos As Object 'SAPFEWSELib.GuiComponentCollection
        Dim control_hijo As Object
        strCadena = ""
        strCadena = control.Id & vbTab
        strCadena = strCadena & "Name: " & control.Name & vbTab
        strCadena = strCadena & "Type: " & control.Type & vbTab
        strCadena = strCadena & "Value: " & control.Text & vbTab
        Try
            If control.Changeable Then
                strCadena = strCadena & "Changeable" & vbTab
            Else
                strCadena = strCadena & "Read only" & vbTab
            End If
        Catch
            strCadena = strCadena & "Read only" & vbTab
        End Try
        Try
            hijos = control.Children
            strCadena = strCadena & "Children:" & hijos.Count
        Catch
            hijos = Nothing
            strCadena = strCadena & "Children: none"
        End Try

        WriteLine(file, strCadena)
        If Not IsNothing(hijos) Then
            For Each control_hijo In control.Children
                Call Me.recursivedump(control_hijo, file)
            Next
            'For contador_control = 0 To control.Children.Count - 1
            '    Try
            '        control_hijo = control.Children.Item(contador_control)
            '    Catch ex As Exception
            '        control_hijo = Nothing
            '    End Try
            '    If Not IsNothing(control_hijo) Then
            '        Call Me.recursivedump(control_hijo, file)
            '    End If
            'Next contador_control
        End If
    End Sub
    Sub dumpscreen(fname As String)
        Dim file_handle As Integer
        file_handle = FreeFile()
        If System.IO.File.Exists(fname) Then
            If System.IO.File.Exists(fname + ".bak") Then
                System.IO.File.Delete(fname + ".bak")
            End If
            System.IO.File.Move(fname, fname + ".bak")
        End If
        FileOpen(file_handle, fname, OpenMode.Output, OpenAccess.Write)
        Me.recursivedump(sess.ActiveWindow, file_handle)
        FileClose(file_handle)
    End Sub
    Sub screenshot(fname As String, format As String)
        Dim id_format As Integer
        Select Case LCase(format)
            Case "bmp"
                id_format = 0
            Case "jpg"
                id_format = 1
            Case "png"
                id_format = 2
            Case "gif"
                id_format = 3
        End Select
        If System.IO.File.Exists(fname) Then
            If System.IO.File.Exists(fname + ".bak") Then
                System.IO.File.Delete(fname + ".bak")
            End If
            System.IO.File.Move(fname, fname + ".bak")
        End If
        sess.ActiveWindow.Restore()
        sess.ActiveWindow.SetFocus()
        sess.ActiveWindow.HardCopy(fname, id_format)
    End Sub
    Sub send(cmd As String)
        Me.sess.SendCommand(cmd)
    End Sub
    Sub setfield(fldName As String, fldType As String, fldValue As String)
        Dim srchType As String
        srchType = ""
        Dim fieldnotfound As Boolean
        fieldnotfound = False
        Select Case LCase(fldType)
            Case "text", "textbox"
                If Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Count = 1 Then
                    srchType = "GuiCTextField"
                ElseIf Me.sess.ActiveWindow.FindAllByName(fldName, "GuiTextField").Count = 1 Then
                    srchType = "GuiTextField"
                ElseIf Me.sess.ActiveWindow.FindAllByName(fldName, "GuiPasswordField").Count = 1 Then
                    srchType = "GuiPasswordField"
                End If
                If srchType.Length > 1 Then
                    Me.sess.ActiveWindow.FindAllByName(fldName, srchType).Item(0).Text = fldValue
                Else : fieldnotfound = True
                End If
            Case "check", "checkbox"
                If Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCheckBox").Count = 1 Then
                    srchType = "GuiCheckBox"
                End If
                If srchType.Length > 1 Then
                    Me.sess.ActiveWindow.FindAllByName(fldName, srchType).Item(0).Selected = LCase(fldValue).Equals("true")
                Else : fieldnotfound = True
                End If
            Case "table", "column", "list"
                Dim fldValues As String()
                Dim singleVal As String
                Dim counter As Integer
                fldValues = fldValue.Split("|".ToCharArray, System.StringSplitOptions.RemoveEmptyEntries)
                counter = 0
                For Each singleVal In fldValues
                    'do something
                    Select Case counter 'en dos palabras IM - PRESIONANTE !!!! 
                        Case 0 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(0).Text = singleVal
                        Case 1 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(1).Text = singleVal
                        Case 2 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(2).Text = singleVal
                        Case 3 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(3).Text = singleVal
                        Case 4 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(4).Text = singleVal
                        Case 5 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(5).Text = singleVal
                        Case 6 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(6).Text = singleVal
                        Case 7 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(7).Text = singleVal
                        Case 8 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(8).Text = singleVal
                        Case 9 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(9).Text = singleVal
                        Case 0 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(0).Text = singleVal
                        Case 11 : Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Item(11).Text = singleVal
                    End Select
                    counter = counter + 1
                Next
            Case Else
                fieldnotfound = True
        End Select
        If fieldnotfound Then
            Throw New System.Exception("Field not found")
        End If
    End Sub
    Function getfield(fldName As String, fldType As String)
        Dim srchType As String, retval As String
        srchType = ""
        retval = "Not found"
        Dim fieldnotfound As Boolean
        fieldnotfound = False
        Select Case LCase(fldType)
            Case "text", "textbox"
                If Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField").Count = 1 Then
                    srchType = "GuiCTextField"
                ElseIf Me.sess.ActiveWindow.FindAllByName(fldName, "GuiTextField").Count = 1 Then
                    srchType = "GuiTextField"
                ElseIf Me.sess.ActiveWindow.FindAllByName(fldName, "GuiPasswordField").Count = 1 Then
                    srchType = "GuiPasswordField"
                End If
                If srchType.Length > 1 Then
                    retval = Me.sess.ActiveWindow.FindAllByName(fldName, srchType).Item(0).Text
                Else : fieldnotfound = True
                End If
            Case "check", "checkbox"
                If Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCheckBox").Count = 1 Then
                    srchType = "GuiCheckBox"
                End If
                If srchType.Length > 1 Then
                    retval = Me.sess.ActiveWindow.FindAllByName(fldName, srchType).Item(0).Selected
                Else : fieldnotfound = True
                End If
            Case "table", "column", "list"
                Dim singleVal As String
                Dim counter As Integer
                Dim scr_value
                counter = 0
                singleVal = ""
                For Each scr_value In Me.sess.ActiveWindow.FindAllByName(fldName, "GuiCTextField")
                    'do something
                    singleVal = scr_value.Text
                Next
            Case "label", "any"
                retval = Me.sess.ActiveWindow.FindAllByName(fldName, 0).Item(0).Selected
            Case Else
                fieldnotfound = True
        End Select
        If fieldnotfound Then
            Throw New System.Exception("Field not found")
        End If
        getfield = retval
    End Function
    Function Eval(cmd As String) As Boolean
        Dim command As String, tmp As String
        Dim arguments As String()
        Dim retval As Boolean = True
        command = LCase((cmd.Split("("))(0))
        If cmd.Split("(")(1).Length > 1 Then
            tmp = Mid(cmd, Len(command) + 3, Len(cmd) - Len(command) - 4)
            arguments = Text.RegularExpressions.Regex.Split(tmp, """,""")
        Else
            arguments = Nothing
        End If
        Select Case command
            Case "end", "exit"
                Me.endScript = True
            Case "attach"
                Me.objWrapper = CreateObject("SapROTWr.SapROTWrapper")
                Me.objRotSAPGUI = objWrapper.GetROTEntry("SAPGUI")
                Me.app = objRotSAPGUI.GetScriptingEngine
                If IsNothing(app) Then
                    MsgBox("Error no running sapguit to attach", MsgBoxStyle.Critical)
                Else
                    Me.conn = app.Children(0)
                    Me.sess = conn.Children(0)
                End If
            Case "logon"
                If arguments.Length = 5 Then
                    Me.logon(arguments(0), arguments(1), arguments(2), arguments(3), arguments(4))
                End If
                If arguments.Length = 6 Then
                    Me.logon(arguments(0), arguments(1), arguments(2), arguments(3), arguments(4), arguments(5))
                End If
            Case "dumpscreen"
                Me.dumpscreen(arguments(0))
            Case "screenshot"
                Me.screenshot(arguments(0), arguments(1))
            Case "setfield"
                Me.setfield(arguments(0), arguments(1), arguments(2))
            Case "getstatus"
                MsgBox("Status: " & Me.getStatus(), MsgBoxStyle.Information + MsgBoxStyle.MsgBoxSetForeground, "LBScript message")
            Case "sendkey"
                Me.sess.ActiveWindow.SendVKey(arguments(0))
            Case "send", "okcode", "sendokcode"
                Me.send(arguments(0))
            Case "transaction"
                Me.sess.EndTransaction()
                Me.sess.StartTransaction(arguments(0))
            Case "msgbox", "wait"
                MsgBox(arguments(0), MsgBoxStyle.Information + MsgBoxStyle.MsgBoxSetForeground, "LBScript message")
            Case "disconnect"
                Me.sess.EndTransaction()
                Me.disconnect()
            Case "shell"
                Dim argument As String, shell_cmd As String
                shell_cmd = ""
                For Each argument In arguments
                    shell_cmd = shell_cmd & argument & " "
                Next
                Shell(shell_cmd, AppWinStyle.NormalFocus, False)
            Case "waitfor"
                Dim argument As String, shell_cmd As String
                shell_cmd = ""
                For Each argument In arguments
                    shell_cmd = shell_cmd & argument & " "
                Next
                Shell(shell_cmd, AppWinStyle.NormalFocus, True)
            Case "hide"
                'MsgBox("Hide app")
                'Me.sess.ActiveWindow.Visualize(False)
            Case "show"
                'MsgBox("Show app")
                'Me.sess.ActiveWindow.Visualize(True)
            Case "getfield"
                MsgBox("Field " & arguments(0) & " = " & Me.getfield(arguments(0), arguments(1)))
            Case "loadscript"
                'This does not always work, only if connecting to an existing sapgui session via a connect() function, 
                'or if saplogon window is open in the system before starting the script
                If Not IsNothing(Me.objRotSAPGUI) Then
                    System.Diagnostics.Process.Start("cscript", arguments(0))
                Else
                    MsgBox("Error, Saplogon window not started before the script, ROT sapgui empty", vbCritical)
                End If
            Case Else
                'ignore it ? raise error ? allow other instructions ?
                retval = False
        End Select
        Eval = retval
    End Function
End Class
