Imports NMALib

Public Class NMAonNewFileService

    ' NOTE you should install this service as WHS/Andrew  then change to local system account once installed.

    Dim WithEvents n As New Monitor.Notifier
    Dim LastNotification As New Generic.Dictionary(Of String, DateTime)
    Dim settings As New Monitor.SettingsReader

    Private Function GetApplicationPath() As String

        Dim strServicePathLong As String = System.Reflection.Assembly.GetExecutingAssembly.Location.ToString()
        Dim intServicePathLength As Integer = Len(Strings.Right(strServicePathLong, Len(strServicePathLong) - InStrRev(strServicePathLong, "\")))
        Dim strServicePath As String = Strings.Left(strServicePathLong, Len(strServicePathLong) - intServicePathLength)

        Return strServicePath

    End Function


    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.

        'Logging.WriteToEventLog("Starting Service", "New File Notification", EventLogEntryType.Information)

        Try

            settings.ReadSettings(GetApplicationPath)

            For Each mi As Monitor.SettingsReader.MonitorItem In settings.ItemsToMonitor
                n.AddWatcher(mi.Name, mi.Path, mi.Filter, mi.IncludeSubDirectories)
            Next

        Catch ex As Exception
            Logging.WriteToEventLog(ex.Message & ControlChars.NewLine & ex.InnerException.ToString, "New File Notification", EventLogEntryType.Error)
        End Try

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.

        'Logging.WriteToEventLog("Stopping Service", "New File Notification", EventLogEntryType.Information)

        n.ClearWatchers()
    End Sub

    Private Sub n_Created(ByVal name As String, ByVal fileName As String) Handles n.Created
        Try
            Dim okToSend As Boolean = True

            If LastNotification.ContainsKey(name) Then
                If LastNotification(name).AddMinutes(settings.Frequency) > Date.Now Then
                    okToSend = False
                End If
            Else
                LastNotification.Add(name, Date.Now)
            End If

            If okToSend Then

                'Create a client/notification.
                Dim Client As New NMAClient()

                'Create the notification command
                Dim notification As New NMANotification
                notification.Priority = NMANotificationPriority.Normal
                notification.Event = settings.NMAEvent
                notification.Description = settings.NMAMessage.Replace("{1}", name)

                Select Case settings.NMAPriority.ToUpper
                    Case "EMERGENCY", "E"
                        notification.Priority = NMANotificationPriority.Emergency
                    Case "HIGH", "H"
                        notification.Priority = NMANotificationPriority.High
                    Case "MODERATE", "M"
                        notification.Priority = NMANotificationPriority.Moderate
                    Case "NORMAL", "N"
                        notification.Priority = NMANotificationPriority.Normal
                    Case "VERYLOW", "L"
                        notification.Priority = NMANotificationPriority.VeryLow
                    Case Else
                        notification.Priority = NMANotificationPriority.Normal
                End Select

                Client.PostNotification(notification)

                Console.WriteLine(name & ":" & fileName)

                LastNotification(name) = Date.Now
            End If

        Catch ex As Exception
            Logging.WriteToEventLog(ex.Message & ControlChars.NewLine & ex.InnerException.ToString, "New File Notification", EventLogEntryType.Error)
        End Try

    End Sub



End Class
