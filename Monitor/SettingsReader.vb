Public Class SettingsReader

    Public Class MonitorItem
        Public Name As String
        Public Path As String
        Public Filter As String
        Public IncludeSubDirectories As Boolean        
    End Class

    Public ItemsToMonitor As New Generic.List(Of MonitorItem)

    Public Frequency As Integer
    Public NMAEvent As String
    Public NMAMessage As String
    Public NMAPriority As String


    Public Sub ReadSettings(ByVal path As String)

        ItemsToMonitor.Clear()

        Dim x As New Xml.XmlDocument()

        Try
            x.Load(IO.Path.Combine(path, "Settings.xml"))
        Catch ex As Exception
            Throw New Exception("Unable to load settings.xml", ex)
        End Try

        Dim rootSettingNode As Xml.XmlNode
        rootSettingNode = x.SelectSingleNode("/Settings")

        If rootSettingNode IsNot Nothing Then
            NMAEvent = rootSettingNode.Attributes.GetNamedItem("event").Value
            NMAMessage = rootSettingNode.Attributes.GetNamedItem("message").Value
            NMAPriority = rootSettingNode.Attributes.GetNamedItem("priority").Value

            Frequency = CInt(rootSettingNode.Attributes.GetNamedItem("frequency").Value)
        Else
            Throw New Exception("No Settings Node found")
        End If

        'Read items to monitor    
        Dim settingNodes As Xml.XmlNodeList
        settingNodes = x.SelectNodes("/Settings/Setting")

        If Not settingNodes Is Nothing Then
            For Each settingNode As Xml.XmlNode In settingNodes
                Dim mi As New MonitorItem
                mi.Name = settingNode.Attributes.GetNamedItem("name").Value
                mi.Path = settingNode.Attributes.GetNamedItem("path").Value
                mi.Filter = settingNode.Attributes.GetNamedItem("filter").Value
                If settingNode.Attributes.GetNamedItem("includesubdirectories").Value.ToLower = "true" Then
                    mi.IncludeSubDirectories = True
                Else
                    mi.IncludeSubDirectories = False
                End If

                ItemsToMonitor.Add(mi)
            Next
        Else
            Throw New Exception("No Settings/Setting Node found")
        End If

    End Sub




End Class
