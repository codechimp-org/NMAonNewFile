Public Class Notifier

    Private watcherNames As New Generic.Dictionary(Of Integer, String)
    Private watchers As New Generic.List(Of IO.FileSystemWatcher)

    Public Event Created(ByVal name As String, ByVal fileName As String)

    Public Sub AddWatcher(ByVal name As String, ByVal path As String, ByVal filter As String, ByVal includeSubDirectories As Boolean)

        Dim w As New IO.FileSystemWatcher()

        w.Path = path
        w.Filter = filter
        w.IncludeSubdirectories = includeSubDirectories

        watcherNames.Add(w.GetHashCode, name)

        AddHandler w.Created, AddressOf WatcherCreatedHandler

        watchers.Add(w)

        w.EnableRaisingEvents = True

    End Sub

    Public Sub ClearWatchers()
        For Each w As IO.FileSystemWatcher In watchers
            w.EnableRaisingEvents = False
            RemoveHandler w.Created, AddressOf WatcherCreatedHandler
        Next

        watchers.Clear()
        watcherNames.Clear()
    End Sub

    Private Sub WatcherCreatedHandler(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs)
        Dim watcher As IO.FileSystemWatcher

        watcher = DirectCast(sender, IO.FileSystemWatcher)



        RaiseEvent Created(watcherNames.Item(watcher.GetHashCode), e.FullPath)

    End Sub
End Class
