' The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

Imports Windows.UI
''' <summary>
''' An empty page that can be used on its own or navigated to within a Frame.
''' </summary>
Public NotInheritable Class MainPage
    Inherits Page
    ''' Void Page_Load Method
    Private Async Sub Page_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Dim storageFile = Await Windows.Storage.StorageFile.GetFileFromApplicationUriAsync(New Uri("ms-appx:///CortanaExampleCommands.xml"))
        Await Windows.ApplicationModel.VoiceCommands.VoiceCommandDefinitionManager.InstallCommandDefinitionsFromStorageFileAsync(storageFile)
    End Sub


    Public Sub CreateRectangle(ByVal color As Color)
        Dim random As Random = New Random()
        Dim left = random.[Next](0, 300)
        Dim top = random.[Next](0, 300)
        Dim rect = New Windows.UI.Xaml.Shapes.Rectangle()
        rect.Height = 100
        rect.Width = 100
        rect.Margin = New Thickness(left, top, 0, 0)
        rect.Fill = New SolidColorBrush(color)
        LayoutGrid.Children.Add(rect)

    End Sub

    Public Sub CreateCircle(ByVal color As Color)
        Dim random As Random = New Random()
        Dim left = random.[Next](300)
        Dim top = random.[Next](300)
        Dim circle = New Windows.UI.Xaml.Shapes.Ellipse()
        circle.Height = 100
        circle.Width = 100
        circle.Margin = New Thickness(left, top, 0, 0)
        circle.Fill = New SolidColorBrush(color)
        LayoutGrid.Children.Add(circle)
    End Sub
End Class
