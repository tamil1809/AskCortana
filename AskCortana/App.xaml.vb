Imports Windows.UI
''' <summary>
''' Provides application-specific behavior to supplement the default Application class.
''' </summary>
NotInheritable Class App
    Inherits Application

    ''' <summary>
    ''' Invoked when the application is launched normally by the end user.  Other entry points
    ''' will be used when the application is launched to open a specific file, to display
    ''' search results, and so forth.
    ''' </summary>
    ''' <param name="e">Details about the launch request and process.</param>
    Protected Overrides Sub OnLaunched(e As Windows.ApplicationModel.Activation.LaunchActivatedEventArgs)
        Dim rootFrame As Frame = TryCast(Window.Current.Content, Frame)

        ' Do not repeat app initialization when the Window already has content,
        ' just ensure that the window is active

        If rootFrame Is Nothing Then
            ' Create a Frame to act as the navigation context and navigate to the first page
            rootFrame = New Frame()

            AddHandler rootFrame.NavigationFailed, AddressOf OnNavigationFailed

            If e.PreviousExecutionState = ApplicationExecutionState.Terminated Then
                ' TODO: Load state from previously suspended application
            End If
            ' Place the frame in the current Window
            Window.Current.Content = rootFrame
        End If

        If e.PrelaunchActivated = False Then
            If rootFrame.Content Is Nothing Then
                ' When the navigation stack isn't restored navigate to the first page,
                ' configuring the new page by passing required information as a navigation
                ' parameter
                rootFrame.Navigate(GetType(MainPage), e.Arguments)
            End If

            ' Ensure the current window is active
            Window.Current.Activate()
        End If
    End Sub

    ''' <summary>
    ''' Invoked when Navigation to a certain page fails
    ''' </summary>
    ''' <param name="sender">The Frame which failed navigation</param>
    ''' <param name="e">Details about the navigation failure</param>
    Private Sub OnNavigationFailed(sender As Object, e As NavigationFailedEventArgs)
        Throw New Exception("Failed to load Page " + e.SourcePageType.FullName)
    End Sub

    ''' <summary>
    ''' Invoked when application execution is being suspended.  Application state is saved
    ''' without knowing whether the application will be terminated or resumed with the contents
    ''' of memory still intact.
    ''' </summary>
    ''' <param name="sender">The source of the suspend request.</param>
    ''' <param name="e">Details about the suspend request.</param>
    Private Sub OnSuspending(sender As Object, e As SuspendingEventArgs) Handles Me.Suspending
        Dim deferral As SuspendingDeferral = e.SuspendingOperation.GetDeferral()
        ' TODO: Save application state and stop any background activity
        deferral.Complete()
    End Sub

    ''' OnActivateed method for launch app by voice command using cortana command
    Protected Overrides Sub OnActivated(ByVal e As IActivatedEventArgs)
        If e.Kind <> Windows.ApplicationModel.Activation.ActivationKind.VoiceCommand Then
            Return
        End If

        Dim commandArgs = TryCast(e, Windows.ApplicationModel.Activation.VoiceCommandActivatedEventArgs)
        Dim speechRecognitionResult = commandArgs.Result
        Dim voiceCommandName As String = speechRecognitionResult.RulePath(0)
        Dim textSpoken As String = speechRecognitionResult.Text
        Dim spokenColor As String = ""
        Try
            spokenColor = speechRecognitionResult.SemanticInterpretation.Properties("color")(0)
        Catch
        End Try

        Dim color As Windows.UI.Color
        Select Case spokenColor
            Case "Red"
                color = Colors.Red
            Case "Blue"
                color = Colors.Blue
            Case "Yellow"
                color = Colors.Yellow
            Case "Green"
                color = Colors.Green
            Case Else
                color = Colors.Purple
        End Select

        Dim rootFrame As Frame = TryCast(Window.Current.Content, Frame)
        Dim page As MainPage = TryCast(rootFrame.Content, MainPage)
        If page Is Nothing Then
            Return
        End If

        Select Case voiceCommandName
            Case "addRectangle"
                page.CreateRectangle(color)
            Case "addCircle"
                page.CreateCircle(color)
            Case Else
        End Select
    End Sub

End Class
