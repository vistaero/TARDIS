Imports NAudio.Wave
Imports System.Threading
Imports System.Globalization.CultureInfo
Imports Microsoft
Imports Microsoft.Win32
Imports Microsoft.Win32.Registry
Imports System.IO

Public Class Form1

    Public Class LoopStream
        Inherits WaveStream
        Private sourceStream As WaveStream

        Public Sub New(sourceStream As WaveStream)
            Me.sourceStream = sourceStream
            Me.EnableLooping = True
        End Sub

        Public Property EnableLooping() As Boolean
            Get
                Return m_EnableLooping
            End Get
            Set(value As Boolean)
                m_EnableLooping = value
            End Set
        End Property

        Private m_EnableLooping As Boolean

        Public Overrides ReadOnly Property WaveFormat() As WaveFormat
            Get
                Return sourceStream.WaveFormat
            End Get
        End Property

        Public Overrides ReadOnly Property Length() As Long
            Get
                Return sourceStream.Length
            End Get
        End Property

        Public Overrides Property Position() As Long
            Get
                Return sourceStream.Position
            End Get
            Set(value As Long)
                sourceStream.Position = value
            End Set
        End Property

        Public Overrides Function Read(buffer As Byte(), offset As Integer, count As Integer) As Integer
            Dim totalBytesRead As Integer = 0

            While totalBytesRead < count
                Dim bytesRead As Integer = sourceStream.Read(buffer, offset + totalBytesRead, count - totalBytesRead)
                If bytesRead = 0 Then
                    If sourceStream.Position = 0 OrElse Not EnableLooping Then
                        ' something wrong with the source stream
                        Exit While
                    End If
                    ' loop
                    sourceStream.Position = 0
                End If
                totalBytesRead += bytesRead
            End While
            Return totalBytesRead
        End Function
    End Class
    Dim Noise As WaveOut
    Dim NoiseReader As AudioFileReader
    Dim EmergencyFlight As WaveOut
    Dim EmergencyFlightReader As AudioFileReader
    Dim Hum2005 As WaveOut
    Dim Hum2005Reader As AudioFileReader
    Dim Hum2010 As WaveOut
    Dim Hum2010Reader As AudioFileReader
    Dim Hum2013 As WaveOut
    Dim Hum2013Reader As AudioFileReader
    Dim Drum As WaveOut
    Dim DrumReader As AudioFileReader
    Dim EndDrum As WaveOut
    Dim EndDrumReader As AudioFileReader
    Dim CloisterBell As WaveOut
    Dim CloisterBellReader As AudioFileReader
    Dim TimeVortex As WaveOut
    Dim TimeVortexReader As AudioFileReader
    Dim DoorOpen As WaveOut
    Dim DoorOpenReader As AudioFileReader
    Dim DoorClose As WaveOut
    Dim DoorCloseReader As AudioFileReader
    Private CBPlaying As Boolean
    Private TVPlaying As Boolean
    Private Travelling As Boolean
    Private SpaceEnabled As Boolean
    Private Language As String
    Private Closemsg As String
    Private WordShow As String
    Private WordHide As String
    Private WordClose As String
    Private Areyousure As String
    Private DoorPlaying As Boolean
    Private EmergencyFlightPattern As Boolean
    Private DoorState As Boolean
    Sub Initapp()

        ThreadPool.QueueUserWorkItem(Sub(o)
                                         Dim SWFfile As String = Application.StartupPath & "\media\2005\Monitor.swf"
                                         videoP.LoadMovie(0, SWFfile)
                                         videoP.Play()
                                         videoP.Loop = True
                                     End Sub)
        SpaceEnabled = False
        Travelling = False
        Label1.Parent = PictureBox1
        
        ' //////////////////////////
        ' IF IT IS THE FIRST RUN
        ' \\\\\\\\\\\\\\\\\\\\\\\\\\
        If My.Settings.IsFirstTime = True Then
            My.Settings.Reset()
            My.Settings.IsMouseVisible = True
            tabControl1.Visible = True
            CheckBox1.Checked = False
            ' Setting look
            Play2005()
            ' Setting the keys
            My.Settings.Escapekey = Keys.Escape
            My.Settings.Helpkey = Keys.F2
            My.Settings.Startkey = Keys.Enter
            My.Settings.Endkey = Keys.Space
            My.Settings.TVKey = Keys.T
            My.Settings.CBKey = Keys.C
            My.Settings.CloseDoorKey = Keys.W
            My.Settings.OpenDoorKey = Keys.Q
            My.Settings.EmergencyFlightKey = Keys.E
            My.Settings.MouseKey = Keys.M
            My.Settings.HideKey = Keys.H
            My.Settings.T2005Key = Keys.D1
            My.Settings.T2010Key = Keys.D2
            My.Settings.T2013key = Keys.D3
            My.Settings.Fullscreenkey = Keys.F11
            My.Settings.Fullscreen = False
            My.Settings.IsFirstTime = False
            My.Settings.Save()
        Else
            ' //////////////////////////
            ' IF IT IS NOT THE FIRST RUN
            ' \\\\\\\\\\\\\\\\\\\\\\\\\\

            My.Settings.Reload()
            If My.Settings.IsMouseVisible = False Then
                Cursor.Hide()
            End If
            If My.Settings.RunAtStart = True Then
                CheckBox1.Checked = True
            ElseIf My.Settings.RunAtStart = False Then
                CheckBox1.Checked = False
            End If
            If My.Settings.Fullscreen = True Then
                Me.TopMost = True
                Me.WindowState = FormWindowState.Maximized
                My.Settings.Fullscreen = True
            Else
                Me.WindowState = FormWindowState.Normal
                Me.TopMost = False
                My.Settings.Fullscreen = False
            End If
            If My.Settings.ActualHum = "Hum2005" Then
                Play2005()
            End If
            If My.Settings.ActualHum = "Hum2010" Then
                Play2010()
            End If
            If My.Settings.ActualHum = "Hum2013" Then
                Play2013()
            End If
        End If
        Button2005.Text = My.Settings.T2005Key
        Button2010.Text = My.Settings.T2010Key
        Button2013.Text = My.Settings.T2013key
        ButtonCB.Text = My.Settings.CBKey
        ButtonTV.Text = My.Settings.TVKey
        ButtonEnd.Text = My.Settings.Endkey
        ButtonFullscreen.Text = My.Settings.Fullscreenkey
        ButtonMenu.Text = My.Settings.Helpkey
        ButtonStart.Text = My.Settings.Startkey
        ButtonHide.Text = My.Settings.HideKey
        ButtonHideMouse.Text = My.Settings.MouseKey
        ButtonEscape.Text = My.Settings.Escapekey
        ' List language files
        Dim fileNames() As String
        Dim result As String
        fileNames = System.IO.Directory.GetFiles(Application.StartupPath & "\languages\", "*.rtf")
        For Each file In fileNames
            result = Path.GetFileNameWithoutExtension(file)
            Me.LanguageComboBox.Items.Add(result)
        Next
        ' End listing language files
        If My.Settings.LanguageAutoDetect = True Then
            Language = System.Globalization.CultureInfo.CurrentCulture.ToString
            If Language.StartsWith("es") Then
                RichTextBox1.LoadFile(Application.StartupPath & "\languages\spanish.rtf")
                TabHelp.Text = "Ayuda"
                TabOther.Text = "Más"
                Label1.Text = "No hay vídeo. ¿Tienes uno?"
                Button11.Text = "GRAN BOTÓN AMISTOSO"
                Closemsg = "Todos los cambios han sido restaurados. El programa se cerrará. A continuación podrá abrirlo de nuevo."
                CheckBox1.Text = "Iniciar con Windows."
                WordShow = "Mostrar"
                WordHide = "Ocultar"
                WordClose = "Cerrar"
                TabSettings.Text = "Personalizar controles y ajustes de audio"
                LanguageComboBox.SelectedIndex = 1
                Areyousure = "¿Estás seguro?"
                GroupBox2.Text = "Restablecer ajustes de la aplicación"
            End If
        Else
            ' Here must be the language-selection system, for now it does not work and english is selected.
            RichTextBox1.LoadFile(Application.StartupPath & "\languages\english.rtf")
            Label1.Text = "Missing video. Do you have one?"
            Closemsg = "Al changes have been restored. The application will end. You can open it again."
            WordShow = "Show"
            WordHide = "Hide"
            WordClose = "Close"
            LanguageComboBox.SelectedIndex = 0
            Areyousure = "Are you sure?"
            ToolStripMenuItem1.Text = WordHide
            CloseToolStripMenuItem.Text = WordClose
        End If
        
    End Sub

    Sub HideUI()
        If ToolStripMenuItem1.Text = WordHide Then
            Me.Visible = False
            ToolStripMenuItem1.Text = WordShow
            My.Settings.IsUiVisible = False
        Else
            Me.Visible = True
            Me.Focus()
            tabControl1.Visible = False
            ToolStripMenuItem1.Text = WordHide
            My.Settings.IsUiVisible = True
        End If
    End Sub

    Sub CloseApp()
        My.Settings.Save()
        NotifyIcon1.Visible = False
        Application.Exit()
    End Sub

    Sub Bigfriendlyexit()
        Dim respuesta As DialogResult = MessageBox.Show(Areyousure, Areyousure, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)
        Select Case respuesta
            Case DialogResult.Yes
                Cursor.Show()
                My.Settings.IsFirstTime = True
                start_Up(False)
                MsgBox(Closemsg)
                CloseApp()
            Case DialogResult.No
        End Select

    End Sub


    Sub HelpWindow()
        If tabControl1.Visible = True Then
            tabControl1.Visible = False
        Else
            tabControl1.Visible = True
        End If
    End Sub

    Sub StopAnyHum()
        Try
            Hum2005.Stop()
            Hum2005.Dispose()
            Hum2010.Stop()
            Hum2010.Dispose()
            Hum2013.Stop()
            Hum2013.Dispose()
        Catch ex As Exception

        End Try
    End Sub

    Sub Play2005()
        StopAnyHum()
        Hum2005Reader = New AudioFileReader(Application.StartupPath & "\media\2005\Hum.wav")
        Dim looping As New LoopStream(Hum2005Reader)        '  
        Hum2005 = New WaveOut()
        Hum2005.Init(looping)
        Hum2005Reader.Volume = Val(T2005Volume.Value) / 10
        Hum2005.Play()
        PictureBox1.Visible = False
        My.Settings.ActualHum = "Hum2005"
    End Sub

    Sub Play2010()
        StopAnyHum()
        Hum2010Reader = New AudioFileReader(Application.StartupPath & "\media\2010\Hum.wav")
        Dim looping As New LoopStream(Hum2010Reader)        '  
        Hum2010 = New WaveOut()
        Hum2010.Init(looping)
        Hum2010Reader.Volume = Val(T2010Volume.Value) / 10
        Hum2010.Play()
        PictureBox1.Visible = True
        PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2010\Monitor.jpg")
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        My.Settings.ActualHum = "Hum2010"
    End Sub

    Sub Play2013()
        StopAnyHum()
        Hum2013Reader = New AudioFileReader(Application.StartupPath & "\media\2013\Hum.wav")
        Dim looping As New LoopStream(Hum2013Reader)        '  
        Hum2013 = New WaveOut()
        Hum2013.Init(looping)
        Hum2013Reader.Volume = Val(T2013Volume.Value) / 10
        Hum2013.Play()
        PictureBox1.Visible = True
        PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2013\Monitor.jpg")
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        My.Settings.ActualHum = "Hum2013"
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        'BIG FRIENDLY KEY
        If e.KeyCode = Keys.F7 Then
            Bigfriendlyexit()
        End If

        ' Hide Key
        If e.KeyCode = My.Settings.HideKey Then
            HideUI()
        End If

        ' Mouse Hide Key
        If e.KeyCode = My.Settings.MouseKey Then
            If My.Settings.IsMouseVisible = True Then
                Cursor.Hide()
                My.Settings.IsMouseVisible = False
            ElseIf My.Settings.IsMouseVisible = False Then
                Cursor.Show()
                My.Settings.IsMouseVisible = True
            End If
        End If

        ' Help key
        If e.KeyCode = My.Settings.Helpkey Then
            HelpWindow()
        End If

        ' Fullscreen key
        If e.KeyCode = My.Settings.Fullscreenkey Then
            If My.Settings.Fullscreen = True Then
                Me.TopMost = False
                Me.WindowState = FormWindowState.Normal
                My.Settings.Fullscreen = False
            Else
                Me.TopMost = True
                Me.WindowState = FormWindowState.Maximized
                My.Settings.Fullscreen = True
            End If
        End If

        ' Exit key
        If e.KeyCode = My.Settings.Escapekey Then
            CloseApp()
        End If

        ' Look and hum
        ' 2005 TARDIS
        If e.KeyCode = My.Settings.T2005Key Then

            If My.Settings.ActualHum = "Hum2005" Then
            Else
                Play2005()
                My.Settings.ActualHum = "Hum2005"
            End If
        End If

        ' 2010 TARDIS
        If e.KeyCode = My.Settings.T2010Key Then
            If My.Settings.ActualHum = "Hum2010" Then
            Else
                Play2010()
                My.Settings.ActualHum = "Hum2010"
            End If
        End If

        '2012 TARDIS (To-Do)
        If e.KeyCode = My.Settings.T2013key Then
            If My.Settings.ActualHum = "Hum2013" Then
            Else
                Play2013()
                My.Settings.ActualHum = "Hum2013"
            End If
        End If

        ' Start travel
        If e.KeyCode = My.Settings.Startkey Then
            If Travelling = False Then
                Travelling = True
                DrumReader = New AudioFileReader(Application.StartupPath & "\media\2005\Drum.wav")
                Dim looping As New LoopStream(DrumReader)        '  
                Drum = New WaveOut()
                Drum.Init(DrumReader)
                DrumReader.Volume = Val(My.Settings.StartVolume) / 10
                Drum.Play()
                DelayAndNoise.Enabled = True
            End If
        End If

        ' End travel
        If e.KeyCode = My.Settings.Endkey Then
            If SpaceEnabled = True Then
                Travelling = False
                SpaceEnabled = False
                Try
                    Noise.Stop()
                    Noise.Dispose()
                    EmergencyFlight.Stop()
                    EmergencyFlight.Dispose()
                Catch ex As Exception

                End Try
                
                EndDrumReader = New AudioFileReader(Application.StartupPath & "\media\2010\EndDrum.wav")
                Dim looping As New LoopStream(EndDrumReader)        '  
                EndDrum = New WaveOut()
                EndDrum.Init(EndDrumReader)
                EndDrumReader.Volume = Val(My.Settings.EndTravelVolume) / 10
                EndDrum.Play()
                EmergencyFlightPattern = False
            End If
        End If

        ' Time Vortex
        If e.KeyCode = My.Settings.TVKey Then
            If TVPlaying = False Then
                TimeVortexReader = New AudioFileReader(Application.StartupPath & "\media\TimeVortex.wav")
                Dim looping As New LoopStream(TimeVortexReader)        '  
                TimeVortex = New WaveOut()
                TimeVortex.Init(looping)
                TimeVortexReader.Volume = Val(My.Settings.TVVolume) / 10
                TimeVortex.Play()
                TVPlaying = True
            Else
                TimeVortex.Stop()
                TimeVortex.Dispose()
                TVPlaying = False
            End If
        End If

        ' Emergency Flight SFX
        If e.KeyCode = My.Settings.EmergencyFlightKey Then
            EmergencyFlightPattern = True
        End If

        ' Cloister Bell
        If e.KeyCode = My.Settings.CBKey Then
            If CBPlaying = False Then
                CloisterBellReader = New AudioFileReader(Application.StartupPath & "\media\CloisterBell.wav")
                Dim looping As New LoopStream(CloisterBellReader)        '  
                CloisterBell = New WaveOut()
                CloisterBell.Init(looping)
                CloisterBellReader.Volume = Val(My.Settings.CBVolume) / 10
                CloisterBell.Play()
                CBPlaying = True
            Else
                CloisterBell.Stop()
                CloisterBell.Dispose()
                CBPlaying = False
            End If
        End If
        ' Open the door
        If e.KeyCode = My.Settings.OpenDoorKey Then
            If DoorState = False Then
                If DoorTimer.Enabled = False Then
                    DoorTimer.Enabled = True
                    DoorOpenReader = New AudioFileReader(Application.StartupPath & "\media\2005\DoorOpen.wav")       '  
                    DoorOpen = New WaveOut()
                    DoorOpen.Init(DoorOpenReader)
                    DoorOpenReader.Volume = Val(TravellingVolume.Value) / 10
                    DoorOpen.Play()
                    DoorState = True
                End If
            End If
        End If
        ' Close the door
        If e.KeyCode = My.Settings.CloseDoorKey Then
            If DoorState = True Then
                If DoorTimer.Enabled = False Then
                    DoorTimer.Enabled = True
                    DoorCloseReader = New AudioFileReader(Application.StartupPath & "\media\2005\DoorClose.wav")    '  
                    DoorClose = New WaveOut()
                    DoorClose.Init(DoorCloseReader)
                    DoorCloseReader.Volume = Val(TravellingVolume.Value) / 10
                    DoorClose.Play()
                    DoorState = False
                End If
            End If
        End If

    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initapp()
    End Sub

    Private Sub DelayAndNoise_Tick(sender As Object, e As EventArgs) Handles DelayAndNoise.Tick
        If EmergencyFlightPattern = False Then
            NoiseReader = New AudioFileReader(Application.StartupPath & "\media\2010\Noise.wav")
            Dim looping As New LoopStream(NoiseReader)        '  
            Noise = New WaveOut()
            Noise.Init(looping)
            NoiseReader.Volume = Val(TravellingVolume.Value) / 10
            Noise.Play()
        Else
            EmergencyFlightReader = New AudioFileReader(Application.StartupPath & "\media\EmergencyFlight.wav")
            Dim looping As New LoopStream(EmergencyFlightReader)        '  
            EmergencyFlight = New WaveOut()
            EmergencyFlight.Init(looping)
            EmergencyFlightReader.Volume = Val(TravellingVolume.Value) / 10
            EmergencyFlight.Play()
        End If
        Drum = Nothing
        SpaceEnabled = True
        DelayAndNoise.Enabled = False
    End Sub

    Private Sub RichTextBox1_GotFocus(sender As Object, e As EventArgs) Handles RichTextBox1.GotFocus
        Me.Focus()
    End Sub

    Private Sub RichTextBox1_LinkClicked1(sender As Object, e As LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Bigfriendlyexit()

    End Sub

    Private Function start_Up(ByVal bCreate As Boolean) As String
        Const key As String = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        Dim subClave As String = Application.ProductName.ToString
        Dim msg As String = ""
        Try
            Dim Registro As RegistryKey = CurrentUser.CreateSubKey(key, RegistryKeyPermissionCheck.ReadWriteSubTree)
            With Registro
                .OpenSubKey(key, True)
                Select Case bCreate
                    Case True
                        .SetValue(subClave, _
                                  Application.ExecutablePath.ToString)
                    Case False
                        If .GetValue(subClave, "").ToString <> "" Then
                            .DeleteValue(subClave)
                        End If
                End Select
            End With
        Catch ex As Exception
            msg = ex.Message.ToString
        End Try
        Return Nothing
    End Function


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            ' Código para añadir a inicio
            start_Up(True)
            My.Settings.RunAtStart = True
            My.Settings.Save()
        Else
            ' Código para quitar de inicio
            start_Up(False)
            My.Settings.RunAtStart = False
            My.Settings.Save()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        HideUI()

    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        My.Settings.Save()
        NotifyIcon1.Visible = False
        End
    End Sub

    Private Sub ContextMenuStrip1_Closing(sender As Object, e As ToolStripDropDownClosingEventArgs) Handles ContextMenuStrip1.Closing
        Me.Focus()
        Me.TopMost = False
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If tabControl1.Visible = False Then
            Me.Focus()
            Me.TopMost = True
            Me.TopMost = False
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDown(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDown
        Me.Focus()
        Me.TopMost = True
        Me.TopMost = False
    End Sub

    Private Sub BigFriendlyButtonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BigFriendlyButtonToolStripMenuItem.Click
        Bigfriendlyexit()
    End Sub

    Private Sub DoorTimer_Tick(sender As Object, e As EventArgs) Handles DoorTimer.Tick
        DoorTimer.Enabled = False
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        Label8.Visible = True
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False Then
            LanguageComboBox.Enabled = True
            My.Settings.LanguageAutoDetect = False
        Else
            LanguageComboBox.Enabled = False
            My.Settings.LanguageAutoDetect = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        My.Settings.Save()
        Initapp()
    End Sub

    Private Sub T2005Volume_Scroll(sender As Object, e As EventArgs) Handles T2005Volume.Scroll
        Try
            Hum2005Reader.Volume = Val(T2005Volume.Value) / 10
        Catch ex As Exception
        End Try

    End Sub

    Private Sub T2010Volume_Scroll(sender As Object, e As EventArgs) Handles T2010Volume.Scroll
        Try
            Hum2010Reader.Volume = Val(T2010Volume.Value) / 10
        Catch ex As Exception
        End Try

    End Sub

    Private Sub T2013Volume_Scroll(sender As Object, e As EventArgs) Handles T2013Volume.Scroll
        Try
            Hum2013Reader.Volume = Val(T2013Volume.Value) / 10
        Catch ex As Exception
        End Try

    End Sub

    Private Sub StartVolume_Scroll(sender As Object, e As EventArgs) Handles StartVolume.Scroll
        Try
            DrumReader.Volume = Val(StartVolume.Value) / 10
        Catch ex As Exception
        End Try

    End Sub

    Private Sub TravellingVolume_Scroll(sender As Object, e As EventArgs) Handles TravellingVolume.Scroll
        Try
            NoiseReader.Volume = Val(TravellingVolume.Value) / 10
        Catch ex As Exception
        End Try
    End Sub

    Private Sub EndTravelVolume_Scroll(sender As Object, e As EventArgs) Handles EndTravelVolume.Scroll
        Try
            EndDrumReader.Volume = Val(EndTravelVolume.Value) / 10
        Catch ex As Exception
        End Try
    End Sub

    Private Sub TVVolume_Scroll(sender As Object, e As EventArgs) Handles TVVolume.Scroll
        Try
            TimeVortexReader.Volume = Val(TVVolume.Value) / 10
        Catch ex As Exception
        End Try
    End Sub

    Private Sub CBVolume_Scroll(sender As Object, e As EventArgs) Handles CBVolume.Scroll
        Try
            CloisterBellReader.Volume = Val(CBVolume.Value) / 10
        Catch ex As Exception
        End Try
    End Sub
End Class
