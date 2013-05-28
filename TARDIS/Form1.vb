Imports NAudio.Wave
Imports System.Threading
Imports System.Globalization.CultureInfo
Imports Microsoft
Imports Microsoft.Win32
Imports Microsoft.Win32.Registry

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

    Private Hum As WaveOut
    Private Drum As WaveOut
    Private Noise As WaveOut
    Private EndDrum As WaveOut
    Private CloisterBell As WaveOut
    Private TimeVortex As WaveOut
    Private CBPlaying As Boolean
    Private TVPlaying As Boolean
    Private Travelling As Boolean
    Private SpaceEnabled As Boolean
    Private Language As String
    Private Closemsg As String
    Private WordShow As String
    Private WordHide As String
    Private WordClose As String

    Sub CloseApp()
        My.Settings.Save()
        NotifyIcon1.Visible = False
        Application.Exit()
    End Sub

    Sub HelpWindow()
        If tabControl1.Visible = True Then
            tabControl1.Visible = False
            Cursor.Hide()
        Else
            tabControl1.Visible = True
            Cursor.Show()
        End If
    End Sub

    Sub Play2005()
        Dim reader As New WaveFileReader("media/2005Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
        PictureBox1.Visible = False
        My.Settings.ActualHum = "Hum2005"
    End Sub

    Sub Play2010()
        Dim reader As New WaveFileReader("media/2010Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
        PictureBox1.Visible = True
        PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2010Monitor.jpg")
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        My.Settings.ActualHum = "Hum2010"
    End Sub

    Sub Play2013()
        Dim reader As New WaveFileReader("media/2013Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
        PictureBox1.Visible = True
        PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2013Monitor.jpg")
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        My.Settings.ActualHum = "Hum2013"
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

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
                Hum.Stop()
                Hum.Dispose()
                Play2005()
                My.Settings.ActualHum = "Hum2005"
            End If
        End If

        ' 2010 TARDIS
        If e.KeyCode = My.Settings.T2010Key Then
            If My.Settings.ActualHum = "Hum2010" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                Play2010()
                My.Settings.ActualHum = "Hum2010"
            End If
        End If

        '2012 TARDIS (To-Do)
        If e.KeyCode = My.Settings.T2013key Then
            If My.Settings.ActualHum = "Hum2013" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                Play2013()
                My.Settings.ActualHum = "Hum2013"
            End If
        End If

        ' End look and hum

        ' Start travel
        If e.KeyCode = My.Settings.Startkey Then
            If Travelling = False Then
                Travelling = True
                Dim reader As New WaveFileReader("media/Drum.wav")
                Drum = New WaveOut()
                Drum.Init(reader)
                Drum.Play()
                DelayAndNoise.Enabled = True
            End If
        End If

        ' End travel
        If e.KeyCode = My.Settings.Endkey Then
            If SpaceEnabled = True Then
                Travelling = False
                SpaceEnabled = False
                Noise.Stop()
                Noise.Dispose()
                Dim reader As New WaveFileReader("media/EndDrum.wav")
                EndDrum = New WaveOut()
                EndDrum.Init(reader)
                EndDrum.Play()
            End If
        End If

        ' Time Vortex
        If e.KeyCode = My.Settings.TVKey Then
            If TVPlaying = False Then
                Dim reader As New WaveFileReader("media/TimeVortex.wav")
                Dim looping As New LoopStream(reader)
                TimeVortex = New WaveOut()
                TimeVortex.Init(looping)
                TimeVortex.Play()
                TVPlaying = True
            Else
                TimeVortex.Stop()
                TimeVortex.Dispose()
                TVPlaying = False
            End If
        End If
        ' Cloister Bell
        If e.KeyCode = My.Settings.CBKey Then
            If CBPlaying = False Then
                Dim reader As New WaveFileReader("media/CloisterBell.wav")
                Dim looping As New LoopStream(reader)
                CloisterBell = New WaveOut()
                CloisterBell.Init(looping)
                CloisterBell.Play()
                CBPlaying = True
            Else
                CloisterBell.Stop()
                CloisterBell.Dispose()
                CBPlaying = False
            End If
        End If

    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ThreadPool.QueueUserWorkItem(Sub(o)
                                         Dim SWFfile As String = Application.StartupPath & "\media\2005Monitor.swf"
                                         videoP.LoadMovie(0, SWFfile)
                                         videoP.Play()
                                         videoP.Loop = True
                                     End Sub)
        SpaceEnabled = False
        Travelling = False
        Label1.Parent = PictureBox1
        Language = System.Globalization.CultureInfo.CurrentCulture.ToString
        If Language.StartsWith("es") Then
            RichTextBox1.LoadFile(Application.StartupPath & "\languages\spanish.rtf")
            tabPage1.Text = "Ayuda"
            tabPage2.Text = "Sonido"
            tabPage3.Text = "Controles"
            TabPage4.Text = "Otros"
            SoundTabText.Text = "Dispositivos de salida y controles de volumen (Por hacer)."
            ControlsTextBox.Text = "Personalizar los controles (Por hacer)."
            Label1.Text = "No hay vídeo. ¿Tienes uno?"
            Button11.Text = "GRAN BOTÓN AMISTOSO"
            Closemsg = "Todos los cambios han sido restaurados. El programa se cerrará. A continuación podrá abrirlo de nuevo."
            CheckBox1.Text = "Iniciar con Windows."
            WordShow = "Mostrar"
            WordHide = "Ocultar"
            WordClose = "Cerrar"
        Else
            RichTextBox1.LoadFile(Application.StartupPath & "\languages\english.rtf")
            SoundTabText.Text = "Output devices and volume settings (TO-DO)."
            ControlsTextBox.Text = "Personalize the controls (TO-DO)."
            Label1.Text = "Missing video. Do you have one?"
            Closemsg = "Al changes have been restored. The application will end. You can open it again."
            WordShow = "Show"
            WordHide = "Hide"
            WordClose = "Close"
        End If
        ToolStripMenuItem1.Text = WordHide
        CloseToolStripMenuItem.Text = WordClose
        ' //////////////////////////
        ' IF IT IS THE FIRST RUN
        ' \\\\\\\\\\\\\\\\\\\\\\\\\\
        If My.Settings.IsFirstTime = True Then
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
            My.Settings.T2005Key = Keys.D1
            My.Settings.T2010Key = Keys.D2
            My.Settings.T2013key = Keys.D3
            My.Settings.Fullscreenkey = Keys.F11
            My.Settings.Fullscreen = False
            My.Settings.IsFirstTime = False
            My.Application.SaveMySettingsOnExit = True
            My.Settings.IsFirstTime = False
        Else
            ' //////////////////////////
            ' IF IT IS NOT THE FIRST RUN
            ' \\\\\\\\\\\\\\\\\\\\\\\\\\
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
    End Sub

    Private Sub videoP_GotFocus1(sender As Object, e As EventArgs)
        Me.Focus()
    End Sub

    Private Sub videoP_MouseCaptureChanged(sender As Object, e As EventArgs)
        Me.Focus()
    End Sub

    Private Sub Help_Tick(sender As Object, e As EventArgs)
        HelpWindow()
    End Sub

    Private Sub DelayAndNoise_Tick(sender As Object, e As EventArgs) Handles DelayAndNoise.Tick
        Dim reader As New WaveFileReader("media/Noise.wav")
        Dim looping As New LoopStream(reader)
        Noise = New WaveOut()
        Noise.Init(looping)
        Noise.Play()
        Drum = Nothing
        SpaceEnabled = True
        DelayAndNoise.Enabled = False
    End Sub

    Private Sub RichTextBox1_LinkClicked1(sender As Object, e As LinkClickedEventArgs) Handles RichTextBox1.LinkClicked
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        My.Settings.IsFirstTime = True
        start_Up(False)
        MsgBox(Closemsg)
        CloseApp()
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
        If ToolStripMenuItem1.Text = WordHide Then
            Me.Visible = False
            ToolStripMenuItem1.Text = WordShow
        Else
            Me.Visible = True
            ToolStripMenuItem1.Text = WordHide
        End If

    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        My.Settings.Save()
        NotifyIcon1.Visible = False
        End
    End Sub
End Class
