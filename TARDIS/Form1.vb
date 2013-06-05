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
    Private Door As WaveOut
    Private EmergencyFlight As WaveOut
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
    Private EmergencyFlightPlaying As Boolean
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
        Else
            RichTextBox1.LoadFile(Application.StartupPath & "\languages\english.rtf")
            Label1.Text = "Missing video. Do you have one?"
            Closemsg = "Al changes have been restored. The application will end. You can open it again."
            WordShow = "Show"
            WordHide = "Hide"
            WordClose = "Close"
            LanguageComboBox.SelectedIndex = 0
            Areyousure = "Are you sure?"
        End If
        ToolStripMenuItem1.Text = WordHide
        CloseToolStripMenuItem.Text = WordClose
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
            Hum.Volume = "0,5"
            ' Setting the keys
            My.Settings.Escapekey = Keys.Escape
            My.Settings.Helpkey = Keys.F2
            My.Settings.Startkey = Keys.Enter
            My.Settings.Endkey = Keys.Space
            My.Settings.TVKey = Keys.T
            My.Settings.CBKey = Keys.C
            My.Settings.MouseKey = Keys.M
            My.Settings.HideKey = Keys.H
            My.Settings.T2005Key = Keys.D1
            My.Settings.T2010Key = Keys.D2
            My.Settings.T2013key = Keys.D3
            My.Settings.Fullscreenkey = Keys.F11
            My.Settings.T2005Volume = 5
            My.Settings.T2010Volume = 5
            My.Settings.T2013Volume = 5
            My.Settings.StartVolume = 5
            My.Settings.TravellingVolume = 5
            My.Settings.EndTravelVolume = 5
            My.Settings.TVVolume = 5
            My.Settings.CBVolume = 5
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
    End Sub

    Sub HideUI()
        If ToolStripMenuItem1.Text = WordHide Then
            Me.Visible = False
            ToolStripMenuItem1.Text = WordShow
            My.Settings.IsUiVisible = False
        Else
            Me.Visible = True
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

    Sub Play2005()
        Dim reader As New WaveFileReader(Application.StartupPath & "\media\2005\Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Volume = My.Settings.T2005Volume
        Hum.Play()
        PictureBox1.Visible = False
        My.Settings.ActualHum = "Hum2005"
    End Sub

    Sub Play2010()
        Dim reader As New WaveFileReader(Application.StartupPath & "\media\2010\Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
        PictureBox1.Visible = True
        PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2010\Monitor.jpg")
        PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
        My.Settings.ActualHum = "Hum2010"
    End Sub

    Sub Play2013()
        Dim reader As New WaveFileReader(Application.StartupPath & "\media\2013\Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
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

        ' Start travel
        If e.KeyCode = My.Settings.Startkey Then
            If Travelling = False Then
                Travelling = True
                Dim inputStream As WaveChannel32
                Dim drumreader As New WaveFileReader(Application.StartupPath & "\media\2005\Drum.wav")
                inputStream = New WaveChannel32(drumreader)
                Drum = New WaveOut()
                Drum.Init(drumreader)
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
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\2010\EndDrum.wav")
                EndDrum = New WaveOut()
                EndDrum.Init(reader)
                EndDrum.Play()
            End If
        End If

        ' Time Vortex
        If e.KeyCode = My.Settings.TVKey Then
            If TVPlaying = False Then
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\TimeVortex.wav")
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

        ' Emergency Flight SFX
        If e.KeyCode = Keys.E Then
            If EmergencyFlightPlaying = False Then
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\EmergencyFlight.wav")
                Dim looping As New LoopStream(reader)
                EmergencyFlight = New WaveOut()
                EmergencyFlight.Init(looping)
                EmergencyFlight.Play()
                EmergencyFlightPlaying = True
            Else
                EmergencyFlight.Stop()
                EmergencyFlight.Dispose()
                EmergencyFlightPlaying = False
            End If
        End If

        ' Cloister Bell
        If e.KeyCode = My.Settings.CBKey Then
            If CBPlaying = False Then
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\CloisterBell.wav")
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
        ' Open the door
        If e.KeyCode = Keys.Q Then
            If DoorTimer.Enabled = False Then
                DoorTimer.Enabled = True
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\2005\DoorOpen.wav")
                Dim looping As New LoopStream(reader)
                Door = New WaveOut()
                Door.Init(reader)
                Door.Play()
            End If
            

        End If
        ' Close the door
        If e.KeyCode = Keys.W Then
            If DoorTimer.Enabled = False Then
                DoorTimer.Enabled = True
                Dim reader As New WaveFileReader(Application.StartupPath & "\media\2005\DoorClose.wav")
                Dim looping As New LoopStream(reader)
                Door = New WaveOut()
                Door.Init(reader)
                Door.Play()
            End If
        End If

    End Sub

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Initapp()
    End Sub

    Private Sub DelayAndNoise_Tick(sender As Object, e As EventArgs) Handles DelayAndNoise.Tick
        Dim reader As New WaveFileReader(Application.StartupPath & "\media\2010\Noise.wav")
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

    Private Sub T2005Volume_Scroll(sender As Object, e As EventArgs) Handles T2005Volume.MouseUp
        Hum.Volume = Val(T2005Volume.Value) / 10
        T2010Volume.Value = T2005Volume.Value
        T2013Volume.Value = T2005Volume.Value
    End Sub

    Private Sub T2010Volume_Scroll(sender As Object, e As EventArgs) Handles T2010Volume.MouseUp
        Hum.Volume = Val(T2010Volume.Value) / 10
        T2005Volume.Value = T2010Volume.Value
        T2013Volume.Value = T2010Volume.Value
    End Sub

    Private Sub T2013Volume_Scroll(sender As Object, e As EventArgs) Handles T2013Volume.MouseUp
        Hum.Volume = Val(T2013Volume.Value) / 10
        T2005Volume.Value = T2013Volume.Value
        T2010Volume.Value = T2013Volume.Value
    End Sub

    Private Sub StartVolume_Scroll(sender As Object, e As EventArgs) Handles StartVolume.Scroll
        My.Settings.StartVolume = Val(StartVolume.Value) / 10
    End Sub

End Class
