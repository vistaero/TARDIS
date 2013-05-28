Imports NAudio.Wave
Imports System.Threading
Imports System.Globalization.CultureInfo

Public Class Form1
    Inherits System.Windows.Forms.Form
    Dim SWFfile As String = Application.StartupPath & "\media\2005Monitor.swf"

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
    Private ActualHum As String
    Private Travelling As Boolean
    Private SpaceEnabled As Boolean
    ' Valores que deben ser guardados en un archivo
    Private Language As String
    Private Fullscreenkey As String
    Private Escapekey As String
    Private Helpkey As String
    Private Startkey As String
    Private Endkey As String
    Private TVKey As String
    Private T2005Key As String
    Private T2010Key As String
    Private T2013key As String


    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        ' Help key
        If e.KeyCode = Helpkey Then
            Help.Enabled = False
            If tabControl1.Visible = True Then
                tabControl1.Visible = False
                Cursor.Hide()
            ElseIf tabControl1.Visible = False Then
                tabControl1.Visible = True
                Cursor.Show()
            End If
        End If

        ' Fullscreen key
        If e.KeyCode = Fullscreenkey Then
            If Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
            ElseIf Me.WindowState = FormWindowState.Normal Then
                Me.WindowState = FormWindowState.Maximized

            End If
        End If

        ' Exit key
        If e.KeyCode = Escapekey Then
            End
        End If

        ' Look and hum
        ' 2005 TARDIS
        If e.KeyCode = T2005Key Then
            If ActualHum = "Hum2005" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                Dim reader As New WaveFileReader("media/2005Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                PictureBox1.Visible = False
                ActualHum = "Hum2005"
            End If
        End If

        ' 2010 TARDIS
        If e.KeyCode = T2010Key Then
            If ActualHum = "Hum2010" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                Dim reader As New WaveFileReader("media/2010Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                PictureBox1.Visible = True
                PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2010Monitor.jpg")
                PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
                ActualHum = "Hum2010"
            End If
        End If

        '2012 TARDIS (To-Do)
        If e.KeyCode = T2013key Then
            If ActualHum = "Hum2013" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                Dim reader As New WaveFileReader("media/2013Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                PictureBox1.Visible = True
                PictureBox1.Image = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2013Monitor.jpg")
                PictureBox1.BackgroundImageLayout = ImageLayout.Stretch
                ActualHum = "Hum2013"
            End If
        End If

        ' End look and hum

        ' Start travel
        If e.KeyCode = Startkey Then
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
        If e.KeyCode = Endkey Then
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

        ' Time vortex
        If e.KeyCode = TVKey Then

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SpaceEnabled = False
        Travelling = False
        Dim reader As New WaveFileReader("media/2005Hum.wav")
        Dim looping As New LoopStream(reader)
        Hum = New WaveOut()
        Hum.Init(looping)
        Hum.Play()
        ActualHum = "Hum2005"
        videoP.LoadMovie(0, SWFfile)
        videoP.Play()
        videoP.Loop = True
        Label1.Parent = PictureBox1
        Language = System.Globalization.CultureInfo.CurrentCulture.ToString
        If Language.StartsWith("es") Then
            RichTextBox1.LoadFile(Application.StartupPath & "\media\spanish.rtf")
            tabPage1.Text = "Ayuda"
            tabPage2.Text = "Sonido"
            tabPage3.Text = "Controles"
            SoundTabText.Text = "Dispositivos de salida y controles de volumen (Por hacer)."
            ControlsTextBox.Text = "Personalizar los controles (Por hacer)."
            Label1.Text = "No hay vídeo. ¿Tienes uno?"
        Else
            RichTextBox1.LoadFile(Application.StartupPath & "\media\english.rtf")
            SoundTabText.Text = "Output devices and volume settings (TO-DO)."
            ControlsTextBox.Text = "Personalize the controls (TO-DO)."
            Label1.Text = "Missing video. Do you have one?"
        End If
        ' Setting the keys
        Fullscreenkey = Keys.F11
        Escapekey = Keys.Escape
        Helpkey = Keys.F2
        Startkey = Keys.Enter
        Endkey = Keys.Space
        TVKey = Keys.T
        T2005Key = Keys.D1
        T2010Key = Keys.D2
        T2013key = Keys.D3
    End Sub

    Private Sub videoP_GotFocus1(sender As Object, e As EventArgs) Handles videoP.GotFocus
        Me.Focus()
    End Sub

    Private Sub videoP_MouseCaptureChanged(sender As Object, e As EventArgs) Handles videoP.MouseCaptureChanged
        Me.Focus()
    End Sub

    Private Sub Help_Tick(sender As Object, e As EventArgs) Handles Help.Tick
        If tabControl1.Visible = True Then
            tabControl1.Visible = False
            Cursor.Hide()
        ElseIf tabControl1.Visible = False Then
            tabControl1.Visible = True
            Cursor.Show()
        End If

        Help.Enabled = False
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

    Private Sub RichTextBox1_LinkClicked(sender As Object, e As LinkClickedEventArgs)
        System.Diagnostics.Process.Start(e.LinkText)
    End Sub


End Class
