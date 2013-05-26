Imports NAudio.Wave
Imports System.Threading

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

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        ' Help key
        If e.KeyCode = Keys.F2 Then
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
        If e.KeyCode = Keys.F11 Then
            If Me.WindowState = FormWindowState.Maximized Then
                Me.WindowState = FormWindowState.Normal
            ElseIf Me.WindowState = FormWindowState.Normal Then
                Me.WindowState = FormWindowState.Maximized

            End If
        End If

        ' Exit key
        If e.KeyCode = Keys.Escape Then
            End
        End If

        ' Look and hum
        ' 2005 TARDIS
        If e.KeyCode = Keys.D1 Then
            If ActualHum = "Hum2005" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                videoP.Visible = True
                Dim reader As New WaveFileReader("media/2005Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                ActualHum = "Hum2005"
            End If
        End If

        ' 2010 TARDIS
        If e.KeyCode = Keys.D2 Then
            If ActualHum = "Hum2010" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                videoP.Visible = False
                Dim reader As New WaveFileReader("media/2010Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                ActualHum = "Hum2010"
            End If
        End If

        '2012 TARDIS (To-Do)
        If e.KeyCode = Keys.D3 Then
            If ActualHum = "Hum2013" Then
            Else
                Hum.Stop()
                Hum.Dispose()
                videoP.Visible = False
                Dim reader As New WaveFileReader("media/2013Hum.wav")
                Dim looping As New LoopStream(reader)
                Hum = New WaveOut()
                Hum.Init(looping)
                Hum.Play()
                ActualHum = "Hum2013"
            End If
        End If

        ' End look and hum

        ' Start travel
        If e.KeyCode = Keys.Enter Then
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
        If e.KeyCode = Keys.Space Then
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
        If e.KeyCode = Keys.T Then

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
        Me.BackgroundImage = System.Drawing.Image.FromFile(Application.StartupPath & "\media\2010Monitor.jpg")
        Me.BackgroundImageLayout = ImageLayout.Zoom
    End Sub

    Private Sub videoP_GotFocus1(sender As Object, e As EventArgs) Handles videoP.GotFocus
        Me.Focus()
    End Sub

    Private Sub videoP_MouseCaptureChanged(sender As Object, e As EventArgs) Handles videoP.MouseCaptureChanged
        Me.Focus()
    End Sub

    Private Sub label2_Click(sender As Object, e As EventArgs) Handles label2.Click
        System.Diagnostics.Process.Start("http://girl-on-the-moon.deviantart.com/art/THE-ICONS-OF-RASSILON-116296443")
    End Sub

    Private Sub label5_Click(sender As Object, e As EventArgs) Handles label5.Click
        System.Diagnostics.Process.Start("http://insertcredit.net/doctorwho/tardisscreen/")
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
End Class
