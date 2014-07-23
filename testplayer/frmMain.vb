
'status - 2 tags:A,l,j,C,N,u,t,y,L,c,K,o,r
Imports asyncSockets

Public Class frmMain
    Dim metadata As New Dictionary(Of String, String)
    Dim WithEvents Client As New IPClient
    Delegate Sub HandlesMessage(Message As clsClientMessage)
    Const MY_SERVER As String = "http://192.168.1.126:9000"

    Dim trackdata As New cTrackInfo
    Dim playlist As New cplayListInfo


    Enum eStreamType
        local
        radio
        pandora
        sirius
        rhapsody
        unknown
    End Enum

#Region "Embedded Classes"
    Class cTrackInfo
        Property ID As String

        Property AlbumArtist As String = ""
        Property Title As String = ""
        Property Artist As String = ""
        Property Album As String = ""

        Property AlbumArtURI As String = ""
        Property AlbumTrackNumber As String = ""
        Property AlbumTrackCount As String = ""
        Property Genre As String = ""
        Property Year As String = ""

        Property Bitrate As String = ""
        Property Duration As String = ""
        Property FileURL As String = ""
        Property StreamType As eStreamType = eStreamType.unknown
    End Class

    Class cplayListInfo

        Property ID As String = ""
        Property NumberOfTracks As String = ""

        Property CurrentTrackIndex As String = "0"
        Property CurrentTrackElapsedTime As String = "00:00"        '// Time
        Property Name As String = ""
        Property timestamp As String = "00:00"                       '// when this changes, it's time to reload the playlist.
        Property tracks As New List(Of cTrackInfo)
        Property shuffle As Boolean = False
        Property repeat As Boolean = False
        Property mode As String = "UNKNOWN"

    End Class


#End Region



#Region "Initialization and TearDown"
    Private Sub Init()
        '// to do
    End Sub

    Private Sub TearDown()
        If Client.IsConnected Then Client.ShutDown()
    End Sub

#End Region

#Region "RichTextBox Routines"

    Private Sub RTBAddText(rtb As RichTextBox, text As String, textColor As Color, size As Single, Optional isBold As Boolean = False, Optional CRLF As Boolean = True, Optional clearBox As Boolean = False)
        If clearBox Then
            rtb.Clear()
        End If
        rtb.Select(rtb.TextLength, 0)
        Dim fnt As Font = rtb.SelectionFont

        Dim style As FontStyle = FontStyle.Regular
        If isBold Then style = FontStyle.Bold

        rtb.SelectionFont = New Font(fnt.Name, size, style)

        rtb.SelectionColor = textColor
        rtb.AppendText(text)
        If CRLF Then rtb.AppendText(vbCrLf)
    End Sub

    Private Sub LBLAddText(lbl As Label, Text As String, textColor As Color, size As Single, Optional CRLF As Boolean = True, Optional clearLabel As Boolean = False)
        If clearLabel Then lbl.Text = ""
        Dim fnt As Font = lbl.Font
        If fnt.Size <> size Then
            lbl.Font = New Font(fnt.Name, size)
        End If
        lbl.ForeColor = textColor
        lbl.Text += Text
        If CRLF Then lbl.Text += vbCrLf
    End Sub

    Private Sub TestLoad()
        RTBAddText(RichTextBox1, "Highway To Hell", Color.White, 20, True, True, True)
        RTBAddText(RichTextBox1, "AC/DC", Color.White, 20, False, False)
        RTBAddText(RichTextBox1, ", Various Artists", Color.Silver, 20, False, True)
        RTBAddText(RichTextBox1, "The Very Best of Metal", Color.Silver, 20, False, False)
        RTBAddText(RichTextBox1, " (1999)", Color.DarkGray, 20, False, True)

        RTBAddText(RichTextBox1, "", Color.Silver, 20, True)
        LBLAddText(Label1, "Track:" & vbTab & "1", Color.DarkGray, 16, True, True)
        LBLAddText(Label1, "Duration:" & vbTab & Now.ToShortTimeString, Color.DarkGray, 16)
        LBLAddText(Label1, "Queue:" & vbTab & "11 of 82", Color.DarkGray, 16, False)
    End Sub

#End Region

#Region "Callbacks and Callback Processing"

    Private Sub Client_ClientMessage(Message As clsClientMessage) Handles Client.ClientMessage
        If Me.InvokeRequired() Then
            Me.Invoke(New HandlesMessage(AddressOf OnClientMessage), Message)
        Else
            OnClientMessage(Message)
        End If
    End Sub

    Private Sub OnClientMessage(Message As clsClientMessage)
        Select Case Message.MessageType
            Case clsClientMessage.eIPMessageType.Connected
                lblStatus.Text = Message.MessageText
            Case clsClientMessage.eIPMessageType.Disconnected
                lblStatus.Text = Message.MessageText
            Case clsClientMessage.eIPMessageType.MessageReceived
                TextBox1.Text = Message.MessageText
                ProcessMessage(Message.MessageText)
            Case clsClientMessage.eIPMessageType.MessageSent
                lblStatus.Text = "Sent: " & Message.MessageText
            Case clsClientMessage.eIPMessageType.ErrorConnecting
                lblStatus.Text = "error:" & Message.MessageText

        End Select
    End Sub


    Private Sub ProcessMessage(msg As String)
        Dim params() As String = msg.Split(" ")
        ListBox1.Items.Clear()
        metadata.Clear()
        For Each item As String In params
            'item = URLDecode2(item)
            item = Uri.UnescapeDataString(item)


            Dim nameval() As String = item.Split(":")
            Dim attr As String = nameval(0)
            Dim value As String = item.Replace(nameval(0) & ":", "")

            If nameval.Length > 1 Then
                If Not metadata.ContainsKey(attr) Then
                    metadata.Add(attr, value)

                End If
                ListBox1.Items.Add(attr & ":" & value)
            Else
                ListBox1.Items.Add(item)
            End If

        Next
        ProcessMetaData()
    End Sub

    Sub ProcessMetaData()
        trackdata = New cTrackInfo
        playlist = New cplayListInfo
        For Each item As KeyValuePair(Of String, String) In metadata
            Select Case item.Key
                Case "title"
                    trackdata.Title = item.Value
                Case "artist", "trackartist"
                    trackdata.Artist = item.Value
                Case "band", "albumartist"
                    trackdata.AlbumArtist = item.Value
                Case "album"
                    trackdata.Album = item.Value
                Case "year"
                    trackdata.Year = item.Value
                Case "tracknum"
                    trackdata.AlbumTrackNumber = Val(item.Value) + 1
                Case "duration"
                    Dim iSpan As TimeSpan = TimeSpan.FromSeconds(Val(item.Value))

                    trackdata.Duration = iSpan.Minutes.ToString.PadLeft(2, "0"c) & ":" & _
                        iSpan.Seconds.ToString.PadLeft(2, "0"c)
                Case "coverid"
                    trackdata.AlbumArtURI = MY_SERVER & "/music/" & item.Value.Trim & "/cover.jpg"
                Case "playlist_tracks"
                    playlist.NumberOfTracks = item.Value
                Case "playlist index"
                    playlist.CurrentTrackIndex = Val(item.Value) + 1
                Case "playlist_timestamp"
                    playlist.timestamp = item.Value
                Case "playlist_id"
                    playlist.ID = item.Value
                Case "artwork_url"
                    trackdata.AlbumArtURI = item.Value

                Case "time"
                    Dim iSpan As TimeSpan = TimeSpan.FromSeconds(Val(item.Value))

                    playlist.CurrentTrackElapsedTime = iSpan.Minutes.ToString.PadLeft(2, "0"c) & ":" & _
                        iSpan.Seconds.ToString.PadLeft(2, "0"c)
                Case "url"

                    If item.Value.StartsWith("file:") Then
                        trackdata.StreamType = eStreamType.local
                    ElseIf item.Value.StartsWith("pandora") Then
                        trackdata.StreamType = eStreamType.pandora
                    ElseIf item.Value.StartsWith("sirius") Then
                        trackdata.StreamType = eStreamType.sirius
                    ElseIf item.Value.StartsWith("rhapsody") Then
                        trackdata.StreamType = eStreamType.rhapsody
                    ElseIf item.Value.StartsWith("http") Then
                        trackdata.StreamType = eStreamType.radio
                    Else
                        trackdata.StreamType = eStreamType.unknown
                    End If
                    trackdata.FileURL = item.Value
                Case "remote_title", "current_title", "playlist_name"
                    playlist.Name = item.Value
                Case "shuffle"
                    playlist.shuffle = item.Value
                Case "repeat"
                    playlist.repeat = item.Value
                Case "bitrate"
                    trackdata.Bitrate = item.Value
                Case "mode"
                    playlist.mode = item.Value
            End Select
        Next

        '// clean up weird radio stuff
        If trackdata.Title.StartsWith("text=") Then
            trackdata.Title = trackdata.Title.Split(" ")(0).Split(Chr(34))(1)
        End If

        '// deal with radio related imageart
        If trackdata.AlbumArtURI.StartsWith("imageproxy/") Then
            trackdata.AlbumArtURI = MY_SERVER & "/" & trackdata.AlbumArtURI
            trackdata.AlbumArtURI = Uri.UnescapeDataString(trackdata.AlbumArtURI)
        End If

        If trackdata.AlbumArtURI.StartsWith("/plugins") Then
            trackdata.AlbumArtURI = MY_SERVER & trackdata.AlbumArtURI
            trackdata.AlbumArtURI = Uri.UnescapeDataString(trackdata.AlbumArtURI)
        End If
        Debug.Print("Art:" & trackdata.AlbumArtURI)
        UpdateGUIElements()

    End Sub

    Sub UpdateGUIElements()
        RTBAddText(RichTextBox1, trackdata.Title, Color.White, 20, True, True, True)
        RTBAddText(RichTextBox1, trackdata.Artist, Color.White, 20, False, False)
        '// handle the case where Artist and AlbumArtist are  different
        If (trackdata.Artist <> trackdata.AlbumArtist) And trackdata.AlbumArtist <> "" Then
            RTBAddText(RichTextBox1, ", " & trackdata.AlbumArtist, Color.Silver, 20, False, True)
        Else
            RTBAddText(RichTextBox1, " ", Color.Silver, 20, False, True)
        End If

        RTBAddText(RichTextBox1, trackdata.Album, Color.Silver, 20, False, False)
        '// If we have a valid year
        If trackdata.Year <> "" And Val(trackdata.Year) <> 0 Then
            RTBAddText(RichTextBox1, " (" & trackdata.Year & ")", Color.DarkGray, 20, False, True)
        Else
            RTBAddText(RichTextBox1, "", Color.DarkGray, 20, False, True)
        End If


        RTBAddText(RichTextBox1, "", Color.Silver, 20, True)
        LBLAddText(Label1, "Track:" & vbTab & trackdata.AlbumTrackNumber, Color.DarkGray, 16, True, True)
        LBLAddText(Label1, "Bitrate:" & vbTab & trackdata.Bitrate, Color.DarkGray, 16, True)
        LBLAddText(Label1, "Position:" & vbTab & playlist.CurrentTrackElapsedTime, Color.DarkGray, 16, True)
        LBLAddText(Label1, "Duration:" & vbTab & trackdata.Duration, Color.DarkGray, 16)
        LBLAddText(Label1, "Queue:" & vbTab & " " & playlist.CurrentTrackIndex & " of " & playlist.NumberOfTracks, Color.DarkGray, 16, True)
        Dim prefix As String = ""

        Select Case trackdata.StreamType
            Case eStreamType.local
                If playlist.Name <> "" Then
                    prefix = "Playlist: "
                Else
                    prefix = "Local Library: "
                End If

            Case eStreamType.pandora
                prefix = "Pandora: "
            Case eStreamType.rhapsody
                prefix = "Rhapsody: "
            Case eStreamType.sirius
                prefix = "Sirius: "
            Case eStreamType.radio
                prefix = "Radio: "
            Case eStreamType.unknown
                prefix = "Unknown: "
        End Select

        If playlist.Name <> "" Then
            LBLAddText(Label1, prefix & vbTab & " " & playlist.Name, Color.DarkGray, 16, False)
        Else
            LBLAddText(Label1, prefix & vbTab & " " & trackdata.Album, Color.DarkGray, 16, False)

        End If

        If trackdata.AlbumArtURI <> "" Then
            '// we have a URL!
            If PictureBox1.ImageLocation <> trackdata.AlbumArtURI Then
                '// its different than the current
                'PictureBox1.Load(Nothing)           'delete current
                PictureBox1.Image = Nothing
                'PictureBox1.WaitOnLoad = False
                PictureBox1.LoadAsync(trackdata.AlbumArtURI)
            End If
        Else
            PictureBox1.Image = Nothing           'delete current
        End If
    End Sub
#End Region

#Region "Form Events"

    Private Sub PictureBox1_LoadCompleted(sender As Object, e As System.ComponentModel.AsyncCompletedEventArgs) Handles PictureBox1.LoadCompleted
        If e.Error IsNot Nothing Then
            Debug.Print("ERROR: " & e.Error.Message)
        ElseIf e.Cancelled Then
            Debug.Print("CANCELLED: " & e.Cancelled)
        Else
            Debug.Print("Load Complete")
        End If

        'Debug.Print("Picture Loaded?:" & e.Error.Message & "|" & e.Cancelled.ToString & "|" & e.UserState.ToString)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Init()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        TearDown()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        TextBox1.Text = ""
        ListBox1.Items.Clear()
    End Sub

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        ConnecttoLMS(txtIP.Text, Val(txtPort.Text))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ProcessMessage(TextBox1.Text)
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        SendMessage(txtSend.Text)
    End Sub

    Private Sub btnDisconnect_Click(sender As Object, e As EventArgs) Handles btnDisconnect.Click
        Client.ShutDown()
    End Sub

#End Region

#Region "TCP Client Functions"
    Private Sub ConnecttoLMS(addr As String, Port As Integer)
        Client.Connect("192.168.1.126", 9090)
    End Sub

    Private Sub SendMessage(message As String)
        If Not Client.IsConnected Then ConnecttoLMS(txtIP.Text, Val(txtPort.Text))

        Client.SendMessage(message & vbCrLf)
    End Sub
#End Region


End Class