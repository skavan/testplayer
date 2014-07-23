Imports System.Text
Imports System.Threading
Imports System.Net
Imports System.Net.Sockets
Imports System.ComponentModel

Public Class classServerMessage
    Public Enum eIPMessageType
        Listening = 1
        ClientConnected = 2
        MessageReceived = 3
        MessageSent = 4
        Disconnected = 5
        CommandPair = 6
    End Enum
    Public MessageType As eIPMessageType
    Public MessageText As String
    Public Client As SocketClient
End Class

Public Class IPServer

#Region "Objects and Variables"

    Public Enum eConnectionType
        PersistantConnection = 0
        ConnectReceiveClientDisconnect = 1
        ConnectReceiveServerDisconnect = 2
    End Enum

    Public Enum eMessageType
        Delimited = 0
        Raw = 1
    End Enum

    Public Enum eResponseType
        None = 0
        OK = 1
        Mirror = 2
        BroadcastAll = 4
        BroadcastExSender = 8
    End Enum

    Public Delegate Function MessageHandler(ByVal Message As classServerMessage) As String

    '// Exposed variables
    Public Shared ServerIP As String = "127.0.0.1"
    Public Shared ServerPort As Integer = 8000
    Public Shared MessageTerminator As String = vbCrLf
    Public Shared ResponseType As eResponseType = eResponseType.None


    ''' <summary>
    ''' Asynchronous message pump that delivers messages across threads. The event handler must be located on the thread that instantiated the class.
    ''' </summary>
    ''' <remarks>The event through which messages are sent back to the host program/class.
    ''' the trick is to make sure RaiseEvent is called on the thread that created the class (i.e. the Hosts thread).
    ''' </remarks>
    Public Shared Event AsyncMessage(ByVal Message As classServerMessage)

    '// Internal "globals"
    Private Shared _listener As Socket
    Private Shared _messageDelegate As MessageHandler
    Private Shared _ConnectionType As eConnectionType = eConnectionType.PersistantConnection
    Private Shared _MessageType As eMessageType = eMessageType.Delimited
    Private Shared _ClientArray As ArrayList = New ArrayList

    '// Async threading variables
    Private Shared _AsyncOperation As AsyncOperation
    Private Shared _instance As Threading.SendOrPostCallback

    'Private receiveDone As New ManualResetEvent(False)
    Private Shared _timerDelegate As TimerCallback = AddressOf OnTimeout
    Private Shared _stateTimer As Timer

#End Region


    Public Sub New()
        '// used for posting messages back to the calling program
        _AsyncOperation = AsyncOperationManager.CreateOperation(Nothing)
        '// used to RaiseEvents back to the host program on its thread.
        _instance = New Threading.SendOrPostCallback(AddressOf TriggerMessageEvent)
    End Sub

    Public Sub New(ByVal MessageProcessingFunction As MessageHandler)
        '// used for posting messages back to the calling program
        _AsyncOperation = AsyncOperationManager.CreateOperation(Nothing)
        '// used to RaiseEvents back to the host program on its thread.
        _instance = New Threading.SendOrPostCallback(AddressOf TriggerMessageEvent)

        _messageDelegate = MessageProcessingFunction
    End Sub

    '// Fire up the server on the provided IP and Port
    '// Then return control to the calling program
    Public Shared Sub StartServer(ByVal ServerAddress As String, ByVal Port As Integer, _
                            ByVal connectionType As eConnectionType, _
                            ByVal messageType As eMessageType)

        _ConnectionType = connectionType
        _MessageType = messageType

        '// Setup Server IP Settings
        If ServerAddress <> "" Then ServerIP = ServerAddress
        If Port <> 0 Then ServerPort = Port

        Dim localIPAddress As IPAddress = IPAddress.Parse(ServerIP)
        Dim localEndPoint As New IPEndPoint(localIPAddress, ServerPort)

        '// Initialize Server and Start Listening
        DebugMsg("Listening on " & ServerIP & " Port " & ServerPort)


        _listener = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        _listener.Bind(localEndPoint)
        _listener.Listen(10)
        _listener.BeginAccept(AddressOf OnConnectRequest, _listener)
        SendMessages(Nothing, classServerMessage.eIPMessageType.Listening, "Listening on " & ServerIP & " Port " & ServerPort)
        GC.Collect()
        GC.WaitForPendingFinalizers()

    End Sub

    '// Assign New Connections to the client list then
    '// loop around (async) waiting for another connection request
    '// Setup the ReceivingData Handler via the NewConnection function.
    Private Shared Sub OnConnectRequest(ByVal ar As IAsyncResult)

        Dim listener As Socket = CType(ar.AsyncState, Socket)
        OnConnection(listener.EndAccept(ar))
        listener.BeginAccept(AddressOf OnConnectRequest, listener)

    End Sub

    Public Shared Function IsConnected() As Boolean
        Return _listener.Connected
    End Function

    '// The Callback function triggered by a ConnectionRequest
    '// Within this function, we 'arm' the OnDataReceived callback.
    Private Shared Sub OnConnection(ByVal sockClient As Socket)

        Dim client As SocketClient = New SocketClient(sockClient)
        '// Add the client to the list of clients
        _ClientArray.Add(client)
        DebugMsg("Client " & client.Sock.RemoteEndPoint.ToString & ", joined")

        SendMessages(client, classServerMessage.eIPMessageType.ClientConnected, "Client " & client.Sock.RemoteEndPoint.ToString & ", joined")
        '// Send a Welcome message back.
        Dim now As DateTime = DateTime.Now
        Dim strDateLine As String = "Welcome " + now.ToString("G") + "" '& Microsoft.VisualBasic.Chr(10) & "" & Microsoft.VisualBasic.Chr(13) & ""
        Dim byteDateLine As Byte() = System.Text.Encoding.ASCII.GetBytes(strDateLine.ToCharArray)

        client.Sock.Send(byteDateLine, byteDateLine.Length, 0)

        '// Rearm the ReceiveCallBack
        client.SetupReceiveCallback()

    End Sub

    '// The Callback function triggered after a Socket.BeginReceive
    '// It is triggered when a message is received (bytesRead>0) or
    '// when the connection is lost (bytesRead=0).
    Public Shared Sub OnDataReceived(ByVal ar As IAsyncResult)
        '// dispose of the timeout timer.
        If Not IsNothing(_stateTimer) Then _stateTimer.Dispose()

        Dim content As String = ""

        ' Retrieve the state object and the handler socket
        ' from the asynchronous state object.
        Dim client As SocketClient = CType(ar.AsyncState, SocketClient)
        Dim handler As Socket = client.Sock

        ' Read data from client socket. 
        Dim bytesRead As Integer = 0

        Try
            bytesRead = handler.EndReceive(ar)
        Catch ex As Exception
            Debug.Print("CONNECTION FAILURE - Bytes Read: " & bytesRead)
        End Try


        ' If bytesRead>0 then a message has been received.
        ' Else, the socket disconnected.
        If bytesRead > 0 Then

            '// There might be more data, so store the data received so far.
            client.Message += (Encoding.ASCII.GetString(client.buffer, 0, bytesRead))

            Dim sp, ep As Integer
            Dim messages As New List(Of String)

            If _MessageType = eMessageType.Delimited Then

                '// If we're in MessageTerminator mode then extract the messages.
                If client.Message.Contains(MessageTerminator) Then
                    Do While client.Message.IndexOf(MessageTerminator, sp) > -1
                        ep = client.Message.IndexOf(MessageTerminator, sp)
                        messages.Add(client.Message.Substring(sp, ep - sp))
                        sp = ep + MessageTerminator.Length
                    Loop
                    If sp < (client.Message.Length) Then
                        client.Message = client.Message.Substring(sp, client.Message.Length - sp)
                    Else
                        client.Message = ""
                    End If
                End If

            End If

            '// and send each message to the parent object in a safe, cross threaded way.
            For Each message As String In messages

                SendMessages(client, classServerMessage.eIPMessageType.MessageReceived, message)

                '// the sleep is to give the receiving host some time to breathe between messages.
                'System.Threading.Thread.Sleep(15)
                DebugMsg("Message Received: " & message)
            Next

            '// if required, disconnect the client and destroy the timer.
            'If Me.ConnectionType = eConnectionType.ConnectReceiveServerDisconnect Then
            '    DisconnectClient(client)
            '    If Not IsNothing(stateTimer) Then stateTimer.Dispose()
            '    Return
            'End If
            '// re-engage the callback
            client.SetupReceiveCallback()

        Else
            '// bytesRead =0 - if the connection was disconnected, clean it up.
            CloseClient(client)

            'client.Sock.Shutdown(SocketShutdown.Both)
            'client.Sock.Close()
            DebugMsg("0 Bytes Arrived:" & ThreadCount())

        End If

        '// just dispose of the timeout timer again for safety. Weird things happen when one doesn't.
        If Not IsNothing(_stateTimer) Then _stateTimer.Dispose()
        '// trigger the timeout timer
        _stateTimer = New Timer(_timerDelegate, client, 100, Timeout.Infinite)

    End Sub

    '// The callback function triggered after a Socket.BeginSend
    Private Shared Sub OnDataSent(ByVal ar As IAsyncResult)
        ' Retrieve the state object and the handler socket
        ' from the asynchronous state object.
        Dim client As SocketClient = CType(ar.AsyncState, SocketClient)
        Dim handler As Socket = client.Sock

        Dim bytesSent As Integer = handler.EndSend(ar)
        DebugMsg("OnDataSent")
        If _ConnectionType = eConnectionType.ConnectReceiveServerDisconnect Then
            DisconnectClient(client)
        End If
    End Sub

    '// Called when we've timed out waiting for a new OnReceivedData event.
    Private Shared Sub OnTimeout(ByVal stateinfo As Object)
        Dim client As SocketClient = CType(stateinfo, SocketClient)
        If client.Message <> "" Then
            DebugMsg("Timeout Message:" & client.Message)
            SendMessages(client, classServerMessage.eIPMessageType.MessageReceived, client.Message)
            client.Message = ""
            'client.sb = New StringBuilder
        Else
            DebugMsg("Timeout Empty")
        End If
        If _ConnectionType = eConnectionType.ConnectReceiveServerDisconnect Then
            'DisconnectClient(client)
        End If
    End Sub

    '// Stop a client from Sending/Receiving
    '// This triggers an OnDataReceive of 0 bytes.
    Private Shared Sub DisconnectClient(ByVal client As SocketClient)
        If Not IsNothing(_stateTimer) Then _stateTimer.Dispose()
        client.Sock.Shutdown(SocketShutdown.Both)
    End Sub

    '// Close the client and release resources
    Private Shared Sub CloseClient(ByVal client As SocketClient)
        DebugMsg("Client: " & client.FriendlyName & ", disconnected")
        _ClientArray.Remove(client)
        SendMessages(client, classServerMessage.eIPMessageType.Disconnected, "Client " & client.Sock.RemoteEndPoint.ToString & ", disconnected")
    End Sub

    Public Shared Sub SendMessage(ByVal Address As String, ByVal Port As String, ByVal Message As String)
        For Each clientSend As SocketClient In _ClientArray
            If clientSend.ClientIP = Address And clientSend.ClientPort.ToString = Port Then
                Dim byteDateLine As Byte() = Text.Encoding.ASCII.GetBytes(Message.ToCharArray)
                clientSend.Sock.Send(byteDateLine)
            End If
        Next
    End Sub
    '// Synchronous Send
    Public Shared Sub BroadcastMessage(ByVal Message As String)
        '// broadcast message to all clients
        For Each clientSend As SocketClient In _ClientArray
            Try
                Dim byteDateLine As Byte() = Text.Encoding.ASCII.GetBytes(Message.ToCharArray)
                clientSend.Sock.Send(byteDateLine)
            Catch
                DisconnectClient(clientSend)
                Return
            End Try
        Next
    End Sub

    '// DataProcessing
    Private Sub MessageProcessing(ByVal client As SocketClient, ByVal Message As String)
        'If Client.Name = "" Then Client.Name = Client.Sock.RemoteEndPoint.ToString

        '// Name Value Pairs
        If Message.Contains("=") Then
            Dim Name As String
            Dim Value As String
            Name = Message.Split("=")(0)
            Value = Message.Split("=")(1)
            Select Case Name.ToLower
                Case "name"
                    client.Name = Value
            End Select
        End If

    End Sub

    '// Receives a message on the senders thread and using m_AsyncOperation
    '// pushes the message across to the thread held by m_instance
    Private Shared Sub SendMessages(ByVal Client As SocketClient, ByVal MessageType As classServerMessage.eIPMessageType, ByVal Message As String)

        Dim mySync As New Threading.SynchronizationContext
        Dim myMessage As classServerMessage

        myMessage = CreateMessage(Client, MessageType, Message)
        Dim response As String = _messageDelegate(myMessage)

        '// synchronous message processing and async sending or the response.
        If response <> "" Then
            If MessageType = classServerMessage.eIPMessageType.MessageReceived Then
                Dim byteData As Byte() = Encoding.ASCII.GetBytes(response)
                'Client.Sock.Send(byteData)
                Client.Sock.BeginSend(byteData, 0, byteData.Length, 0, New AsyncCallback(AddressOf OnDataSent), Client)
            End If
        End If

        '// async send
        _AsyncOperation.Post(_instance, myMessage)


    End Sub

    '// A helper function to create a clientMessageObject
    Private Shared Function CreateMessage(ByVal Client As SocketClient, ByVal MessageType As classServerMessage.eIPMessageType, ByVal Message As String)
        Dim ClientMessage As New classServerMessage
        ClientMessage.MessageType = MessageType
        ClientMessage.MessageText = Message
        ClientMessage.Client = Client
        Return ClientMessage
    End Function

    '// The delegated function to trigger an event to the host program.
    Private Shared Sub TriggerMessageEvent(ByVal oMessage As Object)
        Dim cMessage As classServerMessage = CType(oMessage, classServerMessage)
        RaiseEvent AsyncMessage(cMessage)
    End Sub

    '// Define how the server behaves when it receives messages.
    '// Does it keep or expect a persistent connection?
    '// Does it disconnect a client after processing a message?
    Public Shared Sub ReconfigureServer(ByVal connectionType As eConnectionType, ByVal messageType As eMessageType)
        _ConnectionType = connectionType
        _MessageType = messageType
    End Sub

    '// Retrieve the current List Of Clients
    Public Shared ReadOnly Property ClientList() As ArrayList
        Get
            Return _ClientArray
        End Get
    End Property

End Class

Public Class SocketClient

    Private m_sock As Socket
    'Private m_byBuff As Byte() = New Byte(50) {}
    Private _Name As String
    Private _FriendlyName As String
    Private _ClientIP As String
    Private _ClientPort As Integer

    Public MessageCount As Integer
    Public ProcessedCount As Integer
    Public IsProcessingMessage As Boolean
    Public PartialMessage As String
    Public ReceiveDone As New ManualResetEvent(False)
    Public ReceivePause As New ManualResetEvent(False)

    Public Const BufferSize As Integer = 50
    Public buffer(BufferSize) As Byte
    '// Received data string.
    Public sb As New Text.StringBuilder()
    Public Message As String


    Public Sub New(ByVal sock As Socket)
        m_sock = sock
        _ClientIP = CType(sock.RemoteEndPoint, IPEndPoint).Address.ToString
        _ClientPort = CType(sock.RemoteEndPoint, IPEndPoint).Port
        _FriendlyName = "[" & m_sock.RemoteEndPoint.ToString & "]"
    End Sub

    Public ReadOnly Property Sock() As Socket
        Get
            Return m_sock
        End Get
    End Property

    Public Sub SetupReceiveCallback() 'ByVal app As IPServer
        Try
            Dim receiveData As AsyncCallback = New AsyncCallback(AddressOf IPServer.OnDataReceived)

            m_sock.BeginReceive(buffer, 0, buffer.Length, SocketFlags.None, receiveData, Me)
        Catch ex As SocketException
            DebugMsg("Receive callback setup failed! " & ex.Message & "|" & ex.ErrorCode & "|" & m_sock.RemoteEndPoint.ToString)

        End Try
    End Sub

    Public Function GetRecievedData(ByVal ar As IAsyncResult) As Byte()
        Dim nBytesRec As Integer = 0
        Try
            nBytesRec = m_sock.EndReceive(ar)
        Catch
        End Try
        Dim byReturn(nBytesRec) As Byte
        Array.Copy(buffer, byReturn, nBytesRec)
        Debug.Print("BUFFER:[" & Text.Encoding.ASCII.GetString(byReturn).Trim(Chr(0)) & "]")
        Return byReturn
    End Function

    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
            _FriendlyName = (value & " [" & m_sock.RemoteEndPoint.ToString & "]").Trim
        End Set
    End Property

    Public ReadOnly Property FriendlyName() As String
        Get
            Return _FriendlyName
        End Get
    End Property

    Public ReadOnly Property ClientIP() As String
        Get
            Return _ClientIP
        End Get

    End Property

    Public ReadOnly Property ClientPort() As Integer
        Get
            Return _ClientPort
        End Get

    End Property



End Class