Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.ComponentModel
Imports System.Text

'// This delegate is used to capture inbound messages from the socket connection.
' Delegate Sub dgMessageReceived(ByVal sNewMessage As String)

'// An object to hold relevant message settings.
'// The MessageType signals the type of event and
'// "Message" carries the payload.
Public Class clsClientMessage

    Public Enum eIPMessageType
        Connecting = 1
        Connected = 2
        MessageReceived = 3
        MessageSent = 4
        Disconnecting = 5
        Disconnected = 6
        ErrorConnecting = 7
    End Enum
    Public MessageType As eIPMessageType
    Public MessageText As String
End Class

'// The core Client Engine. It uses sockets asynchronously, returning control to the calling
'// host immediately. It uses AsyncOperations to signal events back to the host.
Public Class IPClient

#Region "Objects and Variables"
    Const BUFFERLEN As Integer = 10000
    '// key internal variables and objects.
    Private m_sock As Socket
    Private m_byBuff As Byte() = New Byte(BUFFERLEN) {}
    Private m_shuttingDown As Boolean = False

    '// the event through which messages are sent back to the host program/class.
    '// the trick is to make sure RaiseEvent is called on the thread
    '// that created the class (i.e. the Hosts thread).
    Public Event ClientMessage(ByVal Message As clsClientMessage)

    '// Async threading variables
    Private m_AsyncOperation As AsyncOperation
    Private m_instance As Threading.SendOrPostCallback
    'Private m_MessageReceived As dgMessageReceived

#End Region

    '// Upon creation, it is important to initialize m_AsyncOperation so that it captures the "creator thread".
    '// i.e. The host programs thread.
    '// We also setup the two delegate pointers.
    Public Sub New()
        '// used for posting messages back to the calling program
        m_AsyncOperation = AsyncOperationManager.CreateOperation(Nothing)
        '// used for collecting messages from IP Operations
        'm_MessageReceived = New dgMessageReceived(AddressOf OnMessageReceived)
        '// used to RaiseEvents back to the host program on its thread.
        m_instance = New Threading.SendOrPostCallback(AddressOf TriggerMessageEvent)
    End Sub

    Public Function IsConnected() As Boolean
        If IsNothing(m_sock) Then
            Return False
        Else
            Return m_sock.Connected
        End If
    End Function

    '// Connect to a Target Machine.
    Public Function Connect(ByVal TargetIP As String, ByVal TargetPort As Integer) As Boolean

        '// Close existing socket if we try to connect again!
        If Not (m_sock Is Nothing) AndAlso m_sock.Connected Then
            m_sock.Shutdown(SocketShutdown.Both)
            System.Threading.Thread.Sleep(10)
            m_sock.Close()
        End If

        '// Set new socket requirements
        m_sock = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        m_sock.Blocking = False

        '// Define Target Endpoint
        Dim targetServer As IPEndPoint = New IPEndPoint(IPAddress.Parse(TargetIP), TargetPort)

        '// Define callback and start connection proceedings.
        Dim ConnectCallback As AsyncCallback = AddressOf Me.OnConnect

        '// The callback function will get called when the server responds.
        m_sock.BeginConnect(targetServer, ConnectCallback, m_sock)
        DebugMsg("")
        '// Control will return immediately to the caller
        Return True
    End Function

    '// Called by BeginConnect : while trying or failing to establish a connection.
    '// Also sets up the listening mode.
    Public Sub OnConnect(ByVal ar As IAsyncResult)
        DebugMsg("")
        Dim sock As Socket = CType(ar.AsyncState, Socket)

        '// if we're connected...start listening.
        If sock.Connected Then
            Dim sMessage As String = "Connected."
            SendCrossThreadMessage(clsClientMessage.eIPMessageType.Connected, sMessage)
            ReceiveData(sock)
        Else
            '// process error condition here.
            DebugMsg("Connection Failed. Unable to connect to remote machine.")
            Dim sMessage As String = "Failed To Connect."
            SendCrossThreadMessage(clsClientMessage.eIPMessageType.ErrorConnecting, sMessage)
            'MsgBox("Connection Failed. Unable to connect to remote machine.")
        End If
        'Catch ex As Exception
        '    MessageBox.Show(Me, ex.Message, "Unusual error during Connect!")
        'End Try
    End Sub

    Public Sub OnClose(ByVal ar As IAsyncResult)
        DebugMsg("")
        Dim sock As Socket = CType(ar.AsyncState, Socket)
        sock.EndDisconnect(ar)
        'm_sock.Close()
        Dim sMessage As String = "Disconnected."
        SendCrossThreadMessage(clsClientMessage.eIPMessageType.Disconnected, sMessage)
    End Sub

    '// Set up a CallBack for inbound messages
    Public Sub ReceiveData(ByVal sock As Socket)

        Try

            Dim receiveData As AsyncCallback = AddressOf OnReceiveData
            sock.BeginReceive(m_byBuff, 0, m_byBuff.Length, SocketFlags.None, receiveData, sock)
        Catch ex As Exception
            MsgBox(ex.Message & ": Setup Recieve Callback failed!")
        End Try
    End Sub

    '// Process inbound data. Armed by SetupReceiveCallback and Triggered when a message is received.
    Public Sub OnReceiveData(ByVal ar As IAsyncResult)
        DebugMsg("")
        Dim sock As Socket = CType(ar.AsyncState, Socket)
        If m_shuttingDown Then Exit Sub

        Try
            Dim nBytesRec As Integer = sock.EndReceive(ar)
            '// if we've received data then be happy.
            If nBytesRec > 0 Then
                Dim sMessage As String = Encoding.ASCII.GetString(m_byBuff, 0, nBytesRec)
                'm_MessageReceived(sReceived)

                'OnMessageReceived(sMEssage)
                SendCrossThreadMessage(clsClientMessage.eIPMessageType.MessageReceived, sMessage)

                '// arm message event capture again.
                ReceiveData(sock)
            Else
                '// must have got disconnected somehow.
                '// do some stuff here
                DebugMsg("The connection has been terminated")
                SendCrossThreadMessage(clsClientMessage.eIPMessageType.Disconnected, "Connection Terminated.")
                If Not (sock Is Nothing) AndAlso sock.Connected Then
                    sock.Shutdown(SocketShutdown.Both)
                    sock.Close()
                End If

            End If
        Catch ex As Exception
            If sock.Connected = False Then
                SendCrossThreadMessage(clsClientMessage.eIPMessageType.Disconnected, "Connection Lost.")
                DebugMsg("Connection Lost.")
            Else
                DebugMsg(ex.Message & " : Unusual error during Receive!")
            End If

        End Try
    End Sub

    '// Send Data out
    Public Sub SendMessage(ByVal Message As String)
        If m_sock Is Nothing OrElse Not m_sock.Connected Then
            DebugMsg("Must be connected to Send a message")
            MsgBox("Must be connected to Send a message")
            Return
        End If
        Try
            Dim byteDateLine As Byte() = Encoding.ASCII.GetBytes(Message.ToCharArray)
            m_sock.Send(byteDateLine, byteDateLine.Length, 0)

        Catch ex As Exception
            MsgBox(Me, ex.Message, "Send Message Failed!")
            DebugMsg(ex.Message & " - Send Message Failed!")
        End Try
    End Sub

    ''// Triggered when m_MessageReceived is fired
    ''// Event triggers here
    'Public Sub OnMessageReceived(ByVal sMessage As String)
    '    'DebugMsg("MESSAGE RECEIVED: " & sMessage)
    '    SendCrossThreadMessage(classClientMessage.eIPMessageType.MessageReceived, sMessage)

    'End Sub

    '// Receive message data from one thread (the calling thread) and invoke a method on a different thread (the m_AsyncOperation thread). 
    Public Sub SendCrossThreadMessage(ByVal MessageType As clsClientMessage.eIPMessageType, ByVal Message As String)
        DebugMsg("")
        Dim mySync As New Threading.SynchronizationContext
        Dim myMessage As clsClientMessage

        myMessage = CreateMessage(MessageType, Message)
        m_AsyncOperation.Post(m_instance, myMessage)
    End Sub

    '// A helper function to create a clientMessageObject
    Private Function CreateMessage(ByVal MessageType As clsClientMessage.eIPMessageType, ByVal Message As String)
        Dim ClientMessage As New clsClientMessage
        ClientMessage.MessageType = MessageType
        ClientMessage.MessageText = Message
        Return ClientMessage
    End Function

    '// TriggerMessageEvent is called through a delegate (SendCrossThreadMessage)
    '// That's why it has an object parameter.
    '// Inside, the object is recast to clientmessage object and an event raised
    '// The event will be raised on the initial class creation thread,
    '// which in turn, should be the thread of the host.

    Public Sub TriggerMessageEvent(ByVal oMessage As Object)

        Dim cMessage As clsClientMessage = CType(oMessage, clsClientMessage)
        DebugMsg(cMessage.MessageText)
        RaiseEvent ClientMessage(cMessage)
    End Sub

    Public Sub ShutDown()
        If Not (m_sock Is Nothing) AndAlso m_sock.Connected Then
            '// Define callback and start connection proceedings.
            Dim DisconnectCallback As AsyncCallback = AddressOf Me.OnClose
            m_sock.Shutdown(SocketShutdown.Both)
            m_sock.BeginDisconnect(True, DisconnectCallback, m_sock)
            m_shuttingDown = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not (m_sock Is Nothing) AndAlso m_sock.Connected Then
            Try
                m_sock.Shutdown(SocketShutdown.Both)
            Catch ex As Exception
                Debug.Print("Seems to be shutdown already!")
            End Try

            m_sock.Close()
        End If
        MyBase.Finalize()
    End Sub

End Class

