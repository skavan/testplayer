Imports System.Threading
Public Module modUtils
    '// Label a thread and return
    Public Function ThreadID() As String
        Static ThreadCounter As Integer
        Dim T As Thread = Thread.CurrentThread
        If T.Name = "" Then
            ThreadCounter += 1
            T.Name = "Thread " & ThreadCounter
        End If
        Return T.Name
    End Function

    Public Sub DebugMsg(ByVal Message As String)
#If DEBUG Then
        Dim stackFrame As New Diagnostics.StackFrame(1)
        Dim callingSub As String = stackFrame.GetMethod.Name.ToString()
        callingSub = stackFrame.GetMethod.DeclaringType.FullName.ToString() & "." & callingSub
        Debug.Print(callingSub & ". " & Message & " [" & ThreadID() & "]" & " [ThreadCount:" & Process.GetCurrentProcess().Threads.Count & "]")
#End If
    End Sub

    Public Function ThreadCount() As String
        Dim Threads As Integer
        Dim ThreadMax As Integer
        System.Threading.ThreadPool.GetAvailableThreads(Threads, ThreadMax)
        Return Threads & "/" & ThreadMax
    End Function
End Module
