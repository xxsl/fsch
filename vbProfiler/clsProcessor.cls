VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const EMPTY_COMMAND As Byte = 255

Public logx                 As clsLog

Private lastCommand         As Byte

Private strings             As New Collection

Public samples              As New Collection

Public members              As New Collection

Public Sub processCommand(ByRef data As clsSocketData, ByRef Socket As JBSOCKETSERVERLib.ISocket)

    Dim result As Boolean
    
    readCommand data
    
    If (lastCommand <> EMPTY_COMMAND) Then
        result = processInternal(data, Socket)

        Do While result
            lastCommand = EMPTY_COMMAND
            readCommand data
            result = processInternal(data, Socket)
        Loop

    End If

End Sub

Public Sub readCommand(ByRef data As clsSocketData)

    If (lastCommand = EMPTY_COMMAND And data.size > 0) Then
    
        lastCommand = read8X(data)
        
        'logx.xDebug "command is " & lastCommand
               
    End If

End Sub

Public Function processInternal(ByRef data As clsSocketData, ByRef Socket As JBSOCKETSERVERLib.ISocket) As Boolean

    Dim str       As New clsString

    Dim swf       As New clsSWF
    
    Dim sample    As New clsSample
    
    Dim newSample As New clsNewObjectSample
    
    Dim delSample As New clsDeleteObjectSample
    
    Dim gc        As New clsGC
    
    Dim result    As Boolean
        
    If (lastCommand <> EMPTY_COMMAND) Then
        
        If (str.canDeserialise(lastCommand)) Then
            If (str.bytesAvailable(data)) Then
                str.parse data
                logx.xDebug "string=" & str.getValue

                If (strings.count < str.getKey) Then
                    strings.Add str.getValue
                End If

                result = True
            End If

        ElseIf (swf.canDeserialise(lastCommand)) Then

            If (swf.bytesAvailable(data)) Then
                swf.parse data
                logx.xInfo "swf size=" & swf.getSize
                logx.xInfo "swf url=" & strings.Item(swf.getKey)
                'logger "swf loaded"

                'send commands
                Dim commands(0 To 6) As Byte

                commands(0) = C_MEMORY_PROFILING
                commands(1) = 1
                commands(2) = C_MEMORY_PROFILING_STACKTRACES
                commands(3) = 1
                commands(4) = C_PERFORMANCE_PROFILING
                commands(5) = 1
                commands(6) = C_RESUME
                Socket.Write commands
                result = True
            End If

        ElseIf (sample.canDeserialise(lastCommand)) Then

            If (sample.bytesAvailable(data)) Then

                Set sample.strings = strings
                sample.parse data
                '                Dim line As clsStackTraceLine
                '                logger "sample size=" & sample.getValue.size
                '                For Each line In sample.getValue.lines
                '                    logger "line:" & line.name & ":" & line.fileName & "(" & line.line & ")"
                '                Next
                
                '                logger "sample loaded"
                result = True
            End If

        ElseIf (newSample.canDeserialise(lastCommand)) Then

            If (newSample.bytesAvailable(data)) Then
                Set newSample.strings = strings
                newSample.parse data
                logx.xDebug "create: id=" & newSample.getId & " type=" & newSample.getType & " size=" & newSample.getSize
                samples.Add newSample, "" & newSample.getId
                result = True
                
            End If

        ElseIf (delSample.canDeserialise(lastCommand)) Then

            If (delSample.bytesAvailable(data)) Then
                delSample.parse data
                logx.xDebug "delete: id=" & delSample.getId & " size=" & newSample.getSize
                samples.Remove "" & delSample.getId
                result = True
            End If
        ElseIf (gc.canDeserialise(lastCommand)) Then

            If (gc.bytesAvailable(data)) Then
                gc.parse data
                logx.xInfo "gc"
                result = True
            End If
        End If
    
    End If

    processInternal = result
End Function


Private Sub Class_Initialize()
    lastCommand = EMPTY_COMMAND
End Sub


'    void process()
'            throws IOException
'    {
'        int command;
'        while ((command = this.in.read()) != -1)
'        {
'            Switch (Command)
'            {
'                Case 0:
'                    readSample();
'                    break;
'                Case 1:
'                    readNewObjectSample();
'                    break;
'                Case 2:
'                    readDeleteObjectSample();
'                    break;
'                Case 3:
'                    readMembers();
'                    break;
'                Case 4:
'                    readSwf();
'                    break;
'                Case 5:
'                    this.playerSession.done();
'                    break;
'                Case 6:
'                    readString();
'                    break;
'                Case 7:
'                    readCounts();
'                    break;
'            }
'
'            throw new IOException("Unknown 'agent -> profiler' command: " + command);
'        }
'    }

