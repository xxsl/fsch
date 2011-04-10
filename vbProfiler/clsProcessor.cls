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

Public Sub processCommand(ByRef Data As clsSocketData)

    Dim result As Boolean
    
    readCommand Data
    
    If (lastCommand <> EMPTY_COMMAND) Then
        result = processInternal(Data)

        Do While result
            lastCommand = EMPTY_COMMAND
            readCommand Data
            result = processInternal(Data)
        Loop

    End If

End Sub

Public Sub readCommand(ByRef Data As clsSocketData)

    If (lastCommand = EMPTY_COMMAND And Data.Size > 0) Then
    
        lastCommand = read8X(Data)
        
        'logx.xDebug "command is " & lastCommand
               
    End If

End Sub

Public Function processInternal(ByRef Data As clsSocketData) As Boolean

    Dim str       As New clsString

    Dim swf       As New clsSWF
    
    Dim sample    As New clsSample
    
    Dim newSample As New clsNewObjectSample
    
    Dim delSample As New clsDeleteObjectSample
    
    Dim gc        As New clsGC
    
    Dim result    As Boolean
        
    If (lastCommand <> EMPTY_COMMAND) Then
        
        If (str.canDeserialise(lastCommand)) Then
            If (str.bytesAvailable(Data)) Then
                str.parse Data
                logx.xDebug "string=" & str.getValue

                If (strings.Count < str.getKey) Then
                    strings.Add str.getValue
                End If

                result = True
            End If

        ElseIf (swf.canDeserialise(lastCommand)) Then

            If (swf.bytesAvailable(Data)) Then
                swf.parse Data
                logx.xInfo "swf size=" & swf.getSize
                logx.xInfo "swf url=" & strings.item(swf.getKey)
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
                Data.socket.SendData commands
                result = True
            End If

        ElseIf (sample.canDeserialise(lastCommand)) Then

            If (sample.bytesAvailable(Data)) Then

                Set sample.strings = strings
                sample.parse Data
                '                Dim line As clsStackTraceLine
                '                logger "sample size=" & sample.getValue.size
                '                For Each line In sample.getValue.lines
                '                    logger "line:" & line.name & ":" & line.fileName & "(" & line.line & ")"
                '                Next
                
                '                logger "sample loaded"
                result = True
            End If

        ElseIf (newSample.canDeserialise(lastCommand)) Then

            If (newSample.bytesAvailable(Data)) Then
                Set newSample.strings = strings
                newSample.parse Data
                logx.xDebug "create: id=" & newSample.getId & " type=" & newSample.getType & " size=" & newSample.getSize
                samples.Add newSample, "" & newSample.getId
                result = True
                
            End If

        ElseIf (delSample.canDeserialise(lastCommand)) Then

            If (delSample.bytesAvailable(Data)) Then
                delSample.parse Data
                logx.xDebug "delete: id=" & delSample.getId & " size=" & newSample.getSize
                samples.Remove "" & delSample.getId
                result = True
            End If
        ElseIf (gc.canDeserialise(lastCommand)) Then

            If (gc.bytesAvailable(Data)) Then
                gc.parse Data
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

