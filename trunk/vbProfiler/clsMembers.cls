VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMembers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'    private void readMembers() throws IOException
'    {
'        int count = read32();
'
'        for (int i = 0; i < count; i++)
'        {
'            int id = read32();
'            int length = read32();
'
'            if (length <= 0)
'                continue;
'            NameValuePairs nameValuePairs = new NameValuePairs(length);
'
'            for (int j = 0; j < length; j++)
'            {
'                String memberName = this.playerSession.getString(read32());
'                boolean isId = readBoolean();
'
'                if (isId)
'                {
'                    int memberId = read32();
'                    nameValuePairs.set(j, memberName, memberId);
'                } else
'                {
'                    nameValuePairs.set(j, memberName);
'                }
'            }
'
'            this.playerSession.members(id, nameValuePairs);
'        }
'
'        this.playerSession.membersDone();
'    }

Public strings As Collection

Public members As Collection

Private id     As Long

Private size   As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_MEMBERS)
End Function

Public Sub parse(ByRef data As clsSocketData)

    Dim count As Long, i As Long, j As Long

    count = read32X(data)
    
    For i = 0 To count

        Dim id As Long, length As Long

        id = read32X(data)
        length = read32X(data)
        
        If (length <= 0) Then

            Dim nameValue As New clsNameValuePairs

            nameValue.setSize length

            For j = 0 To length

                Dim name As String, isId As Boolean, memberId As Long

                name = strings.Item(read32X(data))
                isId = (read32X(data) = 1)

                If (isId) Then
                    memberId = read32X(data)
                    nameValue.setData j, name, memberId
                Else
                    nameValue.setData j, name
                End If

            Next j

        End If

        members.Add nameValue
    Next i

End Sub

Public Function bytesAvailable(ByRef data As clsSocketData) As Boolean

    Dim count As Long, Result As Boolean, pos As Long, i As Long, j As Long

    Result = True

    If (data.size >= 4) Then
    
        count = read32P(data.getBytes, pos)

        For i = 0 To count

            Dim id As Long, length As Long
            
            If (data.size >= pos + 8 And Result) Then
                
                id = read32P(data.getBytes, pos)
                length = read32P(data.getBytes, pos)
                
                If (length <= 0) Then
    
                    For j = 0 To length
    
                        Dim name As String, isId As Boolean, memberId As Long
    
                        If (data.size >= pos + 8) Then
    
                            name = strings.Item(read32P(data.getBytes, pos))
                            isId = (read32P(data.getBytes, pos) = 1)
                            
                            If (data.size >= pos + 4 Or Not isId) Then
                            
                                If (isId) Then
                                    memberId = read32P(data.getBytes, pos)
                                End If

                            Else
                            
                                Result = False
                                
                                Exit For
                            
                            End If

                        Else
                        
                            Result = False
                            
                            Exit For
                            
                        End If
    
                    Next j
    
                End If
    
            Else
            
                Result = False

                Exit For

            End If

        Next i

    Else
        Result = False
    End If

    bytesAvailable = Result
End Function

Public Function getId() As Long
    getId = id
End Function

Public Function getSize() As Long
    getSize = size
End Function

