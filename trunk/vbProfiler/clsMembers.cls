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

Private Size   As Long

Private dataSize As Long

Public Function canDeserialise(firstByte As Byte) As Boolean
    canDeserialise = (firstByte = C_MEMBERS)
End Function

Public Sub parse(ByRef Data As clsSocketData)

    Dim Count As Long, i As Long, j As Long, lng As Long

    
    Data.read32 Count
    
    For i = 0 To Count

        Dim id As Long, Length As Long

        Data.read32 id
        Data.read32 Length

        
        If (Length <= 0) Then

            Dim nameValue As New clsNameValuePairs

            nameValue.setSize Length

            For j = 0 To Length

                Dim NAME As String, isId As Boolean, memberId As Long

                Data.read32 lng
                
                NAME = strings.item(lng)

                Data.read32 lng
                
                isId = (lng = 1)

                If (isId) Then
                    Data.read32 memberId
                    nameValue.setData j, NAME, memberId
                Else
                    nameValue.setData j, NAME
                End If

            Next j

        End If

        members.Add nameValue
    Next i

End Sub

Public Function bytesAvailable(ByRef Data As clsSocketData) As Boolean

    Dim Count As Long, result As Boolean, pos As Long, i As Long, j As Long, lng As Long

    dataSize = 0
    result = True
    Data.getPos pos


    If (Data.bytesAvailable >= 4) Then
        Data.getPos pos
    
        Data.read32P pos, Count
        
        dataSize = dataSize + 4

        For i = 0 To Count

            Dim id As Long, Length As Long
            
            If (Data.bytesAvailable >= pos + 8 And result) Then
                
                Data.read32P pos, id
                Data.read32P pos, Length
                
                dataSize = dataSize + 8
                
                If (Length <= 0) Then
    
                    For j = 0 To Length
    
                        Dim NAME As String, isId As Boolean, memberId As Long
    
                        If (Data.bytesAvailable >= pos + 8) Then
                            
                            Data.read32P pos, lng
                            
                            NAME = strings.item(lng)
                            
                            Data.read32P pos, lng
                            isId = (lng = 1)
                            
                            dataSize = dataSize + 8
                            
                            If (Data.bytesAvailable >= pos + 4 Or Not isId) Then
                            
                                If (isId) Then
                                    Data.read32P pos, memberId
                                    dataSize = dataSize + 4
                                End If

                            Else
                            
                                result = False
                                
                                Exit For
                            
                            End If

                        Else
                        
                            result = False
                            
                            Exit For
                            
                        End If
    
                    Next j
    
                End If
    
            Else
            
                result = False

                Exit For

            End If

        Next i

    Else
        result = False
    End If

    bytesAvailable = result
End Function

Public Function getId() As Long
    getId = id
End Function

Public Function getSize() As Long
    getSize = Size
End Function

