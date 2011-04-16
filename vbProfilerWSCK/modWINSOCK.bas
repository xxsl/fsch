Attribute VB_Name = "modWINSOCK"
Option Explicit


'=============================================================================================================
'
' modWINSOCK Module
' -----------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Last Update : February 27, 2001
'
' VB Versions : 5.0 / 6.0
'
' Requires    : At least Winsock version 1.1 or an operating system that comes with at least Winsock v1.1.
'               Some Winsock APIs require that Winsock 2.0 be installed.  See below for details.
'
' Description : This module gives you full access to all the documented functions of the WSOCK32.DLL and
'               WS2_32.DLL and all of the types (structures) and constants that are required to make use
'               of those functions.
'
' See Also:
' ---------
' http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/winsock/ovrvw3_1436.htm
' http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/winsock/ovrvw3_8vaq.htm
' http://www.stardust.com/winsock/ws_src.htm
' http://www.sockets.com/
' http://www.vbip.com/default.asp
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================



'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' The following table is an alphabetical list of the functions provided by the Windows Sockets API.
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  accept
'  AcceptEx
'  bind
'  closesocket
'  Connect
'  EnumProtocols
'  GetAcceptExSockaddrs
'  GetAddressByName
'  gethostbyaddr
'  gethostbyname
'  gethostname
'  GetNameByType
'  getpeername
'  getprotobyname
'  getprotobynumber
'  getservbyname
'  getservbyport
'  GetService
'  getsockname
'  getsockopt
'  GetTypeByName
'  htonl
'  htons
'  inet_addr
'  inet_ntoa
'  ioctlsocket
'  listen
'  ntohl
'  ntohs
'  recv
'  recvfrom
'  select
'  send
'  sendto
'  SetService
'  setsockopt
'  shutdown
'  socket
'  TransmitFile
'  WSAAccept
'  WSAAddressToString
'  WSAAsyncGetHostByAddr
'  WSAAsyncGetHostByName
'  WSAAsyncGetProtoByName
'  WSAAsyncGetProtoByNumber
'  WSAAsyncGetServByName
'  WSAAsyncGetServByPort
'  WSAAsyncSelect
'  WSACancelAsyncRequest
'  WSACancelBlockingCall
'  WSACleanup
'  WSACloseEvent
'  WSAConnect
'  WSACreateEvent
'  WSADuplicateSocket
'  WSAEnumNameSpaceProviders
'  WSAEnumNetworkEvents
'  WSAEnumProtocols
'  WSAEventSelect
'  WSAGetLastError
'  WSAGetOverlappedResult
'  WSAGetQOSByName
'  WSAGetServiceClassInfo
'  WSAGetServiceClassNameByClassId
'  WSAHtonl
'  WSAHtons
'  WSAInstallServiceClass
'  WSAIoctl
'  WSAIsBlocking
'  WSAJoinLeaf
'  WSALookupServiceBegin
'  WSALookupServiceEnd
'  WSALookupServiceNext
'  WSANtohl
'  WSANtohs
'  WSAProviderConfigChange
'  WSARecv
'  WSARecvDisconnect
'  WSARecvEx
'  WSARecvFrom
'  WSARemoveServiceClass
'  WSAResetEvent
'  WSASend
'  WSASendDisconnect
'  WSASendTo
'  WSASetBlockingHook
'  WSASetEvent
'  WSASetLastError
'  WSASetService
'  WSASocket
'  WSAStartup
'  WSAStringToAddress
'  WSAUnhookBlockingHook
'  WSAWaitForMultipleEvents
'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯



'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' The following functions are only available if you have Windows Sockets version 1.1 or greater installed on
' your computer, or if your operating system comes with at least Winsock 1.1 pre-installed:
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  accept
'  AcceptEx              - Not Supported On Windows 95 or Windows 98
'  bind
'  closesocket
'  connect
'  EnumProtocols         - Obsolete in Winsock 2.0
'  GetAcceptExSockaddrs
'  GetAddressByName      - Obsolete in Winsock 2.0
'  gethostbyaddr
'  gethostbyname
'  gethostname
'  GetNameByType         - Obsolete in Winsock 2.0
'  getpeername
'  getprotobyname
'  getprotobynumber
'  getservbyname
'  getservbyport
'  GetService            - Obsolete in Winsock 2.0
'  getsockname           - Not Supported On Windows 95
'  getsockopt
'  GetTypeByName         - Obsolete in Winsock 2.0
'  htonl
'  htons
'  inet_addr
'  inet_ntoa
'  ioctlsocket
'  listen
'  ntohl
'  ntohs
'  recv
'  recvfrom
'  send
'  sendto
'  SetService            - Obsolete in Winsock 2.0 / Not Supported On Windows 95
'  setsockopt
'  shutdown
'  socket
'  TransmitFile          - Not Supported On Windows 95
'  WSAAsyncGetHostByAddr
'  WSAAsyncGetHostByName
'  WSAAsyncGetProtoByName
'  WSAAsyncGetProtoByNumber
'  WSAAsyncGetServByName
'  WSAAsyncGetServByPort
'  WSAAsyncSelect
'  WSACancelAsyncRequest
'  WSACancelBlockingCall - Obsolete in Winsock 2.0
'  WSACleanup
'  WSACloseEvent
'  WSAGetLastError
'  WSAIsBlocking         - Obsolete in Winsock 2.0
'  WSARecvEx             - Not Supported On Windows 95
'  WSASetBlockingHook    - Obsolete in Winsock 2.0
'  WSASetLastError
'  WSAStartup
'  WSAUnhookBlockingHook - Obsolete in Winsock 2.0
'
'_____________________________________________________________________________________________________________
' The following functions are only available if you have Windows Sockets version 2.0 or greater installed on
' your computer, or if your operating system comes with at least Winsock 2.0 pre-installed:
'
' NOTE : Some of these functions can be found in the WSOCK32.DLL (Windows Sockets v1.1)
'
' NOTE : All of the Winsock 1.1 APIs are supported in Winsock 2.0 for backwards compatibility
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'  select
'  WSAAccept
'  WSAAddressToString
'  WSAConnect
'  WSACreateEvent
'  WSADuplicateSocket
'  WSAEnumNameSpaceProviders
'  WSAEnumNetworkEvents
'  WSAEnumProtocols
'  WSAEventSelect
'  WSAGetOverlappedResult
'  WSAGetQOSByName
'  WSAGetServiceClassInfo
'  WSAGetServiceClassNameByClassId
'  WSAHtonl
'  WSAHtons
'  WSAInstallServiceClass
'  WSAIoctl
'  WSAJoinLeaf
'  WSALookupServiceBegin
'  WSALookupServiceEnd
'  WSALookupServiceNext
'  WSANtohl
'  WSANtohs
'  WSAProviderConfigChange
'  WSARecv
'  WSARecvDisconnect
'  WSARecvFrom
'  WSARemoveServiceClass
'  WSAResetEvent
'  WSASend
'  WSASendDisconnect
'  WSASendTo
'  WSASetEvent
'  WSASetService
'  WSASocket
'  WSAStringToAddress
'  WSAWaitForMultipleEvents
'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




'-------------------------------------------------------------------------------------------------------------
' The following are type definitions that are necisary to understand because when making the transfer from
' C (Win32 API) to Visual Basic, you need to know the data type's size in bytes to match it up with the
' correct corisponding VB data type.
'-------------------------------------------------------------------------------------------------------------
' typedef UINT_PTR SOCKET;
' #define WSAEVENT HANDLE
'-------------------------------------------------------------------------------------------------------------




' Type Declarations
Public Type GUID 'The GUID data type is a text string representing a Class identifier(ID). COM must be able to convert the string to a valid Class ID. All GUIDs must be authored in uppercase. The valid format for a GUID is {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} where X is a hex digit (0,1,2,3,4,5,6,7,8,9,A,B,C,D,E,F).
  Data1    As Long
  Data2    As Integer
  Data3    As Integer
  Data4(8) As Byte
End Type

Public Type BLOB 'Requires Windows Sockets 1.1 or later  (Not supported on Windows 95)
  cbSize    As Long 'ULONG  // Size of the block of data pointed to by pBlobData, in bytes.
  pBlobData As Byte 'BYTE * // Pointer to a block of data.
End Type

Public Type HOSTENT 'Requires Windows Sockets 2.0
  hName     As Long    ' char FAR *       - Official name of the host (PC). If using the DNS or similar resolution system, it is the Fully Qualified Domain Name (FQDN) that caused the server to return a reply. If using a local hosts file, it is the first entry after the IP address.
  hAliases  As Long    ' char FAR * FAR * - NULL terminated array of alternate names.
  hAddrType As Integer ' short            - Type of address being returned.
  hLength   As Integer ' short            - Length of each address, in bytes.
  hAddrList As Long    ' char FAR * FAR * - NULL terminated list of addresses for the host. Addresses are returned in network byte order. The macro h_addr is defined to be h_addr_list[0] for compatibility with older software.
End Type

Public Type IN_ADDR 'Requires Windows Sockets 2.0
  s_b(3) As Byte    'struct {u_char s_b1,s_b2,s_b3,s_b4;} // Address of the host formatted as four u_chars
  s_w(1) As Integer 'struct {u_short s_w1,s_w2;}          // Address of the host formatted as two u_shorts
  S_addr As Long    'u_long                               // Address of the host formatted as a u_long
End Type ' NOTE: Whenver a function calls for a "IN_ADDR" structure, use LONG instead and pass the results of inet_addr("xxx.xxx.xxx.xxx")

' #define s_addr  S_un.S_addr      // Can be used for most tcp & ip code
' #define s_host  S_un.S_un_b.s_b2 // Host on imp
' #define s_net   S_un.S_un_b.s_b1 // Network
' #define s_imp   S_un.S_un_w.s_w2 // Imp
' #define s_impno S_un.S_un_b.s_b4 // Imp #
' #define s_lh    S_un.S_un_b.s_b3 // Logical host

Public Type PROTOENT 'Requires Windows Sockets 2.0
  p_name    As Long    'char FAR *       // Official name of the protocol
  p_aliases As Long    'char FAR * FAR * // Null-terminated array of alternate names
  p_proto   As Integer 'short            // Protocol number, in host byte order
End Type

Public Type PROTOCOL_INFO 'Requires Windows Sockets 1.1 or later
  dwServiceFlags As Long   ' DWORD  // A set of bit flags that specifies the services provided by the protocol.  One or more of the XP_* bit flags may be set.
  iAddressFamily As Long   ' INT    // Value to pass as the af parameter when the socket function is called to open a socket for the protocol. This address family value uniquely defines the structure of protocol addresses, also known as sockaddr structures, used by the protocol.
  iMaxSockAddr   As Long   ' INT    // Maximum length of a socket address supported by the protocol.
  iMinSockAddr   As Long   ' INT    // Minimum length of a socket address supported by the protocol.
  iSocketType    As Long   ' INT    // Value to pass as the type parameter when the socket function is called to open a socket for the protocol. Note that if XP_PSEUDO_STREAM is set in dwServiceFlags, the application can specify SOCK_STREAM as the type parameter to socket, regardless of the value of iSocketType.
  iProtocol      As Long   ' INT    // Value to pass as the protocol parameter when the socket function is called to open a socket for the protocol.
  dwMessageSize  As Long   ' DWORD  // Maximum message size supported by the protocol. This is the maximum size of a message that can be sent from or received by the host. For protocols that do not support message framing, the actual maximum size of a message that can be sent to a given address may be less than this value.  The following special message size values are defined:
                           '             &H0        = The protocol is stream-oriented; the concept of message size is not relevant.
                           '             &HFFFFFFFF = The protocol is message-oriented, but there is no maximum message size.
  lpProtocol     As String ' LPTSTR // Points to a zero-terminated string that supplies a name for the protocol; for example, "SPX2"
End Type

Public Type SERVENT 'Requires Windows Sockets 2.0
  s_name    As Long    'char FAR *       // Official name of the service.
  s_aliases As Long    'char FAR * FAR * // Null-terminated array of alternate names.
  s_port    As Integer 'short            // Port number at which the service can be contacted. Port numbers are returned in network byte order.
  s_proto   As Long    'char FAR *       // Name of the protocol to use when contacting the service.
End Type

Public Type SERVICE_ADDRESS 'Requires Windows Sockets 2.0
  dwAddressType     As Long 'DWORD  // Address family to which the socket address pointed to by lpAddress belongs.
  dwAddressFlags    As Long 'DWORD  // Set of bit flags that specify properties of the address. The following bit flags are defined: SERVICE_ADDRESS_FLAG_RPC_CN, SERVICE_ADDRESS_FLAG_RPC_DG, SERVICE_ADDRESS_FLAG_RPC_NB
  dwAddressLength   As Long 'DWORD  // Size, in bytes, of the address.
  Reserved1         As Long 'DWORD  // Reserved for future use. Must be zero.
  lpAddress         As Byte 'BYTE * // Pointer to a socket address of the appropriate type.
  Reserved2         As Byte 'BYTE * // Reserved for future use. It must be null.
End Type

Public Type SOCKADDR 'Requires Windows Sockets 2.0 (This structure is used with TCP/IP)
  sin_family As Integer    'short
  sin_port   As Integer    'u_short
  sin_addr   As Long       'struct IN_ADDR
  sin_zero   As String * 7 'char[8]
End Type

Public Type SOCKADDR1 'Requires Windows Sockets 2.0 (This is the more generic version of the structure)
  sa_Family As Integer     'u_short
  sa_Data   As String * 13 'char[14]
End Type

Public Type SOCKET_ADDRESS 'Requires Windows Sockets 2.0
  lpSockaddr      As SOCKADDR 'LPSOCKADDR // Pointer to a socket address
  iSockaddrLength As Long     'INT        // Length of the socket address, in bytes.
End Type

Public Type SERVICE_ADDRESSES
  dwAddressCount As Long            'DWORD              // Specifies the number of SERVICE_ADDRESS structures in the Addresses array.
  Addresses()    As SERVICE_ADDRESS 'SERVICE_ADDRESS[1] // An array of SERVICE_ADDRESS data structures. Each SERVICE_ADDRESS structure contains information about a network service address.
End Type

Public Type CSADDR_INFO 'Requires Windows Sockets 1.1 or later.
  LocalAddr   As SOCKET_ADDRESS 'SOCKET_ADDRESS // Specifies a Windows Sockets local address. In a client application, pass this address to the bind function to obtain access to a network service. In a network service, pass this address to the bind function so that the service is bound to the appropriate local address.
  RemoteAddr  As SOCKET_ADDRESS 'SOCKET_ADDRESS // Specifies a Windows Sockets remote address. There are several uses for this remote address:  1) You can use this remote address to connect to the service through the connect function. This is useful if an application performs send/receive operations that involve connection-oriented protocols.  2) You can use this remote address with the sendto function when you are communicating over a connectionless (datagram) protocol. If you are using a connectionless protocol, such as UDP, sendto is typically the way you pass data
  iSocketType As Long           'INT            // Specifies the type of the Windows socket. The following socket types are defined in Winsock.h: SOCK_STREAM, SOCK_DGRAM, SOCK_RDM, SOCK_SEQPACKET
  iProtocol   As Long           'INT            // Specifies a value to pass as the protocol parameter to the socket function to open a socket for this service.
End Type

Public Type SERVICE_INFO ' Requires Windows Sockets 1.1 or later  (Not supported on Windows 95)
  lpServiceType       As GUID              'LPGUID              // Pointer to a GUID that is the type of the network service.
  lpServiceName       As Long              'LPTSTR              // Pointer to a zero-terminated string that is the name of the network service.  If you are calling the SetService function with the dwNameSpace parameter set to NS_DEFAULT, the network service name must be a common name. A common name is what the network service is commonly known as. An example of a common name for a network service is "My SQL Server". If you are calling the SetService function with the dwNameSpace parameter set to a specific service name, the network service name can be a common name or a distinguished name. A distinguished name distinguishes the service to a unique location with a directory service. An example of a distinguished name for a network service is "MS\\SYS\\NT\\DEV\\My SQL Server".
  lpComment           As String            'LPTSTR              // Pointer to a zero-terminated string that is a comment or description for the network service. For example, "Used for development upgrades."
  lpLocale            As String            'LPTSTR              // Pointer to a zero-terminated string that contains locale information.
  dwDisplayHint       As Long              'DWORD               // Specifies a hint as to how to display the network service in a network browsing user interface. This can be one of the RESOURCEDISPLAYTYPE_* following values.
  dwVersion           As Long              'DWORD               // Version information for the network service. The high word of this value specifies a major version number. The low word of this value specifies a minor version number.
  Reserved            As Long              'DWORD               // Reserved for future use. Must be set to zero.
  lpMachineName       As String            'LPTSTR              // Pointer to a zero-terminated string that is the name of the computer on which the network service is running.
  lpServiceAddress    As SERVICE_ADDRESSES 'LPSERVICE_ADDRESSES // Pointer to a SERVICE_ADDRESSES structure that contains an array of SERVICE_ADDRESS structures. Each SERVICE_ADDRESS structure contains information about a network service address.
  ServiceSpecificInfo As BLOB              'BLOB                // A BLOB structure that specifies service-defined information.  Note that In general, the data pointed to by the BLOB structure's pBlobData member must not contain any pointers. That is because only the network service knows the format of the data; copying the data without such knowledge would lead to pointer invalidation. If the data pointed to by pBlobData contains variable-sized elements, offsets from pBlobData can be used to indicate the location of those elements. There is one exception to this general rule: when pBlobData points to a SERVICE_TYPE_INFO_ABS structure. This is possible because both the SERVICE_TYPE_INFO_ABS structure, and any SERVICE_TYPE_VALUE_ABS structures it contains are predefined, and thus their formats are known to the operating system.
End Type

Public Type NS_SERVICE_INFO 'Requires Windows Sockets 1.1 or later  (Not supported on Windows 95)
  dwNameSpace As Long         'DWORD        // Specifies the name space or a set of default name spaces to which this service information applies.  Use one of the following constant values to specify a name space: NS_DEFAULT, NS_DNS, NS_MS, NS_NDS, NS_NETBT, NS_NIS, NS_SAP, NS_STDA, NS_TCPIP_HOSTS, NS_TCPIP_LOCAL, NS_WINS, NS_X500
  ServiceInfo As SERVICE_INFO 'SERVICE_INFO // A SERVICE_INFO structure that contains information about a network service or network service type.
End Type

Public Type FD_SET ' Requires Windows Sockets 2.0
  fd_count     As Long 'u_int               // How many are SET?
  fd_array(63) As Long 'SOCKET [FD_SETSIZE] // An array of SOCKETs
End Type

Public Type TIMEVAL ' Requires Windows Sockets 2.0
  tv_sec  As Long 'long // Seconds
  tv_usec As Long 'long // Microseconds
End Type

Public Type OVERLAPPED
  Internal     As Long 'ULONG_PTR // Reserved for operating system use. This member, which specifies a system-dependent status, is valid when the GetOverlappedResult function returns without setting the extended error information to ERROR_IO_PENDING.
  InternalHigh As Long 'ULONG_PTR // Reserved for operating system use. This member, which specifies the length of the data transferred, is valid when the GetOverlappedResult function returns TRUE.
  Offset       As Long 'DWORD     // Specifies a file position at which to start the transfer. The file position is a byte offset from the start of the file. The calling process sets this member before calling the ReadFile or WriteFile function. This member is ignored when reading from or writing to named pipes and communications devices and should be zero.
  OffsetHigh   As Long 'DWORD     // Specifies the high word of the byte offset at which to start the transfer. This member is ignored when reading from or writing to named pipes and communications devices and should be zero.
  hEvent       As Long 'HANDLE    // Handle to an event set to the signaled state when the operation has been completed. The calling process must set this member either to zero or a valid event handle before calling any overlapped functions. To create an event object, use the CreateEvent function.
End Type

Public Type WSAOVERLAPPED 'Requires Windows Sockets 2.0
  Internal     As Long 'DWORD   // Reserved for internal use. The Internal member is used internally by the entity that implements overlapped I/O. For service providers that create sockets as installable file system (IFS) handles, this parameter is used by the underlying operating system. Other service providers (non-IFS providers) are free to use this parameter as necessary.
  InternalHigh As Long 'DWORD   // Reserved. Used internally by the entity that implements overlapped I/O. For service providers that create sockets as IFS handles, this parameter is used by the underlying operating system. NonIFS providers are free to use this parameter as necessary.
  Offset       As Long 'DWORD   // Reserved for use by service providers.
  OffsetHigh   As Long 'DWORD   // Reserved for use by service providers.
  hEvent       As Long 'WSEVENT // If an overlapped I/O operation is issued without an I/O completion routine (lpCompletionRoutine is null), then this parameter should either contain a valid handle to a WSAEVENT object or be null. If lpCompletionRoutine is non-null then applications are free to use this parameter as necessary.
End Type

Public Type TRANSMIT_FILE_BUFFERS ' Requires Windows Sockets 1.1 or later (Not supported on Windows 95)
  Head       As Long 'PVOID // Pointer to a buffer that contains data to be transmitted before the file data is transmitted.
  HeadLength As Long 'DWORD // Number of bytes in the buffer pointed to by Head that are to be transmitted.
  Tail       As Long 'PVOID // Pointer to a buffer that contains data to be transmitted after the file data is transmitted.
  TailLength As Long 'DWORD // Number of bytes of data in the buffer pointed to by the Tail member that are to be transmitted.
End Type

Public Type FLOWSPEC ' Flow Specifications for each direction of data flow.
  TokenRate          As Long 'uint32      // In Bytes/sec
  TokenBucketSize    As Long 'uint32      // In Bytes
  PeakBandwidth      As Long 'uint32      // In Bytes/sec
  Latency            As Long 'uint32      // In microseconds
  DelayVariation     As Long 'uint32      // In microseconds
  ServiceType        As Long 'ServiceType (uint32)
  MaxSduSize         As Long 'uint32      // In Bytes
  MinimumPolicedSize As Long 'uint32      // In Bytes
End Type

Public Type WSABUF 'Requires Windows Sockets 2.0
  len As Long 'u_long      // The length of the buffer
  buf As Long 'char FAR *  // The pointer to the buffer
End Type

Public Type QOS ' QualityOfService
  SendingFlowspec   As FLOWSPEC 'FLOWSPEC // The flow spec for data sending
  ReceivingFlowspec As FLOWSPEC 'FLOWSPEC // The flow spec for data receiving */
  ProviderSpecific  As WSABUF   'WSABUF   // Additional provider specific stuff */
End Type

Public Type WSAPROTOCOLCHAIN 'Requires Windows Sockets 2.0
  ChainLen        As Long 'int                        // Length of the chain. The following settings apply: Setting ChainLen to zero indicates a layered protocol, Setting ChainLen to one indicates a base protocol, Setting ChainLen to greater than one indicates a protocol chain
  ChainEntries(6) As Long 'DWORD [MAX_PROTOCOL_CHAIN] // Array of protocol chain entries.
End Type

Public Type WSAPROTOCOL_INFO 'Requires Windows Sockets 2.0
  dwServiceFlags1    As Long             'DWORD                     // Bitmask describing the services provided by the protocol. The following values are possible: XP1_CONNECTIONLESS, XP1_GUARANTEED_DELIVERY, XP1_GUARANTEED_ORDER, XP1_MESSAGE_ORIENTED, XP1_PSEUDO_STREAM, XP1_GRACEFUL_CLOSE, XP1_EXPEDITED_DATA, XP1_CONNECT_DATA, XP1_DISCONNECT_DATA, XP1_INTERRUPT, XP1_SUPPORT_BROADCAST, XP1_SUPPORT_MULTIPOINT, XP1_MULTIPOINT_CONTROL_PLANE, XP1_MULTIPOINT_DATA_PLANE, XP1_QOS_SUPPORTED, XP1_UNI_SEND, XP1_UNI_RECV, XP1_IFS_HANDLES, XP1_PARTIAL_MESSAGE
                                         '                             Note that only one of XP1_UNI_SEND or XP1_UNI_RECV may be set. If a protocol can be unidirectional in either direction, two WSAPROTOCOL_INFOW structures should be used. When neither bit is set, the protocol is considered to be bidirectional.
  dwServiceFlags2    As Long             'DWORD                     // Reserved for additional protocol-attribute definitions.
  dwServiceFlags3    As Long             'DWORD                     // Reserved for additional protocol-attribute definitions.
  dwServiceFlags4    As Long             'DWORD                     // Reserved for additional protocol-attribute definitions.
  dwProviderFlags    As Long             'DWORD                     // Provides information about how this protocol is represented in the protocol catalog. The following flag values are possible: PFL_MULTIPLE_PROTO_ENTRIES, PFL_RECOMMENDED_PROTO_ENTRY, PFL_HIDDEN, PFL_MATCHES_PROTOCOL_ZERO
  ProviderId         As GUID             'GUID                      // Globally unique identifier assigned to the provider by the service provider vendor. This value is useful for instances where more than one service provider is able to implement a particular protocol. An application may use the dwProviderId value to distinguish between providers that might otherwise be indistinguishable.
  dwCatalogEntryId   As Long             'DWORD                     // Unique identifier assigned by the WS2_32.DLL for each WSAPROTOCOL_INFOW structure.
  ProtocolChain      As WSAPROTOCOLCHAIN 'WSAPROTOCOLCHAIN          // If the length of the chain is 0, this WSAPROTOCOL_INFOW entry represents a layered protocol which has Windows Sockets 2 SPI as both its top and bottom edges. If the length of the chain equals 1, this entry represents a base protocol whose Catalog Entry identifier is in the dwCatalogEntryId member of the WSAPROTOCOL_INFOW structure. If the length of the chain is larger than 1, this entry represents a protocol chain which consists of one or more layered protocols on top of a base protocol. The corresponding Catalog Entry identifiers are in the ProtocolChain.ChainEntries array starting with the layered protocol at the top (the zero element in the ProtocolChain.ChainEntries array) and ending with the base protocol. Refer to the Windows Sockets 2 Service Provider Interface specification for more information on protocol chains.
  iVersion           As Long             'int                       // Protocol version identifier.
  iAddressFamily     As Long             'int                       // Value to pass as the address family parameter to the socket/WSASocket function in order to open a socket for this protocol. This value also uniquely defines the structure of protocol addresses SOCKADDRs used by the protocol.
  iMaxSockAddr       As Long             'int                       // Maximum address length.
  iMinSockAddr       As Long             'int                       // Minimum address length.
  iSocketType        As Long             'int                       // Value to pass as the socket type parameter to the socket function in order to open a socket for this protocol.
  iProtocol          As Long             'int                       // Value to pass as the protocol parameter to the socket function in order to open a socket for this protocol.
  iProtocolMaxOffset As Long             'int                       // Maximum value that may be added to iProtocol when supplying a value for the protocol parameter to socket and WSASocket. Not all protocols allow a range of values. When this is the case iProtocolMaxOffset is zero.
  iNetworkByteOrder  As Long             'int                       // Currently these values are manifest constants (BIGENDIAN and LITTLEENDIAN) that indicate either big-endian or little-endian with the values 0 and 1 respectively.
  iSecurityScheme    As Long             'int                       // Indicates the type of security scheme employed (if any). A value of SECURITY_PROTOCOL_NONE is used for protocols that do not incorporate security provisions.
  dwMessageSize      As Long             'DWORD                     // Maximum message size supported by the protocol. This is the maximum size that can be sent from any of the host's local interfaces. For protocols that do not support message framing, the actual maximum that can be sent to a given address may be less. There is no standard provision to determine the maximum inbound message size. The following special values are defined:
                                         '                               0          = The protocol is stream-oriented and hence the concept of message size is not relevant.
                                         '                               &H1        = The maximum outbound (send) message size is dependent on the underlying network MTU (maximum sized transmission unit) and hence cannot be known until after a socket is bound. Applications should use getsockopt to retrieve the value of SO_MAX_MSG_SIZE after the socket has been bound to a local address.
                                         '                               &HFFFFFFFF = The protocol is message-oriented, but there is no maximum limit to the size of messages that may be transmitted.
  dwProviderReserved As Long             'DWORD                     // Reserved for use by service providers.
  szProtocol         As String * 255     'TCHAR [WSAPROTOCOL_LEN+1] // Array of characters that contains a human-readable name identifying the protocol, for example "SPX2". The maximum number of characters allowed is WSAPROTOCOL_LEN, which is defined to be 255.
End Type

Public Type WSANAMESPACE_INFO 'Requires Windows Sockets 2.0
  NSProviderId   As GUID   'GUID   // Unique identifier for this name-space provider.
  dwNameSpace    As Long   'DWORD  // Name space supported by this implementation of the provider.
  fActive        As Long   'BOOL   // If TRUE, indicates that this provider is active. If FALSE, the provider is inactive and is not accessible for queries, even if the query specifically references this provider.
  dwVersion      As Long   'DWORD  // Name space–version identifier.
  lpszIdentifier As String 'LPTSTR // Display string for the provider.
End Type

Public Type WSANETWORKEVENTS 'Requires Windows Sockets 2.0
  lNetworkEvents As Long 'long                // Indicates which of the FD_XXX network events have occurred.
  iErrorCode(7)  As Long 'int [FD_MAX_EVENTS] // An array that contains any associated error codes, with an array index that corresponds to the position of event bits in lNetworkEvents. The identifiers FD_READ_BIT, FD_WRITE_BIT and other can be used to index the iErrorCode array.
End Type

Public Type WSANSCLASSINFO
  lpszName    As String 'LPSTR
  dwNameSpace As Long   'DWORD
  dwValueType As Long   'DWORD
  dwValueSize As Long   'DWORD
  lpValue     As Long   'LPVOID
End Type

Public Type WSASERVICECLASSINFO '(WSASERVICECLASSINFOA) Requires Windows Sockets 2.0
  lpServiceClassId     As GUID           'LPGUID            // Unique Identifier (GUID) for the service class.
  lpszServiceClassName As String         'LPSTR             // Well known associated with the service class.
  dwCount              As Long           'DWORD             // Number of entries in lpClassInfos.
  lpClassInfos()       As WSANSCLASSINFO 'LPWSANSCLASSINFOA // Array of WSANSCLASSINFOW structures that contains information about the service class.
End Type

Public Type WSAVERSION
  dwVersion As Long 'DWORD
  ecHow     As Long 'WSAECOMPARATOR // (Can be COMP_EQUAL [0] or COMP_NOTLESS [1])
End Type

Public Type AFPROTOCOLS 'Requires Windows Sockets 2.0
  iAddressFamily As Long 'INT // Address family to which the query is to be constrained.
  iProtocol      As Long 'INT // Protocol to which the query is to be constrained.
End Type

Public Type WSAQUERYSET 'Requires Windows Sockets 2.0
  dwSize                  As Long        'DWORD         // Must be set to sizeof(WSAQUERYSET). This is a versioning mechanism.
  lpszServiceInstanceName As String      'LPTSTR        // Ignored for queries.
  lpServiceClassId        As GUID        'LPGUID        // (Optional) Referenced string contains service name. The semantics for using wildcards within the string are not defined, but can be supported by certain name space providers.
  lpVersion               As WSAVERSION  'LPWSAVERSION  // (Required) The GUID corresponding to the service class.
  lpszComment             As String      'LPTSTR        // (Optional) References desired version number and provides version comparison semantics (that is, version must match exactly, or version must be not less than the value supplied).
  dwNameSpace             As Long        'DWORD         // Ignored for queries.
  lpNSProviderId          As GUID        'LPGUID        // Identifier of a single name space in which to constrain the search, or NS_ALL to include all name spaces.
  lpszContext             As String      'LPTSTR        // (Optional) References the GUID of a specific name-space provider, and limits the query to this provider only.
  dwNumberOfProtocols     As String      'DWORD         // (Optional) Specifies the starting point of the query in a hierarchical name space.
  lpafpProtocols          As AFPROTOCOLS 'LPAFPROTOCOLS // Size of the protocol constraint array, can be zero.
  lpszQueryString()       As String      'LPTSTR        // (Optional) References an array of AFPROTOCOLS structure. Only services that utilize these protocols will be returned.
  dwNumberOfCsAddrs       As Long        'DWORD         // (Optional) Some name spaces (such as Whois++) support enriched SQL-like queries that are contained in a simple text string. This parameter is used to specify that string.
  lpcsaBuffer             As CSADDR_INFO 'LPCSADDR_INFO // Ignored for queries.
  dwOutputFlags           As Long        'DWORD         // Ignored for queries.
  lpBlob                  As BLOB        'LPBLOB        // (Optional) This is a pointer to a provider-specific entity.
End Type

Public Type WSADATA 'Requires Windows Sockets 2.0
  wVersion       As Integer      'WORD                        // Version of the Windows Sockets specification that the Ws2_32.dll expects the caller to use.
  wHighVersion   As Integer      'WORD                        // Highest version of the Windows Sockets specification that this .dll can support (also encoded as above). Normally this is the same as wVersion.
  szDescription  As String * 256 'char [WSADESCRIPTION_LEN+1] // Null-terminated ASCII string into which the Ws2_32.dll copies a description of the Windows Sockets implementation. The text (up to 256 characters in length) can contain any characters except control and formatting characters: the most likely use that an application can put this to is to display it (possibly truncated) in a status message.
  szSystemStatus As String * 128 'char [WSASYS_STATUS_LEN+1]  // Null-terminated ASCII string into which the WSs2_32.dll copies relevant status or configuration information. The Ws2_32.dll should use this parameter only if the information might be useful to the user or support staff: it should not be considered as an extension of the szDescription parameter.
  iMaxSockets    As Integer      'unsigned short              // Retained for backward compatibility, but should be ignored for Windows Sockets version 2 and later, as no single value can be appropriate for all underlying service providers.
  iMaxUdpDg      As Integer      'unsigned short              // Ignored for Windows Sockets version 2 and onward. iMaxUdpDg is retained for compatibility with Windows Sockets specification 1.1, but should not be used when developing new applications. For the actual maximum message size specific to a particular Windows Sockets service provider and socket type, applications should use getsockopt to retrieve the value of option SO_MAX_MSG_SIZE after a socket has been created.
  lpVendorInfo   As Long         'char far*                   // Ignored for Windows Sockets version 2 and onward. It is retained for compatibility with Windows Sockets specification 1.1. Applications needing to access vendor-specific configuration information should use getsockopt to retrieve the value of option PVD_CONFIG. The definition of this value (if utilized) is beyond the scope of this specification.
                                 '                               NOTE : An application should ignore the iMaxsockets, iMaxUdpDg, and lpVendorInfo members in WSAData if the value in wVersion after a successful call to WSAStartup is at least 2. This is because the architecture of Windows Sockets has been changed in version 2 to support multiple providers, and WSAData no longer applies to a single vendor's stack. Two new socket options are introduced to supply provider-specific information: SO_MAX_MSG_SIZE (replaces the iMaxUdpDg element) and PVD_CONFIG (allows any other provider-specific configuration to occur).
End Type

' Constants - General
Public Const MAX_PATH = 260

' Constants - CSADDR_INFO.iSocketType
Public Const SOCK_STREAM = 1    ' Stream. This is a protocol that sends data as a stream of bytes, with no message boundaries.
Public Const SOCK_DGRAM = 2     ' Datagram. This is a connectionless protocol. There is no virtual circuit setup. There are typically no reliability guarantees. Services use recvfrom to obtain datagrams. The listen and accept functions do not work with datagrams.
Public Const SOCK_RAW = 3       ' Raw-protocol interface
Public Const SOCK_RDM = 4       ' Reliably-Delivered Message. This is a protocol that preserves message boundaries in data.
Public Const SOCK_SEQPACKET = 5 ' Sequenced packet stream. This is a protocol that is essentially the same as SOCK_RDM.

' Constants - EnumProtocols.lpiProtocols
Public Const IPPROTO_TCP = 6      ' TCP/IP, a connection/stream-oriented protocol.
Public Const IPPROTO_UDP = 17     ' User Datagram Protocol (UDP/IP), a connectionless datagram protocol.
Public Const ISOPROTO_TP4 = 29    ' ISO connection-oriented transport protocol.
Public Const NSPROTO_IPX = 1000   ' IPX.
Public Const NSPROTO_SPX = 1256   ' SPX.
Public Const NSPROTO_SPXII = 1257 ' SPX II.

' PROTOCOL_INFO.dwServiceFlags
Public Const XP_CONNECTIONLESS = &H1         ' If this flag is set, the protocol provides connectionless (datagram) service. If this flag is clear, the protocol provides connection-oriented data transfer.
Public Const XP_GUARANTEED_DELIVERY = &H2    ' If this flag is set, the protocol guarantees that all data sent will reach the intended destination. If this flag is clear, there is no such guarantee.
Public Const XP_GUARANTEED_ORDER = &H4       ' If this flag is set, the protocol guarantees that data will arrive in the order in which it was sent. Note that this characteristic does not guarantee delivery of the data, only its order. If this flag is clear, the order of data sent is not guaranteed.
Public Const XP_MESSAGE_ORIENTED = &H8       ' If this flag is set, the protocol is message-oriented. A message-oriented protocol honors message boundaries. If this flag is clear, the protocol is stream oriented, and the concept of message boundaries is irrelevant.
Public Const XP_PSEUDO_STREAM = &H10         ' If this flag is set, the protocol is a message-oriented protocol that ignores message boundaries for all receive operations. This optional capability is useful when you do not want the protocol to frame messages. An application that requires stream-oriented characteristics can open a socket with type SOCK_STREAM for transport protocols that support this functionality, regardless of the value of iSocketType.
Public Const XP_GRACEFUL_CLOSE = &H20        ' If this flag is set, the protocol supports two-phase close operations, also known as graceful close operations. If this flag is clear, the protocol supports only abortive close operations.
Public Const XP_EXPEDITED_DATA = &H40        ' If this flag is set, the protocol supports expedited data, also known as urgent data.
Public Const XP_CONNECT_DATA = &H80          ' If this flag is set, the protocol supports connect data.
Public Const XP_DISCONNECT_DATA = &H100      ' If this flag is set, the protocol supports disconnect data.
Public Const XP_SUPPORTS_BROADCAST = &H200   ' If this flag is set, the protocol supports a broadcast mechanism.
Public Const XP_SUPPORTS_MULTICAST = &H400   ' If this flag is set, the protocol supports a multicast mechanism.
Public Const XP_BANDWIDTH_ALLOCATION = &H800 ' If this flag is set, the protocol supports a mechanism for allocating a guaranteed bandwidth to an application.
Public Const XP_FRAGMENTATION = &H1000       ' If this flag is set, the protocol supports message fragmentation; physical network MTU is hidden from applications.
Public Const XP_ENCRYPTS = &H2000            ' If this flag is set, the protocol supports data encryption.

' Constants - GetAddressByName.dwNameSpace / GetService.dwNameSpace / SetService.dwNameSpace
Public Const NS_DEFAULT = 0      ' A set of default name spaces. The function queries each name space within this set. The set of default name spaces typically includes all the name spaces installed on the system. System administrators, however, can exclude particular name spaces from the set. This is the value that most applications should use for dwNameSpace.
Public Const NS_DNS = 12         ' The Domain Name System used in the Internet for host name resolution.
Public Const NS_NDS = 2          ' The NetWare 4 provider.
Public Const NS_NETBT = 13       ' The NetBIOS over TCP/IP layer. All Windows NT/Windows 2000 systems register their computer names with NetBIOS. This name space is used to convert a computer name to an IP address that uses this registration. Note that NS_NETBT can access a WINS server to perform the resolution.
Public Const NS_SAP = 1          ' The Netware Service Advertising Protocol. This can access the Netware bindery if appropriate. NS_SAP is a dynamic name space that allows registration of services.
Public Const NS_TCPIP_HOSTS = 11 ' Lookup value in the <systemroot>\system32\drivers\etc\hosts file.
Public Const NS_TCPIP_LOCAL = 10 ' Local TCP/IP name resolution mechanisms, including comparisons against the local host name and looks up host names and IP addresses in cache of host to IP address mappings.
Public Const NS_PEER_BROWSE = 3
Public Const NS_WINS = 14
Public Const NS_NBP = 20
Public Const NS_MS = 30
Public Const NS_STDA = 31
Public Const NS_NTDS = 32
Public Const NS_X500 = 40
Public Const NS_NIS = 41
Public Const NS_VNS = 50

' Constants - GetAddressByName.dwResolution
Public Const RES_SOFT_SEARCH = &H1   ' This flag is valid if the name space supports multiple levels of searching.  If this flag is valid and set, the operating system performs a simple and quick search of the name space. This is useful if an application only needs to obtain easy-to-find addresses for the service.  If this flag is valid and clear, the operating system performs a more extensive search of the name space.
Public Const RES_FIND_MULTIPLE = &H2 ' If this flag is set, the operating system performs an extensive search of all name spaces for the service. It asks every appropriate name space to resolve the service name. If this flag is clear, the operating system stops looking for service addresses as soon as one is found.
Public Const RES_SERVICE = &H4       ' If set, the function obtains the address to which a service of the specified type should bind. This is the equivalent of setting lpServiceName to NULL.  If this flag is clear, normal name resolution occurs.

' Constants - gethostbyaddr.type
Public Const AF_UNIX = 1        ' Local to host (pipes, portals)
Public Const AF_INET = 2        ' Internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK = 3     ' Arpanet imp addresses
Public Const AF_PUP = 4         ' Pup protocols: e.g. BSP
Public Const AF_CHAOS = 5       ' Mit CHAOS protocols
Public Const AF_NS = 6          ' XEROX NS protocols
Public Const AF_IPX = AF_NS     ' IPX protocols: IPX, SPX, etc.
Public Const AF_ISO = 7         ' ISO protocols
Public Const AF_OSI = AF_ISO    ' OSI is ISO
Public Const AF_ECMA = 8        ' European computer manufacturers
Public Const AF_DATAKIT = 9     ' Datakit protocols
Public Const AF_CCITT = 10      ' CCITT protocols, X.25 etc
Public Const AF_SNA = 11        ' IBM SNA
Public Const AF_DECnet = 12     ' DECnet
Public Const AF_DLI = 13        ' Direct data link interface
Public Const AF_LAT = 14        ' LAT
Public Const AF_HYLINK = 15     ' NSC Hyperchannel
Public Const AF_APPLETALK = 16  ' AppleTalk
Public Const AF_NETBIOS = 17    ' NetBios-style addresses
Public Const AF_VOICEVIEW = 18  ' VoiceView
Public Const AF_FIREFOX = 19    ' Protocols from Firefox
Public Const AF_UNKNOWN1 = 20   ' Somebody is using this!
Public Const AF_BAN = 21        ' Banyan
Public Const AF_ATM = 22        ' Native ATM Services
Public Const AF_INET6 = 23      ' Internetwork Version 6

' Constants - GetService.dwProperties
Public Const PROP_COMMENT = &H1      ' If this flag is set, the function stores data in the lpComment member of the data structures stored in *lpBuffer.
Public Const PROP_LOCALE = &H2       ' If this flag is set, the function stores data in the lpLocale member of the data structures stored in *lpBuffer.
Public Const PROP_DISPLAY_HINT = &H4 ' If this flag is set, the function stores data in the dwDisplayHint member of the data structures stored in *lpBuffer.
Public Const PROP_VERSION = &H8      ' If this flag is set, the function stores data in the dwVersion member of the data structures stored in *lpBuffer.
Public Const PROP_START_TIME = &H10  ' If this flag is set, the function stores data in the dwTime member of the data structures stored in *lpBuffer.
Public Const PROP_MACHINE = &H20     ' If this flag is set, the function stores data in the lpMachineName member of the data structures stored in *lpBuffer.
Public Const PROP_ADDRESSES = &H100  ' If this flag is set, the function stores data in the lpServiceAddress member of the data structures stored in *lpBuffer.
Public Const PROP_SD = &H200         ' If this flag is set, the function stores data in the ServiceSpecificInfo member of the data structures stored in *lpBuffer.
Public Const PROP_ALL = &H80000000   ' If this flag is set, the function stores data in all of the members of the data structures stored in *lpBuffer.

' Constants - SERVICE_ADDRESS.dwAddressFlags
Public Const SERVICE_ADDRESS_FLAG_RPC_CN = &H1 'If this bit flag is set, the service supports connection-oriented RPC over this transport protocol.
Public Const SERVICE_ADDRESS_FLAG_RPC_DG = &H2 'If this bit flag is set, the service supports datagram-oriented RPC over this transport protocol.
Public Const SERVICE_ADDRESS_FLAG_RPC_NB = &H4 'If this bit flag is set, the service supports NetBIOS RPC over this transport protocol.

' Constants - getsockopt.level (See MSDN for information about .level and .optval)
Public Const SOL_SOCKET = &HFFFF
'Public Const IPPROTO_TCP = 6
'Public Const NSPROTO_IPX = 1000

' Constants - getsockopt.optval / setsockopt.optval
Public Const SO_ACCEPTCONN = &H2                 ' Socket is listening.
Public Const SO_BROADCAST = &H20                 ' Socket is configured for the transmission of broadcast messages.
Public Const SO_DEBUG = &H1                      ' Debugging is enabled.
Public Const SO_DONTROUTE = &H10                 ' Routing is disabled. Not supported on ATM sockets.
Public Const SO_ERROR = &H1007                   ' Retrieves error status and clear.
Public Const SO_GROUP_ID = &H2001                ' Reserved.
Public Const SO_GROUP_PRIORITY = &H2002          ' Reserved.
Public Const SO_KEEPALIVE = &H8                  ' Keep-alives are being sent. Not supported on ATM sockets.
Public Const SO_LINGER = &H80                    ' Returns the current linger options.
Public Const SO_DONTLINGER = Not SO_LINGER       ' If TRUE, the SO_LINGER option is disabled.
Public Const SO_MAX_MSG_SIZE = &H2003            ' Maximum size of a message for message-oriented socket types (for example, SOCK_DGRAM). Has no meaning for stream oriented sockets.
Public Const SO_OOBINLINE = &H100                ' OOB data is being received in the normal data stream. (See section Windows Sockets 1.1 Blocking Routines and EINPROGRESS for a discussion of this topic.)
Public Const SO_PROTOCOL_INFO = &H2004           ' Description of protocol information for protocol that is bound to this socket.
Public Const SO_RCVBUF = &H1002                  ' Buffer size for receives.
Public Const SO_RCVLOWAT = &H1004                ' Receives low watermark.
Public Const SO_RCVTIMEO = &H1006                ' Receives time-out (available in Microsoft implementation of Windows Sockets 2).
Public Const SO_REUSEADDR = &H4                  ' The socket can be bound to an address which is already in use. Not applicable for ATM sockets.
Public Const SO_SNDBUF = &H1001                  ' Buffer size for sends.
Public Const SO_SNDLOWAT = &H1003                ' Sends low watermark.
Public Const SO_SNDTIMEO = &H1005                ' Sends time-out (available in Microsoft implementation of Windows Sockets 2).
Public Const SO_TYPE = &H1008                    ' The type of the socket (for example, SOCK_STREAM).
Public Const PVD_CONFIG = &H3001                 ' An opaque data structure object from the service provider associated with socket s. This object stores the current configuration information of the service provider. The exact format of this data structure is service provider specific.
Public Const TCP_NODELAY = &H1                   ' Disables the Nagle algorithm for send coalescing.
Public Const IPX_PTYPE = &H4000                  ' Obtains the IPX packet type.
Public Const IPX_FILTERPTYPE = &H4001            ' Obtains the receive filter packet type
Public Const IPX_DSTYPE = &H4002                 ' Obtain the value of the data stream field in the SPX header on every packet sent.
Public Const IPX_EXTENDED_ADDRESS = &H4004       ' Find out whether extended addressing is enabled.
Public Const IPX_RECVHDR = &H4005                ' Find out whether the protocol header is sent up on all receive headers.
Public Const IPX_MAXSIZE = &H4006                ' Obtain the maximum data size that can be sent.
Public Const IPX_ADDRESS = &H4007                ' Obtain information about a specific adapter to which IPX is bound. Adapter numbering is base zero. The adapternum member is filled in upon return.
Public Const IPX_GETNETINFO = &H4008             ' Obtain information about a specific IPX network number. If not available in the cache, uses RIP to obtain information.
Public Const IPX_GETNETINFO_NORIP = &H4009       ' Obtain information about a specific IPX network number. If not available in the cache, will not use RIP to obtain information, and returns error.
Public Const IPX_SPXGETCONNECTIONSTATUS = &H400B ' Obtains information about a connected SPX socket.
Public Const IPX_ADDRESS_NOTIFY = &H400C         ' Obtains status notification when changes occur on an adapter to which IPX is bound.
Public Const IPX_MAX_ADAPTER_NUM = &H400D        ' Obtains maximum number of adapters present, numbered as base zero.
Public Const IPX_RERIPNETNUMBER = &H400E         ' Similar to IPX_GETNETINFO, but forces IPX to use RIP for resolution, even if the network information is in the local cache.
Public Const IPX_IMMEDIATESPXACK = &H4010        ' Directs SPX connections not to delay before sending an ACK. Applications without back-and-forth traffic should set this to TRUE to increase performance.
Public Const IPX_STOPFILTERPTYPE = &H4003        ' Stop filtering the filter type set with IPX_FILTERTYPE
Public Const IPX_RECEIVE_BROADCAST = &H400F      ' Indicates broadcast packets are likely on the socket. Set to TRUE by default. Applications that do not use broadcasts should set this to FALSE for better system performance.
'Public Const SO_CONDITIONAL_ACCEPT = ?          ' Returns current socket state, either from a previous call to setsockopt or the system default.
'Public Const SO_EXCLUSIVEADDRUSE = ?            ' Enables a socket to be bound for exclusive access. Requires Windows NT 4.0 SP4 or Windows 2000.

' Constants - inet_addr (return)
Public Const INADDR_NONE = &HFFFFFFFF

' Constants - recv.flags / recvfrom.flags / send.flags / sendto.flags / WSARecv.lpFlags
Public Const MSG_OOB = &H1        ' Process out-of-band data
Public Const MSG_PEEK = &H2       ' Peek at incoming message
Public Const MSG_DONTROUTE = &H4  ' Send without using routing tables
Public Const MSG_PARTIAL = &H8000 ' Partial send or recv for message xport

' Constants - SetService.dwOperation
Public Const SERVICE_REGISTER = &H1    ' Register the network service with the name space. This operation can be used with the SERVICE_FLAG_DEFER and SERVICE_FLAG_HARD bit flags.
Public Const SERVICE_DEREGISTER = &H2  ' Remove from the registry the network service from the name space. This operation can be used with the SERVICE_FLAG_DEFER and SERVICE_FLAG_HARD bit flags.
Public Const SERVICE_FLUSH = &H3       ' Perform any operation that was called with the SERVICE_FLAG_DEFER bit flag set to one.
Public Const SERVICE_ADD_TYPE = &H4    ' Add a service type to the name space.  For this operation, use the ServiceSpecificInfo member of the SERVICE_INFO structure pointed to by lpServiceInfo to pass a SERVICE_TYPE_INFO_ABS structure. You must also set the ServiceType member of the SERVICE_INFO structure. Other SERVICE_INFO members are ignored.
Public Const SERVICE_DELETE_TYPE = &H5 ' Remove a service type, added by a previous call specifying the SERVICE_ADD_TYPE operation, from the name space.

' Constants - SetService.dwFlags
Public Const SERVICE_FLAG_DEFER = &H1 ' This bit flag is valid only if the operation is SERVICE_REGISTER or SERVICE_DEREGISTER.  If this bit flag is one, and it is valid, the name-space provider should defer the registration or deregistration operation until a SERVICE_FLUSH operation is requested.
Public Const SERVICE_FLAG_HARD = &H2  ' This bit flag is valid only if the operation is SERVICE_REGISTER or SERVICE_DEREGISTER.  If this bit flag is one, and it is valid, the name-space provider updates any relevant persistent store information when the operation is performed.  For example: If the operation involves deregistration in a name space that uses a persistent store, the name-space provider would remove the relevant persistent store information.

' Constants - SetService.lpdwStatusFlags
Public Const SET_SERVICE_PARTIAL_SUCCESS = &H1 ' One or more name-space providers were unable to successfully perform the requested operation.

' Constants - shutdown.how
Public Const SD_RECEIVE = &H0 ' If the how parameter is SD_RECEIVE, subsequent calls to the recv function on the socket will be disallowed. This has no effect on the lower protocol layers. For TCP sockets, if there is still data queued on the socket waiting to be received, or data arrives subsequently, the connection is reset, since the data cannot be delivered to the user. For UDP sockets, incoming datagrams are accepted and queued. In no case will an ICMP error packet be generated.
Public Const SD_SEND = &H1    ' If the how parameter is SD_SEND, subsequent calls to the send function are disallowed. For TCP sockets, a FIN will be sent after all data is sent and acknowledged by the receiver.
Public Const SD_BOTH = &H2    ' Setting how to SD_BOTH disables both sends and receives as described above.

' Constants - TransmitFile.dwFlags
Public Const TF_DISCONNECT = &H1        ' Start a transport-level disconnect after all the file data has been queued for transmission.
Public Const TF_REUSE_SOCKET = &H2      ' Prepare the socket handle to be reused. When the TransmitFile request completes, the socket handle can be passed to the AcceptEx function. It is only valid if TF_DISCONNECT is also specified.
Public Const TF_WRITE_BEHIND = &H4      ' Complete the TransmitFile request immediately, without pending. If this flag is specified and TransmitFile succeeds, then the data has been accepted by the system but not necessarily acknowledged by the remote end. Do not use this setting with the TF_DISCONNECT and TF_REUSE_SOCKET flags.
'Public Const TF_USE_DEFAULT_WORKER = ? ' Directs the Windows Sockets service provider to use the system's default thread to process long TransmitFile requests. The system default thread can be adjusted using the following registry parameter as a REG_DWORD:CurrentControlSet\Services\afd\Parameters\TransmitWorker
'Public Const TF_USE_SYSTEM_THREAD = ?  ' Directs the Windows Sockets service provider to use system threads to process long TransmitFile requests.
'Public Const TF_USE_KERNEL_APC = ?     ' Directs the driver to use kernel Asynchronous Procedure Calls (APCs) instead of worker threads to process long TransmitFile requests. Long TransmitFile requests are defined as requests that require more than a single read from the file or a cache; the request therefore depends on the size of the file and the specified length of the send packet.  Use of TF_USE_KERNEL_APC can deliver significant performance benefits. It is possible (though unlikely), however, that the thread in which context TransmitFile is initiated is being used for heavy computations; this situation may prevent APCs from launching. Note that the Windows Sockets kernel mode driver uses normal kernel APCs, which launch whenever a thread is in a wait state, which differs from user-mode APCs, which launch whenever a thread is in an alertable wait state initiated in user mode).

' Constants - WSAPROTOCOL_INFO.dwServiceFlags1
Public Const XP1_CONNECTIONLESS = &H1             ' Provides connectionless (datagram) service. If not set, the protocol supports connection-oriented data transfer.
Public Const XP1_GUARANTEED_DELIVERY = &H2        ' Guarantees that all data sent will reach the intended destination.
Public Const XP1_GUARANTEED_ORDER = &H4           ' Guarantees that data only arrives in the order in which it was sent and that it is not duplicated. This characteristic does not necessarily mean that the data is always delivered, but that any data that is delivered is delivered in the order in which it was sent.
Public Const XP1_MESSAGE_ORIENTED = &H8           ' Honors message boundaries—as opposed to a stream-oriented protocol where there is no concept of message boundaries.
Public Const XP1_PSEUDO_STREAM = &H10             ' A message-oriented protocol, but message boundaries are ignored for all receipts. This is convenient when an application does not desire message framing to be done by the protocol.
Public Const XP1_GRACEFUL_CLOSE = &H20            ' Supports two-phase (graceful) close. If not set, only abortive closes are performed.
Public Const XP1_EXPEDITED_DATA = &H40            ' Supports expedited (urgent) data.
Public Const XP1_CONNECT_DATA = &H80              ' Supports connect data.
Public Const XP1_DISCONNECT_DATA = &H100          ' Supports disconnect data.
Public Const XP1_INTERRUPT = &H4000               ' Bit is reserved.
Public Const XP1_SUPPORT_BROADCAST = &H200        ' Supports a broadcast mechanism.
Public Const XP1_SUPPORT_MULTIPOINT = &H400       ' Supports a multipoint or multicast mechanism. Control and data plane attributes are indicated below.
Public Const XP1_MULTIPOINT_CONTROL_PLANE = &H800 ' Indicates whether the control plane is rooted (value = 1) or nonrooted (value = 0).
Public Const XP1_MULTIPOINT_DATA_PLANE = &H1000   ' Indicates whether the data plane is rooted (value = 1) or nonrooted (value = 0).
Public Const XP1_QOS_SUPPORTED = &H2000           ' Supports quality of service requests.
Public Const XP1_UNI_SEND = &H8000                ' Protocol is unidirectional in the send direction.
Public Const XP1_UNI_RECV = &H10000               ' Protocol is unidirectional in the recv direction.
Public Const XP1_IFS_HANDLES = &H20000            ' Socket descriptors returned by the provider are operating system Installable File System (IFS) handles.
Public Const XP1_PARTIAL_MESSAGE = &H40000        ' The MSG_PARTIAL flag is supported in WSASend and WSASendTo.

' Constants - WSAPROTOCOL_INFO.dwProviderFlags
Public Const PFL_MULTIPLE_PROTO_ENTRIES = &H1  ' Indicates that this is one of two or more entries for a single protocol (from a given provider) which is capable of implementing multiple behaviors. An example of this is SPX which, on the receiving side, can behave either as a message-oriented or a stream-oriented protocol.
Public Const PFL_RECOMMENDED_PROTO_ENTRY = &H2 ' Indicates that this is the recommended or most frequently used entry for a protocol that is capable of implementing multiple behaviors.
Public Const PFL_HIDDEN = &H4                  ' Set by a provider to indicate to the Ws2_32.dll that this protocol should not be returned in the result buffer generated by WSAEnumProtocols. Obviously, a Windows Sockets 2 application should never see an entry with this bit set.
Public Const PFL_MATCHES_PROTOCOL_ZERO = &H8   ' Indicates that a value of zero in the protocol parameter of socket or WSASocket matches this protocol entry.

' Constants - WSAAsyncSelect.lEvent / WSAAsyncSelect (Return)
Public Const FD_READ = &H1                       ' INPUT  = Wants to receive notification of readiness for reading.
                                                 ' OUTPUT = Socket ready for reading.
Public Const FD_WRITE = &H2                      ' INPUT  = Wants to receive notification of readiness for writing.
                                                 ' OUTPUT = Socket ready for writing.
Public Const FD_OOB = &H4                        ' INPUT  = Wants to receive notification of the arrival of OOB data.
                                                 ' OUTPUT = OOB data ready for reading on socket.
Public Const FD_ACCEPT = &H8                     ' INPUT  = Wants to receive notification of incoming connections.
                                                 ' OUTPUT = Socket ready for accepting a new incoming connection.
Public Const FD_CONNECT = &H10                   ' INPUT  = Wants to receive notification of completed connection or multipoint join operation.
                                                 ' OUTPUT = Connection or multipoint join operation initiated on socket completed.
Public Const FD_CLOSE = &H20                     ' INPUT  = Wants to receive notification of socket closure.
                                                 ' OUTPUT = Connection identified by socket has been closed.
Public Const FD_QOS = &H40                       ' INPUT  = Wants to receive notification of socket Quality of Service (QOS) changes.
                                                 ' OUTPUT = Quality of Service associated with socket has changed.
Public Const FD_GROUP_QOS = &H80                 ' INPUT  = Reserved
                                                 ' OUTPUT = Reserved
Public Const FD_ROUTING_INTERFACE_CHANGE = &H100 ' INPUT  = Wants to receive notification of routing interface changes for the specified destination(s).
                                                 ' OUTPUT = Local interface that should be used to send to the specified destination has changed.
Public Const FD_ADDRESS_LIST_CHANGE = &H200      ' INPUT  = Wants to receive notification of local address list changes for the socket's protocol family.
                                                 ' OUTPUT = The list of addresses of the socket's protocol family to which the application client can bind has changed.

' Constants - WSALookupServiceBegin.dwControlFlags
Public Const LUP_DEEP = &H1             ' Queries deep as opposed to just the first level.
Public Const LUP_CONTAINERS = &H2       ' Returns containers only.
Public Const LUP_NOCONTAINERS = &H4     ' Does not return any containers.
Public Const LUP_FLUSHCACHE = &H1000    ' If the provider has been caching information, ignores the cache, and queries the name space itself.
Public Const LUP_FLUSHPREVIOUS = &H2000 ' Used as a value for the dwControlFlags parameter in WSALookupServiceNext. Setting this flag instructs the provider to discard the last result set, which was too large for the supplied buffer, and move on to the next result set.
Public Const LUP_NEAREST = &H8          ' If possible, returns results in the order of distance. The measure of distance is provider specific.
Public Const LUP_RES_SERVICE = &H8000   ' This indicates whether prime response is in the remote or local part of CSADDR_INFO structure. The other part needs to be usable in either case.
Public Const LUP_RETURN_ALIASES = &H400 ' Any available alias information is to be returned in successive calls to WSALookupServiceNext, and each alias returned will have the RESULT_IS_ALIAS flag set.
Public Const LUP_RETURN_NAME = &H10     ' Retrieves the name as lpszServiceInstanceName.
Public Const LUP_RETURN_TYPE = &H20     ' Retrieves the type as lpServiceClassId.
Public Const LUP_RETURN_VERSION = &H40  ' Retrieves the version as lpVersion.
Public Const LUP_RETURN_COMMENT = &H80  ' Retrieves the comment as lpszComment.
Public Const LUP_RETURN_ADDR = &H100    ' Retrieves the addresses as lpcsaBuffer.
Public Const LUP_RETURN_BLOB = &H200    ' Retrieves the private data as lpBlob.
Public Const LUP_RETURN_ALL = &HFF0     ' Retrieves all of the information.

' Constants - WSASetService.essOperation
Public Enum WSAESETSERVICEOP
  RNRSERVICE_REGISTER = 0 ' Register the service. For SAP, this means sending out a periodic broadcast. This is an NOP for the DNS name space. For persistent data stores, this means updating the address information.
  RNRSERVICE_DEREGISTER   ' Remove the service from the registry. For SAP, this means stop sending out the periodic broadcast. This is an NOP for the DNS name space. For persistent data stores this means deleting address information.
  RNRSERVICE_DELETE       ' Delete the service from dynamic name and persistent spaces. For services represented by multiple CSADDR_INFO structures (using the SERVICE_MULTIPLE flag), only the supplied address will be deleted, and this must match exactly the corresponding CSADDR_INFO structure that was supplied when the service was registered.
End Enum

' Constants - WSASetService.dwControlFlags
Public Const SERVICE_MULTIPLE = &H1 ' Controls scope of operation. When clear, service addresses are managed as a group. A register or removal from the registry invalidates all existing addresses before adding the given address set. When set, the action is only performed on the given address set. A register does not invalidate existing addresses and a removal from the registry only invalidates the given set of addresses.

' Constants - WSASocket.dwFlags
Public Const WSA_FLAG_OVERLAPPED = &H1  'This flag causes an overlapped socket to be created. Overlapped sockets can utilize WSASend, WSASendTo, WSARecv, WSARecvFrom, and WSAIoctl for overlapped I/O operations, which allow multiple operations to be initiated and in progress simultaneously. All functions that allow overlapped operation (WSASend, WSARecv, WSASendTo, WSARecvFrom, WSAIoctl) also support nonoverlapped usage on an overlapped socket if the values for parameters related to overlapped operations are NULL.
Public Const WSA_FLAG_MULTIPOINT_C_ROOT = &H2  ' Indicates that the socket created will be a c_root in a multipoint session. Only allowed if a rooted control plane is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to Multipoint and Multicast Semantics for additional information.
Public Const WSA_FLAG_MULTIPOINT_C_LEAF = &H4  'Indicates that the socket created will be a c_leaf in a multicast session. Only allowed if XP1_SUPPORT_MULTIPOINT is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to Multipoint and Multicast Semantics for additional information.
Public Const WSA_FLAG_MULTIPOINT_D_ROOT = &H8  'Indicates that the socket created will be a d_root in a multipoint session. Only allowed if a rooted data plane is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to Multipoint and Multicast Semantics for additional information.
Public Const WSA_FLAG_MULTIPOINT_D_LEAF = &H10 'Indicates that the socket created will be a d_leaf in a multipoint session. Only allowed if XP1_SUPPORT_MULTIPOINT is indicated in the protocol's WSAPROTOCOL_INFO structure. Refer to Multipoint and Multicast Semantics for additional information.

' Constants - WSAWaitForMultipleEvents (Return)
Public Const WSA_WAIT_EVENT_0 = &H0        ' WSA_WAIT_EVENT_0 = WAIT_OBJECT_0 = STATUS_WAIT_0 = 0x00000000
Public Const WSA_WAIT_IO_COMPLETION = &HC0 ' WSA_WAIT_IO_COMPLETION = WAIT_IO_COMPLETION = STATUS_USER_APC = 0x000000C0
Public Const WSA_WAIT_TIMEOUT = &H102      ' WSA_WAIT_TIMEOUT = WAIT_TIMEOUT = 258

' Constants Windows Sockets base error number
Public Const WSABASEERR = 10000

' Constants - Windows Sockets definitions of regular Microsoft C error constants
Public Const WSAEINTR = (WSABASEERR + 4)    ' Interrupted system call
Public Const WSAEBADF = (WSABASEERR + 9)    ' Bad file number
Public Const WSEACCES = (WSABASEERR + 13)   ' Permission denied
Public Const WSAEFAULT = (WSABASEERR + 14)  ' Bad address
Public Const WSAEINVAL = (WSABASEERR + 22)  ' Invalid argument
Public Const WSAEMFILE = (WSABASEERR + 24)  ' Too many open files

' Constants - Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK = (WSABASEERR + 35)      ' A non-blocking socket operation could not be completed immediately
Public Const WSAEINPROGRESS = (WSABASEERR + 36)      ' A blocking operation is currently executing
Public Const WSAEALREADY = (WSABASEERR + 37)         ' An operation was attempted on a non-blocking socket that already had an operation in progress
Public Const WSAENOTSOCK = (WSABASEERR + 38)         ' An operation was attempted on something that is not a socket
Public Const WSAEDESTADDRREQ = (WSABASEERR + 39)     ' A required address was omitted from an operation on a socket
Public Const WSAEMSGSIZE = (WSABASEERR + 40)         ' A message sent on a datagram socket was larger than the internal message buffer or some other network limit, or the buffer used to receive a datagram into was smaller than the datagram itself
Public Const WSAEPROTOTYPE = (WSABASEERR + 41)       ' A protocol was specified in the socket function call that does not support the semantics of the socket type requested
Public Const WSAENOPROTOOPT = (WSABASEERR + 42)      ' An unknown, invalid, or unsupported option or level was specified in a getsockopt or setsockopt call
Public Const WSAEPROTONOSUPPORT = (WSABASEERR + 43)  ' The requested protocol has not been configured into the system, or no implementation for it exists
Public Const WSAESOCKTNOSUPPORT = (WSABASEERR + 44)  ' The support for the specified socket type does not exist in this address family
Public Const WSAEOPNOTSUPP = (WSABASEERR + 45)       ' The attempted operation is not supported for the type of object referenced
Public Const WSAEPFNOSUPPORT = (WSABASEERR + 46)     ' The protocol family has not been configured into the system or no implementation for it exists
Public Const WSAEAFNOSUPPORT = (WSABASEERR + 47)     ' An address incompatible with the requested protocol was used
Public Const WSAEADDRINUSE = (WSABASEERR + 48)       ' Address already in use - Only one usage of each socket address (protocol/network address/port) is normally permitted
Public Const WSAEADDRNOTAVAIL = (WSABASEERR + 49)    ' The requested address is not valid in its context
Public Const WSAENETDOWN = (WSABASEERR + 50)         ' Network is down - This error may be reported at any time if the Windows Sockets implementation detects an underlying failure
Public Const WSAENETUNREACH = (WSABASEERR + 51)      ' Network is unreachable - A socket operation encountered a dead network
Public Const WSAENETRESET = (WSABASEERR + 52)        ' The connection has been broken due to keep-alive activity detecting a failure while the operation was in progress
Public Const WSAECONNABORTED = (WSABASEERR + 53)     ' An established connection was aborted by the software in your host machine
Public Const WSAECONNRESET = (WSABASEERR + 54)       ' An existing connection was forcibly closed by the remote host
Public Const WSAENOBUFS = (WSABASEERR + 55)          ' An operation on a socket could not be performed because the system lacked sufficient buffer space or because a queue was full
Public Const WSAEISCONN = (WSABASEERR + 56)          ' A connect request was made on an already connected socket
Public Const WSAENOTCONN = (WSABASEERR + 57)         ' Socket is not connected - A request to send or receive data was disallowed because the socket is not connected and (when sending on a datagram socket using a sendto call) no address was supplied
Public Const WSAESHUTDOWN = (WSABASEERR + 58)        ' A request to send or receive data was disallowed because the socket had already been shut down in that direction with a previous shutdown call
Public Const WSAETOOMANYREFS = (WSABASEERR + 59)     ' Too many references to some kernel object
Public Const WSAETIMEDOUT = (WSABASEERR + 60)        ' A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond
Public Const WSAECONNREFUSED = (WSABASEERR + 61)     ' No connection could be made because the target machine actively refused it
Public Const WSAELOOP = (WSABASEERR + 62)            ' Too many levels of symbolic links - Cannot translate name
Public Const WSAENAMETOOLONG = (WSABASEERR + 63)     ' Name component or name was too long
Public Const WSAEHOSTDOWN = (WSABASEERR + 64)        ' A socket operation failed because the destination host was down
Public Const WSAEHOSTUNREACH = (WSABASEERR + 65)     ' A socket operation was attempted to an unreachable host
Public Const WSAENOTEMPTY = (WSABASEERR + 66)        ' Cannot remove a directory that is not empty
Public Const WSAEPROCLIM = (WSABASEERR + 67)         ' A Windows Sockets implementation may have a limit on the number of applications that may use it simultaneously
Public Const WSAEUSERS = (WSABASEERR + 68)           ' Ran out of quota
Public Const WSAEDQUOT = (WSABASEERR + 69)           ' Ran out of disk quota
Public Const WSAESTALE = (WSABASEERR + 70)           ' File handle reference is no longer available
Public Const WSAEREMOTE = (WSABASEERR + 71)          ' Item is not available locally

' Constants - Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY = (WSABASEERR + 91)           ' WSAStartup cannot function at this time because the underlying system it uses to provide network services is currently unavailable
Public Const WSAVERNOTSUPPORTED = (WSABASEERR + 92)       ' The Windows Sockets version requested is not supported
Public Const WSANOTINITIALISED = (WSABASEERR + 93)        ' Either the application has not called WSAStartup, or WSAStartup failed
Public Const WSAEDISCON = (WSABASEERR + 101)              ' Disconnect
Public Const WSAENOMORE = (WSABASEERR + 102)              ' No more results can be returned by WSALookupServiceNext
Public Const WSAECANCELLED = (WSABASEERR + 103)           ' A call to WSALookupServiceEnd was made while this call was still processing - The call has been canceled
Public Const WSAEINVALIDPROCTABLE = (WSABASEERR + 104)    ' The procedure call table is invalid
Public Const WSAEINVALIDPROVIDER = (WSABASEERR + 105)     ' The requested service provider is invalid
Public Const WSAEPROVIDERFAILEDINIT = (WSABASEERR + 106)  ' The requested service provider could not be loaded or initialized
Public Const WSASYSCALLFAILURE = (WSABASEERR + 107)       ' A system call that should never fail has failed
Public Const WSASERVICE_NOT_FOUND = (WSABASEERR + 108)    ' No such service is known - The service cannot be found in the specified name space
Public Const WSATYPE_NOT_FOUND = (WSABASEERR + 109)       ' The specified class was not found
Public Const WSA_E_NO_MORE = (WSABASEERR + 110)           ' No more results can be returned by WSALookupServiceNext
Public Const WSA_E_CANCELLED = (WSABASEERR + 111)         ' A call to WSALookupServiceEnd was made while this call was still processing - The call has been canceled
Public Const WSAEREFUSED = (WSABASEERR + 112)             ' A database query failed because it was actively refused

' Constants - Error return codes from GetHostByName() and GetHostByAddr() (when using the resolver)
Public Const WSAHOST_NOT_FOUND = (WSABASEERR + 1001)  ' Host not found - This message indicates that the key (name, address, and so on) was not found
Public Const WSATRY_AGAIN = (WSABASEERR + 1002)       ' Nonauthoritative host not found - This error may suggest that the name service itself is not functioning
Public Const WSANO_RECOVERY = (WSABASEERR + 1003)     ' Nonrecoverable error - This error may suggest that the name service itself is not functioning
Public Const WSANO_DATA = (WSABASEERR + 1004)         ' Valid name, no data record of requested type - This error indicates that the key (name, address, and so on) was not found

' Constants - Define QOS related error return codes
Public Const WSA_QOS_RECEIVERS = (WSABASEERR + 1005)           ' At least one Reserve has arrived
Public Const WSA_QOS_SENDERS = (WSABASEERR + 1006)             ' At least one Path has arrived
Public Const WSA_QOS_NO_SENDERS = (WSABASEERR + 1007)          ' There are no senders
Public Const WSA_QOS_NO_RECEIVERS = (WSABASEERR + 1008)        ' There are no receivers
Public Const WSA_QOS_REQUEST_CONFIRMED = (WSABASEERR + 1009)   ' Reserve has been confirmed
Public Const WSA_QOS_ADMISSION_FAILURE = (WSABASEERR + 1010)   ' Error due to lack of resources
Public Const WSA_QOS_POLICY_FAILURE = (WSABASEERR + 1011)      ' Rejected for administrative reasons - bad credentials
Public Const WSA_QOS_BAD_STYLE = (WSABASEERR + 1012)           ' Unknown or conflicting style
Public Const WSA_QOS_BAD_OBJECT = (WSABASEERR + 1013)          ' Problem with some part of the filterspec or providerspecific buffer in general
Public Const WSA_QOS_TRAFFIC_CTRL_ERROR = (WSABASEERR + 1014)  ' Problem with some part of the flowspec
Public Const WSA_QOS_GENERIC_ERROR = (WSABASEERR + 1015)       ' General QOS error
Public Const WSA_QOS_ESERVICETYPE = (WSABASEERR + 1016)        ' An invalid or unrecognized service type was found in the flowspec
Public Const WSA_QOS_EFLOWSPEC = (WSABASEERR + 1017)           ' An invalid or inconsistent flowspec was found in the QOS structure
Public Const WSA_QOS_EPROVSPECBUF = (WSABASEERR + 1018)        ' Invalid QOS provider-specific buffer
Public Const WSA_QOS_EFILTERSTYLE = (WSABASEERR + 1019)        ' An invalid QOS filter style was used
Public Const WSA_QOS_EFILTERTYPE = (WSABASEERR + 1020)         ' An invalid QOS filter type was used
Public Const WSA_QOS_EFILTERCOUNT = (WSABASEERR + 1021)        ' An incorrect number of QOS FILTERSPECs were specified in the FLOWDESCRIPTOR
Public Const WSA_QOS_EOBJLENGTH = (WSABASEERR + 1022)          ' An object with an invalid ObjectLength field was specified in the QOS provider-specific buffer
Public Const WSA_QOS_EFLOWCOUNT = (WSABASEERR + 1023)          ' An incorrect number of flow descriptors was specified in the QOS structure
Public Const WSA_QOS_EUNKOWNPSOBJ = (WSABASEERR + 1024)        ' An unrecognized object was found in the QOS provider-specific buffer
Public Const WSA_QOS_EPOLICYOBJ = (WSABASEERR + 1025)          ' An invalid policy object was found in the QOS provider-specific buffer
Public Const WSA_QOS_EFLOWDESC = (WSABASEERR + 1026)           ' An invalid QOS flow descriptor was found in the flow descriptor list
Public Const WSA_QOS_EPSFLOWSPEC = (WSABASEERR + 1027)         ' An invalid or inconsistent flowspec was found in the QOS provider-specific buffer
Public Const WSA_QOS_EPSFILTERSPEC = (WSABASEERR + 1028)       ' An invalid FILTERSPEC was found in the QOS provider-specific buffer
Public Const WSA_QOS_ESDMODEOBJ = (WSABASEERR + 1029)          ' An invalid shape discard mode object was found in the QOS provider-specific buffer
Public Const WSA_QOS_ESHAPERATEOBJ = (WSABASEERR + 1030)       ' An invalid shaping rate object was found in the QOS provider-specific buffer
Public Const WSA_QOS_RESERVED_PETYPE = (WSABASEERR + 1031)     ' A reserved policy element was found in the QOS provider-specific buffer

' Constants - Other Related Errors
Public Const SOCKET_ERROR = (-1)         ' Socket Error
Public Const WSA_INVALID_HANDLE = 1609   ' Specified event object handle is invalid.  (An application attempts to use an event object, but the specified handle is not valid)
Public Const WSA_INVALID_PARAMETER = 87  ' One or more parameters are invalid.        (An application used a Windows Sockets function which directly maps to a Win32 function. The Win32 function is indicating a problem with one or more parameters)
Public Const WSA_IO_PENDING = 997        ' Overlapped operations will complete later. (The application has initiated an overlapped operation that cannot be completed immediately. A completion indication will be given later when the operation has been completed)
Public Const WSA_NOT_ENOUGH_MEMORY = 8   ' Insufficient memory available.             (An application used a Windows Sockets function that directly maps to a Win32 function. The Win32 function is indicating a lack of required memory resources)
Public Const WSA_OPERATION_ABORTED = 995 ' Overlapped operation aborted.              (An overlapped operation was canceled due to the closure of the socket, or the execution of the SIO_FLUSH command in WSAIoctl)

'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

' Types - Win32 Related
Public Type POINTAPI
  X As Long 'LONG // X axis coordinate
  Y As Long 'LONG // Y axis coordinate
End Type

Public Type MSG 'The MSG structure contains message information from a thread's message queue.
  hWnd    As Long     'HWND   // Handle to the window whose window procedure receives the message.
  Message As Long     'UINT   // Specifies the message identifier. Applications can only use the low word; the high word is reserved by the system.
  wParam  As Long     'WPARAM // Specifies additional information about the message. The exact meaning depends on the value of the message member.
  lParam  As Long     'LPARAM // Specifies additional information about the message. The exact meaning depends on the value of the message member.
  Time    As Long     'DWORD  // Specifies the time at which the message was posted.
  PT      As POINTAPI 'POINT  // Specifies the cursor position, in screen coordinates, when the message was posted.
End Type

' Constants - Win32 Related - PeekMessage.wRemoveMsg
Public Const PM_REMOVE = &H1   ' Messages are removed from the queue after processing by PeekMessage.
Public Const PM_NOREMOVE = &H0 ' Messages are not removed from the queue after processing by PeekMessage.

' Constants - Win32 Related - FormatMessage.dwFlags
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100 ' Specifies that the lpBuffer parameter is a pointer to a PVOID pointer, and that the nSize parameter specifies the minimum number of TCHARs to allocate for an output message buffer. The function allocates a buffer large enough to hold the formatted message, and places a pointer to the allocated buffer at the address specified by lpBuffer. The caller should use the LocalFree function to free the buffer when it is no longer needed.
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200  ' Specifies that insert sequences in the message definition are to be ignored and passed through to the output buffer unchanged. This flag is useful for fetching a message for later formatting. If this flag is set, the Arguments parameter is ignored.
Public Const FORMAT_MESSAGE_FROM_STRING = &H400     ' Specifies that lpSource is a pointer to a null-terminated message definition. The message definition may contain insert sequences, just as the message text in a message table resource may. Cannot be used with FORMAT_MESSAGE_FROM_HMODULE or FORMAT_MESSAGE_FROM_SYSTEM.
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800    ' Specifies that lpSource is a module handle containing the message-table resource(s) to search. If this lpSource handle is NULL, the current process's application image file will be searched. Cannot be used with FORMAT_MESSAGE_FROM_STRING.
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000    ' Specifies that the function should search the system message-table resource(s) for the requested message. If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system message table if the message is not found in the module specified by lpSource. Cannot be used with FORMAT_MESSAGE_FROM_STRING.  If this flag is specified, an application can pass the result of the GetLastError function to retrieve the message text for a system-defined error.
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000 ' Specifies that the Arguments parameter is not a va_list structure, but instead is just a pointer to an array of values that represent the arguments.

' General Win32 API Declarations
Public Declare Sub SetLastError Lib "KERNEL32" (ByVal dwErrCode As Long)
Public Declare Function DispatchMessage Lib "USER32" Alias "DispatchMessageA" (ByRef lpMSG As MSG) As Long
Public Declare Function FormatMessage Lib "KERNEL32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Public Declare Function GetLastError Lib "KERNEL32" () As Long
Public Declare Function PeekMessage Lib "USER32" Alias "PeekMessageA" (ByRef lpMSG As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Public Declare Function StringFromPointer Lib "KERNEL32" Alias "lstrcpyA" (ByVal Return_String As String, ByVal StringPointer As Long) As Long
Public Declare Function TranslateMessage Lib "USER32" (ByRef lpMSG As MSG) As Long
'_____________________________________________________________________________________________________________
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX





'=============================================================================================================
' accept
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets accept function permits an incoming connection attempt on a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in]  Descriptor identifying a socket that has been placed in a listening state with the listen function. The connection is actually made with the socket that is returned by accept.
' addr    [out] Optional pointer to a buffer that receives the address of the connecting entity, as known to the communications layer. The exact format of the addr parameter is determined by the address family that was established when the socket was created.
' addrlen [out] Optional pointer to an integer that contains the length of addr.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, accept returns a value of type SOCKET that is a descriptor for the new socket. This returned value is a handle for the socket on which the actual connection is made.
' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
' The integer referred to by addrlen initially contains the amount of space pointed to by addr. On return it will contain the actual length in bytes of the address returned.
' ____________________________________________________________________________________________________________
' SOCKET accept (SOCKET s, struct sockaddr FAR *addr, int FAR *addrlen);
'=============================================================================================================
Public Declare Function accept Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef SocketAddress As SOCKADDR, ByRef AddrLen As Long) As Long


'=============================================================================================================
' AcceptEx
'
' Minimum Availability : Windows Sockets 1.1 or later (not supported on Windows 95/98)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets AcceptEx function accepts a new connection, returns the local and remote address, and
' receives the first block of data sent by the client application.
'
' Note : This function is a Microsoft-specific extension to the Windows Sockets specification. For more
' information, see Microsoft Extensions and Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' sListenSocket         [in]  Descriptor identifying a socket that has already been called with the listen function. A server application waits for attempts to connect on this socket.
' sAcceptSocket         [in]  Descriptor identifying a socket on which to accept an incoming connection. This socket must not be bound or connected.
' lpOutputBuffer        [in]  Pointer to a buffer that receives the first block of data sent on a new connection, the local address of the server, and the remote address of the client. The receive data is written to the first part of the buffer starting at offset zero, while the addresses are written to the latter part of the buffer. If this parameter is set to NULL, no receive will be performed, nor will local or remote addresses be available through the use of GetAcceptExSockaddrs function calls.
' dwReceiveDataLength   [in]  Number of bytes in lpOutputBuffer that will be used for actual receive data at the beginning of the buffer. This size should not include the size of the local address of the server, nor the remote address of the client; they are appended to the output buffer. If dwReceiveDataLength is zero, accepting the connection will not result in a receive operation. Instead, AcceptEx completes as soon as a connection arrives, without waiting for any data.
' dwLocalAddressLength  [in]  Number of bytes reserved for the local address information. This value must be at least 16 bytes more than the maximum address length for the transport protocol in use.
' dwRemoteAddressLength [in]  Number of bytes reserved for the remote address information. This value must be at least 16 bytes more than the maximum address length for the transport protocol in use.
' lpdwBytesReceived     [out] Pointer to a DWORD that receives the count of bytes received. This parameter is set only if the operation completes synchronously. If it returns ERROR_IO_PENDING and is completed later, then this DWORD is never set and you must obtain the number of bytes read from the completion notification mechanism.
' lpOverlapped          [in]  An OVERLAPPED structure that is used to process the request. This parameter must be specified; it cannot be null.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, the AcceptEx function completed successfully and a value of TRUE is returned.
' If the function fails, AcceptEx returns FALSE. The WSAGetLastError function can then be called to
' return extended error information. If WSAGetLastError returns ERROR_IO_PENDING, then the operation
' was successfully initiated and is still in progress.
' ____________________________________________________________________________________________________________
' BOOL AcceptEx (SOCKET sListenSocket, SOCKET sAcceptSocket, PVOID lpOutputBuffer, DWORD dwReceiveDataLength, DWORD dwLocalAddressLength, DWORD dwRemoteAddressLength, LPDWORD lpdwBytesReceived, lpOverlapped lpOverlapped);
'=============================================================================================================
Public Declare Function AcceptEx Lib "WSOCK32.DLL" (ByVal sListenSocket As Long, ByVal sAcceptSocket As Long, ByRef lpOutputBuffer As Any, ByVal dwReceiveDataLength As Long, ByVal dwLocalAddressLength As Long, ByVal dwRemoteAddressLength As Long, ByRef lpdwBytesReceived As Long, ByRef lpOverlapped As OVERLAPPED) As Long   'BOOL


'=============================================================================================================
' bind
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets bind function associates a local address with a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in] Descriptor identifying an unbound socket.
' name    [in] Address to assign to the socket from the SOCKADDR structure.
' NameLen [in] Length of the value in the name parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, bind returns zero. Otherwise, it returns SOCKET_ERROR, and a specific error code can be
' retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int bind (SOCKET s, const struct sockaddr FAR *name, int namelen);
'=============================================================================================================
Public Declare Function bind Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Name As SOCKADDR, ByVal NameLen As Long) As Long


'=============================================================================================================
' closesocket
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets closesocket function closes an existing socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s [in] Descriptor identifying the socket to close.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, closesocket returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int closesocket (SOCKET s);
'=============================================================================================================
Public Declare Function closesocket Lib "WSOCK32.DLL" (ByVal hSocket As Long) As Long


'=============================================================================================================
' connect
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets connect function establishes a connection to a specified socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in] Descriptor identifying an unconnected socket.
' name    [in] Name of the socket to which the connection should be established.
' NameLen [in] Length of name.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, connect returns zero. Otherwise, it returns SOCKET_ERROR, and a specific error code
' can be retrieved by calling WSAGetLastError.
'
' On a blocking socket, the return value indicates success or failure of the connection attempt.
'
' With a nonblocking socket, the connection attempt cannot be completed immediately. In this case, connect
' will return SOCKET_ERROR, and WSAGetLastError will return WSAEWOULDBLOCK. In this case, there are three
' possible scenarios:
'   - Use the select function to determine the completion of the connection request by checking to see if
'     the socket is writeable.
'   - If the application is using WSAAsyncSelect to indicate interest in connection events, then the
'     application will receive an FD_CONNECT notification indicating that the connect operation is
'     complete (successfully or not).
'   - If the application is using WSAEventSelect to indicate interest in connection events, then the
'     associated event object will be signaled indicating that the connect operation is complete
'     (successfully or not).
' Until the connection attempt completes on a nonblocking socket, all subsequent calls to connect on the
' same socket will fail with the error code WSAEALREADY, and WSAEISCONN when the connection completes
' successfully. Due to ambiguities in version 1.1 of the Windows Sockets specification, error codes returned
' from connect while a connection is already pending may vary among implementations. As a result, it is not
' recommended that applications use multiple calls to connect to detect connection completion. If they do,
' they must be prepared to handle WSAEINVAL and WSAEWOULDBLOCK error values the same way that they handle
' WSAEALREADY, to assure robust execution.
' If the error code returned indicates the connection attempt failed (that is, WSAECONNREFUSED,
' SAENETUNREACH, WSAETIMEDOUT) the application can call connect again for the same socket.
' ____________________________________________________________________________________________________________
' int connect (SOCKET s, const struct sockaddr FAR *name, int namelen);
'=============================================================================================================
Public Declare Function connect Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Name As SOCKADDR, ByVal NameLen As Long) As Long


'=============================================================================================================
' EnumProtocols
'
' Minimum Availability : Windows Sockets 1.1 or later (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The EnumProtocols function obtains information about a specified set of network protocols that are active
' on a local host.
'
' Important:
' ¯¯¯¯¯¯¯¯¯¯
' The EnumProtocols function is a Microsoft-specific extension to the Windows Sockets 1.1 specification.
' This function is obsolete. For the convenience of Windows Sockets 1.1 developers, the reference material
' is included.  The WSAEnumProtocols function provides equivalent functionality in Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpiProtocols     [in]     Pointer to a null-terminated array of protocol identifiers. The EnumProtocols function obtains information about the protocols specified by this array.  If lpiProtocols is NULL, the function obtains information about all available protocols.  The following protocol identifier values are defined: IPPROTO_TCP, IPPROTO_UDP, ISOPROTO_TP4, NSPROTO_IPX, NSPROTO_SPX, NSPROTO_SPXII,
' lpProtocolBuffer [out]    Pointer to a buffer that the function fills with an array of PROTOCOL_INFO data structures.
' lpdwBufferLength [in,out] Pointer to a variable that, on input, specifies the size, in bytes, of the buffer pointed to by lpProtocolBuffer.  On output, the function sets this variable to the minimum buffer size needed to retrieve all of the requested information. For the function to succeed, the buffer must be at least this size.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is the number of PROTOCOL_INFO data structures written to
' the buffer pointed to by lpProtocolBuffer.  If the function fails, the return value is SOCKET_ERROR (–1).
' To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' INT EnumProtocols (LPINT lpiProtocols, LPVOID lpProtocolBuffer, LPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function EnumProtocols Lib "WSOCK32.DLL" Alias "EnumProtocolsA" (ByVal lpiProtocols As Long, ByRef lpProtocolBuffer As PROTOCOL_INFO, ByRef lpdwBufferLength As Long) As Long


'=============================================================================================================
' GetAcceptExSockaddrs
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets GetAcceptExSockaddrs function parses the data obtained from a call to the AcceptEx
' function and passes the local and remote addresses to a SOCKADDR structure.
'
' Note : This function is a Microsoft-specific extension to the Windows Sockets specification. For more
' information, see Microsoft Extensions and Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpOutputBuffer        [in]  Pointer to a buffer that receives the first block of data sent on a connection resulting from an AcceptEx call. Must be the same lpOutputBuffer parameter that was passed to the AcceptEx function.
' dwReceiveDataLength   [in]  Number of bytes in the buffer used for receiving the first data. This value must be equal to the dwReceiveDataLength parameter that was passed to the AcceptEx function.
' dwLocalAddressLength  [in]  Number of bytes reserved for the local address information. Must be equal to the dwLocalAddressLength parameter that was passed to the AcceptEx function.
' dwRemoteAddressLength [in]  Number of bytes reserved for the remote address information. This value must be equal to the dwRemoteAddressLength parameter that was passed to the AcceptEx function.
' LocalSockaddr         [out] Pointer to the SOCKADDR structure that receives the local address of the connection (the same information that would be returned by the Windows Sockets getsockname function). This parameter must be specified.
' LocalSockaddrLength   [out] Size of the local address. This parameter must be specified.
' RemoteSockaddr        [out] Pointer to the SOCKADDR structure that receives the remote address of the connection (the same information that would be returned by the Windows Sockets getpeername function). This parameter must be specified.
' RemoteSockaddrLength  [out] Size of the local address. This parameter must be specified.
'
' Return:
' ¯¯¯¯¯¯¯
' ( None )
' ____________________________________________________________________________________________________________
' VOID GetAcceptExSockaddrs (PVOID lpOutputBuffer, DWORD dwReceiveDataLength, DWORD dwLocalAddressLength, DWORD dwRemoteAddressLength, LPSOCKADDR *LocalSockaddr, LPINT LocalSockaddrLength, LPSOCKADDR *RemoteSockaddr, LPINT RemoteSockaddrLength);
'=============================================================================================================
Public Declare Sub GetAcceptExSockaddrs Lib "WSOCK32.DLL" (ByRef lpOutputBuffer As Any, ByVal dwReceiveDataLength As Long, ByVal dwLocalAddressLength As Long, ByVal dwRemoteAddressLength As Long, ByRef LocalSockaddr As SOCKADDR, ByRef LocalSockaddrLength As Long, ByRef RemoteSockaddr As SOCKADDR, ByRef RemoteSockaddrLength As Long)


'=============================================================================================================
' GetAddressByName
'
' Minimum Availability : Requires Windows Sockets 1.1 (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The GetAddressByName function queries a name space, or a set of default name spaces, in order to obtain
' network address information for a specified network service. This process is known as service name resolution.
' A network service can also use the function to obtain local address information that it can use with the bind
' function.
'
' The functions detailed in Protocol-Independent Name Resolution provide equivalent functionality in Windows
' Sockets 2.
'
' Important : The GetAddressByName function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers, the
' reference material is as follows.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwNameSpace           [in]     Specifies the name space, or a set of default name spaces, that the operating system will query for network address information.  Use one of the following constants to specify a name space: NS_DEFAULT, NS_DNS, NS_NETBT, NS_SAP, NS_TCPIP_HOSTS, NS_TCPIP_LOCAL
'                                 Most calls to GetAddressByName should use the special value NS_DEFAULT. This lets a client get by with no knowledge of which name spaces are available on an internetwork. The system administrator determines name space access. Name spaces can come and go without the client having to be aware of the changes.
' lpServiceType         [in]     Pointer to a globally unique identifier (GUID) that specifies the type of the network service. The header file Svcguid.h includes definitions of several GUID service types, and macros for working with them.
' lpServiceName         [in]     Pointer to a zero-terminated string that uniquely represents the service name. For example, "MY SNA SERVER".  Setting lpServiceName to NULL is the equivalent of setting dwResolution to RES_SERVICE. The function operates in its second mode, obtaining the local address to which a service of the specified type should bind. The function stores the local address within the LocalAddr member of the CSADDR_INFO structures stored into *lpCsaddrBuffer.
'                                 If dwResolution is set to RES_SERVICE, the function ignores the lpServiceName parameter.
'                                 If dwNameSpace is set to NS_DNS, *lpServiceName is the name of the host.
' lpiProtocols          [in]     Pointer to a zero-terminated array of protocol identifiers. The function restricts a name resolution attempt to name space providers that offer these protocols. This lets the caller limit the scope of the search. If lpiProtocols is NULL, the function obtains information on all available protocols.
' dwResolution          [in]     Set of bit flags that specify aspects of the service name resolution process. The following bit flags are defined: RES_SERVICE, RES_FIND_MULTIPLE, RES_SOFT_SEARCH
' lpServiceAsyncInfo    [in]     Reserved for future use; must be set to NULL.
' lpCsaddrBuffer        [out]    Pointer to a buffer to receive one or more CSADDR_INFO data structures. The number of structures written to the buffer depends on the amount of information found in the resolution attempt. You should assume that multiple structures will be written, although in many cases there will only be one.
' lpdwBufferLength      [in,out] Pointer to a variable that, upon input, specifies the size, in bytes, of the buffer pointed to by lpCsaddrBuffer.  Upon output, this variable contains the total number of bytes required to store the array of CSADDR_INFO structures. If this value is less than or equal to the input value of *lpdwBufferLength, and the function is successful, this is the number of bytes actually stored in the buffer. If this value is greater than the input value of *lpdwBufferLength, the buffer was too small, and the output value of *lpdwBufferLength is the minimal required buffer size.
' lpAliasBuffer         [out]    Pointer to a buffer to receive alias information for the network service. If a name space supports aliases, the function stores an array of zero-terminated name strings into the buffer pointed to by lpAliasBuffer. There is a double zero-terminator at the end of the list. The first name in the array is the service's primary name. Names that follow are aliases. An example of a name space that supports aliases is DNS.  If a name space does not support aliases, it stores a double zero-terminator into the buffer.  This parameter is optional, and can be set to NULL.
' lpdwAliasBufferLength [in,out] Pointer to a variable that, upon input, specifies the size, in bytes, of the buffer pointed to by lpAliasBuffer.  Upon output, this variable contains the total number of bytes required to store the array of name strings. If this value is less than or equal to the input value of *lpdwAliasBufferLength, and the function is successful, this is the number of bytes actually stored in the buffer. If this value is greater than the input value of *lpdwAliasBufferLength, the buffer was too small, and the output value of *lpdwAliasBufferLength is the minimal required buffer size.  If lpAliasBuffer is NULL, lpdwAliasBufferLength is meaningless and can also be NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is the number of CSADDR_INFO data structures written to the
' buffer pointed to by lpCsaddrBuffer.  If the function fails, the return value is SOCKET_ERROR( – 1).
' To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' INT GetAddressByName (DWORD dwNameSpace, LPGUID lpServiceType, LPTSTR lpServiceName, LPINT lpiProtocols, DWORD dwResolution, LPSERVICE_ASYNC_INFO lpServiceAsyncInfo, LPVOID lpCsaddrBuffer, LPDWORD lpdwBufferLength, LPTSTR lpAliasBuffer, LPDWORD lpdwAliasBufferLength);
'=============================================================================================================
Public Declare Function GetAddressByName Lib "WSOCK32.DLL" Alias "GetAddressByNameA" (ByVal dwNameSpace As Long, ByRef lpServiceType As GUID, ByVal lpServiceName As String, ByRef lpiProtocols As String, ByVal dwResolution As Long, ByVal Reserved As Long, ByRef lpCsaddrBuffer As CSADDR_INFO, ByRef lpdwBufferLength As Long, ByVal lpAliasBuffer As String, ByRef lpdwAliasBufferLength As Long) As Long


'=============================================================================================================
' gethostbyaddr
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets gethostbyaddr function retrieves the host information corresponding to a network address.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' addr [in] Pointer to an address in network byte order.
' len  [in] Length of the address.
' type [in] Type of the address, such as the AF_INET address family type (defined as TCP, UDP, and other associated Internet protocols). Address family types and their corresponding values are defined in the winsock2.h header file.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, gethostbyaddr returns a pointer to the HOSTENT structure. Otherwise, it returns a
' NULL pointer, and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct HOSTENT FAR * gethostbyaddr (const char FAR *addr, int len, int type);
'=============================================================================================================
Public Declare Function gethostbyaddr Lib "WSOCK32.DLL" (ByRef Address As Long, ByVal AddrLen As Long, ByVal AddrType As Long) As Long 'HOSTENT


'=============================================================================================================
' gethostbyname
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets gethostbyname function retrieves host information corresponding to a host name from
' a host database.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' name [out] Pointer to the null-terminated name of the host to resolve.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, gethostbyname returns a pointer to the HOSTENT structure described above.
' Otherwise, it returns a NULL pointer and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct HOSTENT FAR *gethostbyname (const char FAR *name);
'=============================================================================================================
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal Name As String) As Long 'HOSTENT


'=============================================================================================================
' gethostname
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets gethostname function returns the standard host name for the local machine.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' name    [out] Pointer to a buffer that receives the local host name.
' NameLen [in]  Length of the buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, gethostname returns zero. Otherwise, it returns SOCKET_ERROR and a specific error code
' can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int gethostname (char FAR *name, int namelen);
'=============================================================================================================
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal HostName As String, ByVal NameLen As Long) As Long


'=============================================================================================================
' GetNameByType
'
' Minimum Availability : Windows Sockets 1.1 or later (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The GetNameByType function obtains the name of a network service. The network service is specified by its
' service type.
' Important : The GetNameByType function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers, the
' reference material is as follows.
' The functions detailed in Protocol-Independent Name Resolution provide equivalent functionality in
' Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpServiceType [in]     Pointer to a globally unique identifier (GUID) that specifies the type of the network service. The header file Svcguid.h includes definitions of several GUID service types, and macros for working with them.
' lpServiceName [out]    Pointer to a buffer to receive a zero-terminated string that uniquely represents the name of the network service.
' dwNameLength  [in,out] Pointer to a variable that, on input, specifies the size of the buffer pointed to by lpServiceName. On output, the variable contains the actual size of the service name string.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is not SOCKET_ERROR ( –1).
' If the function fails, the return value is SOCKET_ERROR ( –1). To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' INT GetNameByType (LPGUID lpServiceType, LPTSTR lpServiceName, DWORD dwNameLength);
'=============================================================================================================
Public Declare Function GetNameByType Lib "WSOCK32.DLL" Alias "GetNameByTypeA" (ByRef lpServiceType As GUID, ByVal lpServiceName As String, ByRef dwNameLength As Long) As Long


'=============================================================================================================
' getpeername
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getpeername function retrieves the name of the peer to which a socket is connected.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in]      Descriptor identifying a connected socket.
' name    [out]     The structure that receives the name of the peer.
' NameLen [in, out] Pointer to the size of the name structure.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getpeername returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int getpeername (SOCKET s, struct sockaddr FAR *name, int FAR *namelen);
'=============================================================================================================
Public Declare Function getpeername Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef PearName As SOCKADDR, ByRef NameLen As Long) As Long


'=============================================================================================================
' getprotobyname
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getprotobyname function retrieves the protocol information corresponding to a protocol
' name.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' name [in] Pointer to a null-terminated protocol name.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getprotobyname returns a pointer to the PROTOENT. Otherwise, it returns a NULL
' pointer and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct PROTOENT FAR * getprotobyname (const char FAR *name);
'=============================================================================================================
Public Declare Function getprotobyname Lib "WSOCK32.DLL" (ByVal Name As String) As Long 'As PROTOENT


'=============================================================================================================
' getprotobynumber
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getprotobynumber function retrieves protocol information corresponding to a protocol number.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' number [in] Protocol number, in host byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getprotobynumber returns a pointer to the PROTOENT structure. Otherwise, it returns a
' NULL pointer and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct PROTOENT FAR * getprotobynumber (int number);
'=============================================================================================================
Public Declare Function getprotobynumber Lib "WSOCK32.DLL" (ByVal Number As Long) As Long 'PROTOENT


'=============================================================================================================
' getservbyname
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getservbyname function retrieves service information corresponding to a service name
' and protocol.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' name  [in] Pointer to a null-terminated service name.
' proto [in] Optional pointer to a null-terminated protocol name. If this pointer is NULL, getservbyname returns the first service entry where name matches the s_name member of the SERVENT structure or the s_aliases member of the SERVENT structure. Otherwise, getservbyname matches both the name and the proto.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getservbyname returns a pointer to the SERVENT structure. Otherwise, it returns a NULL
' pointer and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct servent FAR * getservbyname (const char FAR *name, const char FAR *proto);
'=============================================================================================================
Public Declare Function getservbyname Lib "WSOCK32.DLL" (ByVal Name As String, ByVal Proto As String) As Long 'SERVENT


'=============================================================================================================
' getservbyport
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getservbyport function retrieves service information corresponding to a port and protocol.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' port  [in] Port for a service, in network byte order.
' Proto [in] Optional pointer to a protocol name. If this is NULL, getservbyport returns the first service entry for which the port matches the s_port of the SERVENT structure. Otherwise, getservbyport matches both the port and the proto parameters.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getservbyport returns a pointer to the SERVENT structure. Otherwise, it returns a NULL
' pointer and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' struct servent FAR * getservbyport (int port, const char FAR *proto);
'=============================================================================================================
Public Declare Function getservbyport Lib "WSOCK32.DLL" (ByVal Port As Long, ByVal Proto As String) As Long  'SERVENT


'=============================================================================================================
' GetService
'
' Minimum Availability : Windows Sockets 1.1 or later (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The GetService function obtains information about a network service in the context of a set of default
' name spaces or a specified name space. The network service is specified by its type and name. The
' information about the service is obtained as a set of NS_SERVICE_INFO data structures.
'
' The functions detailed in Protocol-Independent Name Resolution provide equivalent functionality in Windows
' Sockets 2.
'
' Important : The GetService function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers, the
' reference material is as follows.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwNameSpace        [in]     Specifies the name space, or a set of default name spaces, that the operating system queries for information about the specified network service. Use one of the following constants to specify a name space: NS_DEFAULT, NS_DNS, NS_NETBT, NS_SAP, NS_TCPIP_HOSTS, NS_TCPIP_LOCAL
'                             Most calls to GetService should use the special value NS_DEFAULT. This lets a client get by without knowing available name spaces on an internetwork. The system administrator determines name space access. Name spaces can come and go without the client having to be aware of the changes.
' lpGuid             [in]     Pointer to a globally unique identifier (GUID) that specifies the type of the network service. The header file Svcguid.h includes GUID service types from many well-known services within the DNS and SAP name spaces.
' lpServiceName      [in]     Pointer to a zero-terminated string that uniquely represents the service name. For example, "MY SNA SERVER."
' dwProperties       [in]     Set of bit flags that specify the service information that the function obtains. Each of these bit flag constants, other than PROP_ALL, corresponds to a particular member of the SERVICE_INFO data structure. If the flag is set, the function puts information into the corresponding member of the data structures stored in *lpBuffer. The following bit flags are defined: PROP_COMMENT, PROP_LOCALE, PROP_DISPLAY_HINT, PROP_VERSION, PROP_START_TIME, PROP_MACHINE, PROP_ADDRESSES, PROP_SD, PROP_ALL
' lpBuffer           [out]    Pointer to a buffer to receive an array of NS_SERVICE_INFO structures and associated service information. Each NS_SERVICE_INFO structure contains service information in the context of a particular name space.
'                             Note : If dwNameSpace is NS_DEFAULT, the function stores more than one structure into the buffer; otherwise, just one structure is stored.  Each NS_SERVICE_INFO structure contains a SERVICE_INFO structure. The members of these SERVICE_INFO structures will contain valid data based on the bit flags that are set in the dwProperties parameter. If a member's corresponding bit flag is not set in dwProperties, the member's value is undefined. The function stores the NS_SERVICE_INFO structures in a consecutive array, starting at the beginning of the buffer. The pointers in the contained SERVICE_INFO structures point to information that is stored in the buffer between the end of the NS_SERVICE_INFO structures and the end of the buffer.
' lpdwBufferSize     [in,out] Pointer to a variable that, on input, contains the size, in bytes, of the buffer pointed to by lpBuffer. On output, this variable contains the number of bytes required to store the requested information. If this output value is greater than the input value, the function has failed due to insufficient buffer size.
' lpServiceAsyncInfo [in]     Reserved for future use. Must be set to NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is the number of NS_SERVICE_INFO structures stored in *lpBuffer.
' Zero indicates that no structures were stored.  If the function fails, the return value is SOCKET_ERROR
' ( – 1). To get extended error information, call GetLastError.
' ____________________________________________________________________________________________________________
' INT GetService (DWORD dwNameSpace, PGUID lpGuid, LPTSTR lpServiceName, DWORD dwProperties, LPVOID lpBuffer, LPDWORD lpdwBufferSize, LPSERVICE_ASYNC_INFO lpServiceAsyncInfo);
'=============================================================================================================
Public Declare Function GetService Lib "WSOCK32.DLL" Alias "GetServiceA" (ByVal dwNameSpace As Long, ByRef lpGuid As GUID, ByVal lpServiceName As String, ByVal dwProperties As Long, ByRef lpBuffer As NS_SERVICE_INFO, ByRef lpdwBufferSize As Long, ByVal Reserved As Long) As Long


'=============================================================================================================
' getsockname
'
' Minimum Availability : Requires Windows Sockets 1.1 or later (Not supported on Windows 95)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getsockname function retrieves the local name for a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in]      Descriptor identifying a socket.
' name    [out]     Receives the address (name) of the socket.
' NameLen [in, out] Size of the name buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getsockname returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int getsockname (SOCKET s, struct sockaddr FAR *name, int FAR *namelen);
'=============================================================================================================
Public Declare Function getsockname Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Name As SOCKADDR, ByRef NameLen As Long) As Long


'=============================================================================================================
' getsockopt
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' See Also : http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/winsock/wsapiref_8qcy.htm
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets getsockopt function retrieves a socket option.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in]      Descriptor identifying a socket.
' Level   [in]      Level at which the option is defined; the supported levels include SOL_SOCKET and IPPROTO_TCP. See the Windows Sockets 2 Protocol-Specific Annex (a separate document included with the Platform SDK) for more information on protocol-specific levels.
' optname [in]      Socket option for which the value is to be retrieved.
' optval  [out]     Pointer to the buffer in which the value for the requested option is to be returned.
' optlen  [in, out] Pointer to the size of the optval buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, getsockopt returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int getsockopt (SOCKET s, int level, int optname, char FAR *optval, int FAR *optlen);
'=============================================================================================================
Public Declare Function getsockopt Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal Level As Long, ByVal OptName As Long, ByRef OptVal As Any, ByRef OptLen As Long) As Long


'=============================================================================================================
' GetTypeByName
'
' Minimum Availability : Requires Windows Sockets 1.1. (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The GetTypeByName function obtains a service type GUID for a network service specified by name.
'
' Important  The GetTypeByName function is a Microsoft-specific extension to the Windows Sockets 1.1
' specification. This function is obsolete. For the convenience of Windows Sockets 1.1 developers, the
' reference material is as follows.
'
' The functions detailed in Protocol-Independent Name Resolution provide equivalent functionality in
' Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpServiceName [in]  Pointer to a zero-terminated string that uniquely represents the name of the service. For example, "MY SNA SERVER."
' lpServiceType [out] Pointer to a variable to receive a globally unique identifier (GUID) that specifies the type of the network service. The header file Svcguid.h includes definitions of several GUID service types and macros for working with them.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is zero.
' If the function fails, the return value is SOCKET_ERROR( – 1). To get extended error information, call
' GetLastError. GetLastError can return the following extended error value.
' ____________________________________________________________________________________________________________
' INT GetTypeByName (LPTSTR lpServiceName, PGUID lpServiceType);
'=============================================================================================================
Public Declare Function GetTypeByName Lib "WSOCK32.DLL" Alias "GetTypeByNameA" (ByVal lpServiceName As String, ByRef lpServiceType As GUID) As Long


'=============================================================================================================
' htonl
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets htonl function converts a u_long from host to TCP/IP network byte order
' (which is big-endian).
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hostlong [in] 32-bit number in host byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' The htonl function returns the value in TCP/IP's network byte order.
' ____________________________________________________________________________________________________________
' u_long htonl (u_long hostlong);
'=============================================================================================================
Public Declare Function htonl Lib "WSOCK32.DLL" (ByVal HostLong As Long) As Long


'=============================================================================================================
' htons
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets htons function converts a u_short from host to TCP/IP network byte order
' (which is big-endian).
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hostshort [in] 16-bit number in host byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' The htons function returns the value in TCP/IP network byte order.
' ____________________________________________________________________________________________________________
' u_short htons (u_short hostshort);
'=============================================================================================================
Public Declare Function htons Lib "WSOCK32.DLL" (ByVal HostShort As Integer) As Integer


'=============================================================================================================
' inet_addr
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets inet_addr function converts a string containing an (Ipv4) Internet Protocol dotted
' address into a proper address for the IN_ADDR structure.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' cp [in] Null-terminated character string representing a number expressed in the Internet standard ".'' (dotted) notation.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, inet_addr returns an unsigned long value containing a suitable binary representation
' of the Internet address given. If the string in the cp parameter does not contain a legitimate Internet
' address, for example if a portion of an "a.b.c.d" address exceeds 255, then inet_addr returns the value
' INADDR_NONE.
' ____________________________________________________________________________________________________________
' unsigned long inet_addr (const char FAR *cp);
'=============================================================================================================
Public Declare Function inet_addr Lib "WSOCK32.DLL" (ByVal IpAddress As String) As Long


'=============================================================================================================
' inet_ntoa
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets inet_ntoa function converts an (Ipv4) Internet network address into a string in Internet
' standard dotted format.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' in [in] Structure that represents an Internet host address.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, inet_ntoa returns a character pointer to a static buffer containing the text address
' in standard "." notation. Otherwise, it returns NULL.
' ____________________________________________________________________________________________________________
' char FAR * inet_ntoa (struct in_addr in);
'=============================================================================================================
Public Declare Function inet_ntoa Lib "WSOCK32.DLL" (ByVal InAddr As Long) As Long


'=============================================================================================================
' ioctlsocket
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets ioctlsocket function controls the I/O mode of a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s    [in]      Descriptor identifying a socket.
' cmd  [in]      Command to perform on the socket s.
' argp [in, out] Pointer to a parameter for cmd.
'
' Return:
' ¯¯¯¯¯¯¯
' Upon successful completion, the ioctlsocket returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int ioctlsocket (SOCKET s, long cmd, u_long FAR * argp);
'=============================================================================================================
Public Declare Function ioctlsocket Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal Cmd As Long, ByRef CmdParam As Long) As Long


'=============================================================================================================
' listen
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets listen function places a socket a state where it is listening for an incoming connection.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in] Descriptor identifying a bound, unconnected socket.
' backlog [in] Maximum length of the queue of pending connections. If set to SOMAXCONN, the underlying service provider responsible for socket s will set the backlog to a maximum reasonable value. There is no standard provision to obtain the actual backlog value.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, listen returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int listen (SOCKET s, int backlog);
'=============================================================================================================
Public Declare Function listen Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal BackLog As Long) As Long


'=============================================================================================================
' ntohl
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets ntohl function converts a u_long from TCP/IP network order to host byte order
' (which is little-endian on Intel processors).
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' netlong [in] 32-bit number in TCP/IP network byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' The ntohl function always returns a value in host byte order. If the netlong parameter was already in
' host byte order, then no operation is performed.
' ____________________________________________________________________________________________________________
' u_long ntohl (u_long netlong);
'=============================================================================================================
Public Declare Function ntohl Lib "WSOCK32.DLL" (ByVal NetLong As Long) As Long


'=============================================================================================================
' ntohs
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets ntohs function converts a u_short from TCP/IP network byte order to host byte order
' (which is little-endian on Intel processors).
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' netshort [in] 16-bit number in TCP/IP network byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' The ntohs function returns the value in host byte order. If the netshort parameter was already in host
' byte order, then no operation is performed.
' ____________________________________________________________________________________________________________
' u_short ntohs (u_short netshort);
'=============================================================================================================
Public Declare Function ntohs Lib "WSOCK32.DLL" (ByVal NetShort As Integer) As Integer


'=============================================================================================================
' recv
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets recv function receives data from a connected socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s     [in]  Descriptor identifying a connected socket.
' buf   [out] Buffer for the incoming data.
' len   [in]  Length of buf.
' flags [in]  Flag specifying the way in which the call is made.  This can be MSG_PEEK or MSG_OOB.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, recv returns the number of bytes received. If the connection has been gracefully closed,
' the return value is zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be
' retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int recv (SOCKET s, char FAR *buf, int len, int flags);
'=============================================================================================================
Public Declare Function recv Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Buffer As Any, ByVal BufferLength As Long, ByVal Flags As Long) As Long


'=============================================================================================================
' recvfrom
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets recvfrom function receives a datagram and stores the source address.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in]      Descriptor identifying a bound socket.
' buf     [out]     Buffer for the incoming data.
' len     [in]      Length of buf.
' flags   [in]      Indicator specifying the way in which the call is made.  This can be MSG_PEEK or MSG_OOB.
' from    [out]     Optional pointer to a buffer that will hold the source address upon return.
' fromlen [in, out] Optional pointer to the size of the from buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, recvfrom returns the number of bytes received. If the connection has been gracefully
' closed, the return value is zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific error
' code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int recvfrom (SOCKET s, char FAR* buf, int len, int flags, struct sockaddr FAR *from, int FAR *fromlen);
'=============================================================================================================
Public Declare Function recvfrom Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Buffer As Any, ByVal BufferLength As Long, ByVal Flags As Long, ByRef From As SOCKADDR, ByRef FromLen As Long) As Long


'=============================================================================================================
' select
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets select function determines the status of one or more sockets, waiting if necessary,
' to perform synchronous I/O.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' nfds      [in]      Ignored. The nfds parameter is included only for compatibility with Berkeley sockets.
' readfds   [in, out] Optional pointer to a set of sockets to be checked for readability.
' writefds  [in, out] Optional pointer to a set of sockets to be checked for writability
' exceptfds [in, out] Optional pointer to a set of sockets to be checked for errors.
' Timeout   [in]      Maximum time for select to wait, provided in the form of a TIMEVAL structure. Set the timeout parameter to NULL for blocking operation.
'
' Return:
' ¯¯¯¯¯¯¯
' The select function returns the total number of socket handles that are ready and contained in the fd_set
' structures, zero if the time limit expired, or SOCKET_ERROR if an error occurred. If the return value is
' SOCKET_ERROR, WSAGetLastError can be used to retrieve a specific error code.
' ____________________________________________________________________________________________________________
' int select (int nfds, fd_set FAR *readfds, fd_set FAR *writefds, fd_set FAR *exceptfds, const struct timeval FAR *timeout);
'=============================================================================================================
Public Declare Function select_API Lib "WSOCK32.DLL" Alias "select" (ByVal Reserved As Long, ByRef ReadFds As FD_SET, ByRef WriteFds As FD_SET, ByRef ExceptFds As FD_SET, ByRef Timeout As TIMEVAL) As Long


'=============================================================================================================
' send
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets send function sends data on a connected socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s     [in] Descriptor identifying a connected socket.
' buf   [in] Buffer containing the data to be transmitted.
' len   [in] Length of the data in buf.
' flags [in] Indicator specifying the way in which the call is made.  This can be MSG_DONTROUTE or MSG_OOB.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, send returns the total number of bytes sent, which can be less than the number
' indicated by len for nonblocking sockets. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int send (SOCKET s, const char FAR *buf, int len, int flags);
'=============================================================================================================
Public Declare Function send Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Buffer As Any, ByVal BufferLength As Long, ByVal Flags As Long) As Long


'=============================================================================================================
' sendto
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets sendto function sends data to a specific destination.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s     [in] Descriptor identifying a (possibly connected) socket.
' buf   [in] Buffer containing the data to be transmitted.
' len   [in] Length of the data in buf.
' flags [in] Indicator specifying the way in which the call is made.  This can be MSG_DONTROUTE or MSG_OOB.
' to    [in] Optional pointer to the address of the target socket.
' tolen [in] Size of the address in to.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, sendto returns the total number of bytes sent, which can be less than the number
' indicated by len. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be
' retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int sendto (SOCKET s, const char FAR *buf, int len, int flags, const struct sockaddr FAR *to, int tolen);
'=============================================================================================================
Public Declare Function sendto Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Buffer As Any, ByVal BufferLength As Long, ByVal Flags As Long, ByRef ToSocket As SOCKADDR, ByVal ToLength As Long) As Long


'=============================================================================================================
' SetService
'
' Minimum Availability : Windows Sockets 1.1 or later (Not supported on Windows 95 - Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The SetService function registers or removes from the registry a network service within one or more name
' spaces. The function can also add or remove a network service type within one or more name spaces.
'
' Important  The SetService function is obsolete. For the convenience of Windows Sockets 1.1 developers,
' the reference material is as follows.
'
' The functions detailed in Protocol-Independent Name Resolution provide equivalent functionality in
' Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' dwNameSpace        [in]  Name space, or a set of default name spaces, within which the function will operate.  Use one of the following constants to specify a name space: NS_DEFAULT, NS_DNS, NS_NDS, NS_NETBT, NS_SAP, NS_TCPIP_HOSTS, NS_TCPIP_LOCAL
' dwOperation        [in]  Specifies the operation that the function will perform. Use one of the following values to specify an operation: SERVICE_REGISTER, SERVICE_DEREGISTER, SERVICE_FLUSH, SERVICE_ADD_TYPE, SERVICE_DELETE_TYPE
' dwFlags            [in]  Set of bit flags that modify the function's operation. You can set one or more of the following bit flags: SERVICE_FLAG_DEFER, SERVICE_FLAG_HARD
' lpServiceInfo      [in]  Pointer to a SERVICE_INFO structure that contains information about the network service or service type.
' lpServiceAsyncInfo [in]  Reserved for future use. Must be set to NULL.
' lpdwStatusFlags    [out] Set of bit flags that receive function status information. The following bit flag is defined: SET_SERVICE_PARTIAL_SUCCESS
'
' Return:
' ¯¯¯¯¯¯¯
' If the function fails, the return value is SOCKET_ERROR. To get extended error information, call
' GetLastError.
' ____________________________________________________________________________________________________________
' INT SetService (DWORD dwNameSpace, DWORD dwOperation, DWORD dwFlags, LPSERVICE_INFO lpServiceInfo, LPSERVICE_ASYNC_INFO lpServiceAsyncInfo, LPDWORD lpdwStatusFlags);
'=============================================================================================================
Public Declare Function SetService Lib "WSOCK32.DLL" Alias "SetServiceA" (ByVal dwNameSpace As Long, ByVal dwOperation As Long, ByVal dwFlags As Long, ByRef lpServiceInfo As SERVICE_INFO, ByVal Reserved As Long, ByRef lpdwStatusFlags As Long) As Long


'=============================================================================================================
' setsockopt
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' See Also : http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/winsock/wsapiref_94aa.htm
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets setsockopt function sets a socket option.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s       [in] Descriptor identifying a socket.
' Level   [in] Level at which the option is defined; the supported levels include SOL_SOCKET and IPPROTO_TCP. See the Windows Sockets 2 Protocol-Specific Annex (a separate document included with the Platform SDK) for more information on protocol-specific levels.
' OptName [in] Socket option for which the value is to be set.
' OptVal  [in] Pointer to the buffer in which the value for the requested option is supplied.
' OptLen  [in] Size of the optval buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, setsockopt returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int setsockopt (SOCKET s, int level, int optname, const char FAR *optval, int optlen);
'=============================================================================================================
Public Declare Function setsockopt Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal Level As Long, ByVal OptionName As Long, ByRef OptionValue As Any, ByVal OptionLength As Long) As Long


'=============================================================================================================
' shutdown
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets shutdown function disables sends or receives on a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s   [in] Descriptor identifying a socket.
' how [in] Flag that describes what types of operation will no longer be allowed.  This can be one of the following flags: SD_RECEIVE, SD_SEND, SD_BOTH
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, shutdown returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a specific
' error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int shutdown (SOCKET s, int how);
'=============================================================================================================
Public Declare Function shutdown Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal How As Long) As Long


'=============================================================================================================
' socket
'
' Minimum Availability : Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets socket function creates a socket that is bound to a specific service provider.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' af       [in] Address family specification.
' type     [in] Type specification for the new socket. The following are the type specifications supported: SOCK_STREAM, SOCK_DGRAM, SOCK_RAW, SOCK_RDM, SOCK_SEQPACKET
' protocol [in] Protocol to be used with the socket that is specific to the indicated address family.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, socket returns a descriptor referencing the new socket. Otherwise, a value of
' INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' SOCKET socket (int af, int type, int protocol);
'=============================================================================================================
Public Declare Function socket Lib "WSOCK32.DLL" (ByVal AddressFamily As Long, ByVal SocketType As Long, ByVal Protocol As Long) As Long


'=============================================================================================================
' TransmitFile
'
' Minimum Availability : Windows Sockets 1.1 or later (Not supported on Windows 95)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets TransmitFile function transmits file data over a connected socket handle. This
' function uses the operating system's cache manager to retrieve the file data, and provides high-
' performance file data transfer over sockets.
'
' Note : This function is a Microsoft-specific extension to the Windows Sockets specification. For more
' information, see Microsoft Extensions and Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hSocket               [in] Handle to a connected socket. The TransmitFile function will transmit the file data over this socket. The socket specified by hSocket must be a connection-oriented socket; the TransmitFile function does not support datagram sockets. Sockets of type SOCK_STREAM, SOCK_SEQPACKET, or SOCK_RDM are connection-oriented sockets.
' hFile                 [in] Handle to the open file that the TransmitFile function transmits. Since operating system reads the file data sequentially, you can improve caching performance by opening the handle with FILE_FLAG_SEQUENTIAL_SCAN. The hFile parameter is optional; if the hFile parameter is NULL, only data in the header and/or the tail buffer is transmitted; any additional action, such as socket disconnect or reuse, is performed as specified by the dwFlags parameter.
' nNumberOfBytesToWrite [in] Number of file bytes to transmit. The TransmitFile function completes when it has sent the specified number of bytes, or when an error occurs, whichever occurs first.  Set nNumberOfBytesToWrite to zero in order to transmit the entire file.
' nNumberOfBytesPerSend [in] Size of each block of data sent in each send operation, in bytes. This specification is used by Windows sockets layer. To select the default send size, set nNumberOfBytesPerSend to zero.  The nNumberOfBytesPerSend parameter is useful for message protocols that have limitations on the size of individual send requests.
' lpOverlapped          [in] Pointer to an OVERLAPPED structure. If the socket handle has been opened as overlapped, specify this parameter in order to achieve an overlapped (asynchronous) I/O operation. By default, socket handles are opened as overlapped.  You can use lpOverlapped to specify an offset within the file at which to start the file data transfer by setting the Offset and OffsetHigh member of the OVERLAPPED structure. If lpOverlapped is NULL, the transmission of data always starts at the current byte offset in the file.  When lpOverlapped is not NULL, the overlapped I/O might not finish before TransmitFile returns. In that case, the TransmitFile function returns FALSE, and GetLastError returns ERROR_IO_PENDING. This enables the caller to continue processing while the file transmission operation completes. Windows will set the event specified by the hEvent member of the OVERLAPPED structure, or the socket specified by hSocket, to the signaled state upon completion of the data transmission request.
' lpTransmitBuffers     [in] Pointer to a TRANSMIT_FILE_BUFFERS data structure that contains pointers to data to send before and after the file data is sent. Set the lpTransmitBuffers parameter to NULL if you want to transmit only the file data.
' dwFlags               [in] The dwFlags parameter has six settings: TF_DISCONNECT, TF_REUSE_SOCKET, TF_USE_DEFAULT_WORKER, TF_USE_SYSTEM_THREAD, TF_USE_KERNEL_APC, TF_WRITE_BEHIND
'
' Return:
' ¯¯¯¯¯¯¯
' If the TransmitFile function succeeds, the return value is TRUE. Otherwise, the return value is FALSE.
'
' To get extended error information, call GetLastError. The function returns FALSE if an overlapped I/O
' operation is not complete before TransmitFile returns. In that case, GetLastError returns ERROR_IO_PENDING.
' ____________________________________________________________________________________________________________
' BOOL TransmitFile (SOCKET hSocket, HANDLE hFile, DWORD nNumberOfBytesToWrite, DWORD nNumberOfBytesPerSend, LPOVERLAPPED lpOverlapped, LPTRANSMIT_FILE_BUFFERS lpTransmitBuffers, DWORD dwFlags);
'=============================================================================================================
Public Declare Function TransmitFile Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal hFile As Long, ByVal nNumberOfBytesToWrite As Long, ByVal nNumberOfBytesPerSend As Long, ByRef lpOverlapped As OVERLAPPED, ByRef lpTransmitBuffers As TRANSMIT_FILE_BUFFERS, ByVal dwFlags As Long) As Long


'=============================================================================================================
' WSAAccept
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAccept function conditionally accepts a connection based on the return value of a
' condition function, provides QOS flow specifications, and allows the transfer of connection data.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s              [in]      Descriptor identifying a socket that is listening for connections after a call to the listen function.
' addr           [out]     Optional pointer to a buffer that receives the address of the connecting entity, as known to the communications layer. The exact format of the addr parameter is determined by the address family established when the socket was created.
' AddrLen        [in, out] Optional pointer to an integer that contains the length of the address addr.
' lpfnCondition  [in]      Procedure instance address of the optional, application-supplied condition function that will make an accept/reject decision based on the caller information passed in as parameters.  (See ConditionFunc callback function below)
' dwCallbackData [in]      Callback data passed back to the application as the value of the dwCallbackData parameter of the condition function. This parameter is not interpreted by Windows Sockets.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAAccept returns a value of type SOCKET that is a descriptor for the accepted socket.
' Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be retrieved by calling
' WSAGetLastError.
'
' The integer referred to by addrlen initially contains the amount of space pointed to by addr. On return
' it will contain the actual length in bytes of the address returned.
' ____________________________________________________________________________________________________________
' SOCKET WSAAccept (SOCKET s, struct sockaddr FAR *addr, LPINT addrlen, LPCONDITIONPROC lpfnCondition, DWORD dwCallbackData);
'=============================================================================================================
Public Declare Function WSAAccept Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef SocketAddress As SOCKADDR, ByRef AddressLength As Long, ByVal lpfnCondition As Long, ByVal dwCallbackData As Long) As Long


'=============================================================================================================
' WSAAddressToString
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAddressToString function converts all components of a SOCKADDR structure into a
' human-readable string representation of the address.
'
' This is intended to be used mainly for display purposes. If the caller wants the translation to be done
' by a particular provider, it should supply the corresponding WSAPROTOCOL_INFO structure in the
' lpProtocolInfo parameter.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpsaAddress             [in]      Pointer to the SOCKADDR structure to translate into a string.
' dwAddressLength         [in]      Length of the address in SOCKADDR, which may vary in size with different protocols.
' lpProtocolInfo          [in]      (Optional) The WSAPROTOCOL_INFO structure for a particular provider. If this is NULL, the call is routed to the provider of the first protocol supporting the address family indicated in lpsaAddress.
' lpszAddressString       [in, out] Buffer that receives the human-readable address string.
' lpdwAddressStringLength [in, out] On input, the length of the AddressString buffer. On output, returns the length of the string actually copied into the buffer. If the supplied buffer is not large enough, the function fails with a specific error of WSAEFAULT and this parameter is updated with the required size in bytes.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAAddressToString returns a value of zero. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAAddressToString (LPSOCKADDR lpsaAddress, DWORD dwAddressLength, LPWSAPROTOCOL_INFO lpProtocolInfo, LPTSTR lpszAddressString, LPDWORD lpdwAddressStringLength);
'=============================================================================================================
Public Declare Function WSAAddressToString Lib "WS2_32.DLL" Alias "WSAAddressToStringA" (ByRef lpsaAddress As SOCKADDR, ByVal dwAddressLength As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFO, ByVal lpszAddressString As String, ByRef lpdwAddressStringLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetHostByAddr
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetHostByAddr function asynchronously retrieves host information that
' corresponds to an address.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that will receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' addr   [in]  Pointer to the network address for the host. Host addresses are stored in network byte order.
' len    [in]  Length of the address.
' type   [in]  Type of the address.
' buf    [out] Pointer to the data area to receive the HOSTENT data. The data area must be larger than the size of a HOSTENT structure because the supplied data area is used by Windows Sockets to contain a HOSTENT structure and all of the data referenced by members of the HOSTENT structure. A buffer of MAXGETHOSTSTRUCT (1024) bytes is recommended.
' buflen [in]  Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It does
' not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetHostByAddr returns a nonzero value of type HANDLE that is the asynchronous
' task handle (not to be confused with a Windows HTASK) for the request. This value can be used in two ways.
' It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used to match up
' asynchronous operations and completion messages by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetHostByAddr returns a zero value, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetHostByAddr (HWND hWnd, unsigned int wMsg, const char FAR *addr, int len, int type, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetHostByAddr Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Address As Long, ByVal AddressLength As Long, ByVal AddressType As Long, ByRef Buffer As Any, ByVal BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetHostByName
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetHostByName function asynchronously retrieves host information
' corresponding to a host name.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that will receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' name   [in]  Pointer to the null-terminated name of the host.
' buf    [out] Pointer to the data area to receive the HOSTENT data. The data area must be larger than the size of a HOSTENT structure because the supplied data area is used by Windows Sockets to contain a HOSTENT structure and all of the data referenced by members of the HOSTENT structure. A buffer of MAXGETHOSTSTRUCT (1024) bytes is recommended.
' buflen [in]  Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It does
' not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetHostByName returns a nonzero value of type HANDLE that is the asynchronous
' task handle (not to be confused with a Windows HTASK) for the request. This value can be used in two ways.
' It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used to match up
' asynchronous operations and completion messages by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetHostByName returns a zero value, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetHostByName (HWND hWnd, unsigned int wMsg, const char FAR *name, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetHostByName Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal HostName As String, ByRef Buffer As Any, ByVal BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetProtoByName
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetProtoByName function gets protocol information corresponding to a
' protocol name asynchronously.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that will receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' name   [in]  Pointer to the null-terminated protocol name to be resolved.
' buf    [out] Pointer to the data area to receive the PROTOENT data. The data area must be larger than the size of a PROTOENT structure because the data area is used by Windows Sockets to contain a PROTOENT structure and all of the data that is referenced by members of the PROTOENT structure. A buffer of MAXGETHOSTSTRUCT (1024) bytes is recommended.
' buflen [out] Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It does
' not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetProtoByName returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value can be used
' in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used to match
' up asynchronous operations and completion messages, by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetProtoByName returns a zero value, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetProtoByName (HWND hWnd, unsigned int wMsg, const char FAR *name, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetProtoByName Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal ProtocolName As String, ByRef Buffer As Any, ByRef BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetProtoByNumber
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetProtoByNumber function asynchronously retrieves protocol information
' corresponding to a protocol number.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that will receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' Number [in]  Protocol number to be resolved, in host byte order.
' buf    [out] Pointer to the data area to receive the PROTOENT data. The data area must be larger than the size of a PROTOENT structure because the data area is used by Windows Sockets to contain a PROTOENT structure and all of the data that is referenced by members of the PROTOENT structure. A buffer of MAXGETHOSTSTRUCT bytes is recommended.
' buflen [in]  Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It
' does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetProtoByNumber returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value can be
' used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used
' to match up asynchronous operations and completion messages, by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetProtoByNumber returns a zero value,
' and a specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetProtoByNumber (HWND hWnd, unsigned int wMsg, int number, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetProtoByNumber Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal ProtocolNumber As Long, ByRef Buffer As Any, ByVal BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetServByName
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetServByName function asynchronously retrieves service information
' corresponding to a service name and port.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that should receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' name   [in]  Pointer to a null-terminated service name.
' Proto  [in]  Pointer to a protocol name. This can be NULL, in which case WSAAsyncGetServByName will search for the first service entry for which s_name or one of the s_aliases matches the given name. Otherwise, WSAAsyncGetServByName matches both name and proto.
' buf    [out] Pointer to the data area to receive the SERVENT data. The data area must be larger than the size of a SERVENT structure because the data area supplied is used by Windows Sockets to contain a SERVENT structure and all of the data that is referenced by members of the SERVENT structure. A buffer of MAXGETHOSTSTRUCT bytes is recommended.
' buflen [in]  Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It
' does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetServByName returns a nonzero value of type HANDLE that is the
' asynchronous task handle for the request (not to be confused with a Windows HTASK). This value can be
' used in two ways. It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used
' to match up asynchronous operations and completion messages, by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncServByName returns a zero value, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetServByName (HWND hWnd, unsigned int wMsg, const char FAR *name, const char FAR *proto, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetServByName Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal ServiceName As String, ByVal ProtocolName As String, ByRef Buffer As Any, ByVal BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncGetServByPort
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncGetServByPort function gets service information corresponding to a port
' and protocol asynchronously.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hWnd   [in]  Handle of the window that should receive a message when the asynchronous request completes.
' wMsg   [in]  Message to be received when the asynchronous request completes.
' Port   [in]  Port for the service, in network byte order.
' Proto  [in]  Pointer to a protocol name. This can be NULL, in which case WSAAsyncGetServByPort will search for the first service entry for which s_port match the given port. Otherwise, WSAAsyncGetServByPort matches both port and proto.
' buf    [out] Pointer to the data area to receive the SERVENT data. The data area must be larger than the size of a SERVENT structure because the data area supplied is used by Windows Sockets to contain a SERVENT structure and all of the data that is referenced by members of the SERVENT structure. A buffer of MAXGETHOSTSTRUCT bytes is recommended.
' buflen [in]  Size of data area for the buf parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value specifies whether or not the asynchronous operation was successfully initiated. It
' does not imply success or failure of the operation itself.
'
' If no error occurs, WSAAsyncGetServByPort returns a nonzero value of type HANDLE that is the asynchronous
' task handle for the request (not to be confused with a Windows HTASK). This value can be used in two ways.
' It can be used to cancel the operation using WSACancelAsyncRequest, or it can be used to match up
' asynchronous operations and completion messages, by examining the wParam message parameter.
'
' If the asynchronous operation could not be initiated, WSAAsyncGetServByPort returns a zero value, and a
' specific error number can be retrieved by calling WSAGetLastError.
'
' The following error codes can be set when an application window receives a message. As described above,
' they can be extracted from the lParam in the reply message using the WSAGETASYNCERROR macro.
' ____________________________________________________________________________________________________________
' HANDLE WSAAsyncGetServByPort (HWND hWnd, unsigned int wMsg, int port, const char FAR *proto, char FAR *buf, int buflen);
'=============================================================================================================
Public Declare Function WSAAsyncGetServByPort Lib "WSOCK32.DLL" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal ProtocolName As String, ByRef Buffer As Any, ByVal BufferLength As Long) As Long


'=============================================================================================================
' WSAAsyncSelect
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAAsyncSelect function requests Windows message-based notification of network
' events for a socket.
'
' Note:
' ¯¯¯¯¯
' Winsock will not continually flood an application with messages for a particular network event.
' Having successfully posted notification of a particular event to an application window, no further
' message(s) for that network event will be posted to the application window until the application makes
' the function call that implicitly reenables notification of that network event.
'
' Network Event:               Re-enabling Function:
' -------------------------------------------------------
' FD_READ                      recv, recvfrom, WSARecv, or WSARecvFrom
' FD_WRITE                     send, sendto, WSASend, or WSASendTo
' FD_OOB                       recv, recvfrom, WSARecv, or WSARecvFrom
' FD_ACCEPT                    accept or WSAAccept unless the error code is WSATRY_AGAIN indicating that the condition function returned CF_DEFER
' FD_CONNECT                   NONE
' FD_CLOSE                     NONE
' FD_QOS                       WSAIoctl with command SIO_GET_QOS
' FD_GROUP_QOS                 Reserved
' FD_ROUTING_INTERFACE_CHANGE  WSAIoctl with command SIO_ROUTING_INTERFACE_CHANGE.
' FD_ADDRESS_LIST_CHANGE       WSAIoctl with command SIO_ADDRESS_LIST_CHANGE.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s      [in] Descriptor identifying the socket for which event notification is required.
' hWnd   [in] Handle identifying the window that will receive a message when a network event occurs.
' wMsg   [in] Message to be received when a network event occurs.
' lEvent [in] Bitmask that specifies a combination of network events in which the application is interested.  (See the FD_* constant declarations)
'
' Return:
' ¯¯¯¯¯¯¯
' If the WSAAsyncSelect function succeeds, the return value is zero provided the application's declaration
' of interest in the network event set was successful. Otherwise, the value SOCKET_ERROR is returned, and
' a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAAsyncSelect (SOCKET s, HWND hWnd, unsigned int wMsg, long lEvent);
'=============================================================================================================
Public Declare Function WSAAsyncSelect Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long


'=============================================================================================================
' WSACancelAsyncRequest
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSACancelAsyncRequest function cancels an incomplete asynchronous operation.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hAsyncTaskHandle [in] Handle that specifies the asynchronous operation to be canceled.
'
' Return:
' ¯¯¯¯¯¯¯
' The value returned by WSACancelAsyncRequest is zero if the operation was successfully canceled.
' Otherwise, the value SOCKET_ERROR is returned, and a specific error number can be retrieved by calling
' WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSACancelAsyncRequest (HANDLE hAsyncTaskHandle);
'=============================================================================================================
Public Declare Function WSACancelAsyncRequest Lib "WSOCK32.DLL" (ByVal hAsyncTaskHandle As Long) As Long


'=============================================================================================================
' WSACancelBlockingCall
'
' Minimum Availability : Requires Windows Sockets 1.1 (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' CancelS a blocking call which is currently in progress.
'
' The WSACancelBlockingCall function has been removed in compliance with the Windows Sockets 2
' specification, revision 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll and Windows Sockets 2 applications should not
' use this function. Windows Sockets 1.1 applications that call this function are still supported through
' the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during calls to
' blocking functions. Instead of using blocking hooks, an applications should use a separate thread
' (separate from the main GUI thread) for network activity.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' The value returned by WSACancelBlockingCall() is 0 if the operation was successfully canceled.
' Otherwise the value SOCKET_ERROR is returned, and a specific error number may be retrieved by calling
' WSAGetLastError().
' ____________________________________________________________________________________________________________
' int WSACancelBlockingCall (void);
'=============================================================================================================
Public Declare Function WSACancelBlockingCall Lib "WSOCK32.DLL" ()


'=============================================================================================================
' WSACleanup
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSACleanup function terminates use of the WS2_32.DLL / WSOCK32.DLL
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
'
' Attempting to call WSACleanup from within a blocking hook and then failing to check the return code is
' a common programming error in Windows Socket 1.1 applications. If an application needs to quit while a
' blocking call is outstanding, the application must first cancel the blocking call with
' WSACancelBlockingCall then issue the WSACleanup call once control has been returned to the application.
'
' In a multithreaded environment, WSACleanup terminates Windows Sockets operations for all threads.
' ____________________________________________________________________________________________________________
' int  WSACleanup (void);
'=============================================================================================================
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long


'=============================================================================================================
' WSACloseEvent
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSACloseEvent function closes an open event object handle.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEvent [in] Object handle identifying the open event.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is TRUE.  If the function fails, the return value is FALSE.
' To get extended error information, call WSAGetLastError.
' ____________________________________________________________________________________________________________
' BOOL WSACloseEvent (WSAEVENT hEvent);
'=============================================================================================================
Public Declare Function WSACloseEvent Lib "WS2_32.DLL" (ByVal hEvent As Long) As Long


'=============================================================================================================
' WSAConnect
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAConnect function establishes a connection to another socket application, exchanges
' connect data, and specifies needed quality of service based on the supplied FLOWSPEC structure.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s            [in]  Descriptor identifying an unconnected socket.
' name         [in]  Name of the socket in the other application to which to connect.
' namelen      [in]  Length of the name.
' lpCallerData [in]  Pointer to the user data that is to be transferred to the other socket during connection establishment.
' lpCalleeData [out] Pointer to the user data that is to be transferred back from the other socket during connection establishment.
' lpSQOS       [in]  Pointer to the FLOWSPEC structures for socket s, one for each direction.
' lpGQOS       [in]  Reserved. Should be NULL.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAConnect returns zero. Otherwise, it returns SOCKET_ERROR, and a specific error
' code can be retrieved by calling WSAGetLastError. On a blocking socket, the return value indicates
' success or failure of the connection attempt.
'
' With a nonblocking socket, the connection attempt cannot be completed immediately. In this case,
' WSAConnect will return SOCKET_ERROR, and WSAGetLastError will return WSAEWOULDBLOCK; the application
' could therefore:
'   - Use select to determine the completion of the connection request by checking if the socket is writeable.
'   - If your application is using WSAAsyncSelect to indicate interest in connection events, then your application will receive an FD_CONNECT notification when the connect operation is complete(successful or not).
'   - If your application is using WSAEventSelect to indicate interest in connection events, then the associated event object will be signaled when the connect operation is complete (successful or not).
'
' For a nonblocking socket, until the connection attempt completes all subsequent calls to WSAConnect on
' the same socket will fail with the error code WSAEALREADY.
'
' If the return error code indicates the connection attempt failed (that is, WSAECONNREFUSED, WSAENETUNREACH,
' WSAETIMEDOUT) the application can call WSAConnect again for the same socket.
' ____________________________________________________________________________________________________________
' int  WSAConnect (SOCKET s, const struct sockaddr FAR *name, int namelen, LPWSABUF lpCallerData, LPWSABUF lpCalleeData, LPQOS lpSQOS, LPQOS lpGQOS);
'=============================================================================================================
Public Declare Function WSAConnect Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal SocketName As SOCKADDR, ByVal NameLength As Long, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long) As Long


'=============================================================================================================
' WSACreateEvent
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSACreateEvent function creates a new event object.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSACreateEvent returns the handle of the event object.
' Otherwise, the return value is WSA_INVALID_EVENT. To get extended error information, call WSAGetLastError.
' ____________________________________________________________________________________________________________
' WSAEVENT WSACreateEvent (void);
'=============================================================================================================
Public Declare Function WSACreateEvent Lib "WS2_32.DLL" () As Long ' WSAEVENT = HANDLE = Long


'=============================================================================================================
' WSADuplicateSocket
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSADuplicateSocket function returns a WSAPROTOCOL_INFO structure that can be used
' to create a new socket descriptor for a shared socket. The WSADuplicateSocket function cannot be used
' on a QOS-enabled socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s              [in]  Descriptor identifying the local socket.
' dwProcessId    [in]  Process identifier of the target process in which the duplicated socket will be used.
' lpProtocolInfo [out] Pointer to a buffer, allocated by the client, that is large enough to contain a WSAPROTOCOL_INFO structure. The service provider copies the protocol information structure contents to this buffer.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSADuplicateSocket returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSADuplicateSocket (SOCKET s, DWORD dwProcessId, LPWSAPROTOCOL_INFO lpProtocolInfo);
'=============================================================================================================
Public Declare Function WSADuplicateSocket Lib "WS2_32.DLL" Alias "WSADuplicateSocketA" (ByVal hSocket As Long, ByVal dwProcessId As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFO) As Long


'=============================================================================================================
' WSAEnumNameSpaceProviders
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAEnumNameSpaceProviders function retrieves information about available name spaces.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpdwBufferLength [in, out] On input, the number of bytes contained in the buffer pointed to by lpnspBuffer. On output (if the function fails, and the error is WSAEFAULT), the minimum number of bytes to pass for the lpnspBuffer to retrieve all the requested information. The passed-in buffer must be sufficient to hold all of the name space information.
' lpnspBuffer      [out]     Buffer that is filled with WSANAMESPACE_INFO structures. The returned structures are located consecutively at the head of the buffer. Variable sized information referenced by pointers in the structures point to locations within the buffer located between the end of the fixed sized structures and the end of the buffer. The number of structures filled in is the return value of WSAEnumNameSpaceProviders.
'
' Return:
' ¯¯¯¯¯¯¯
' The WSAEnumNameSpaceProviders function returns the number of WSANAMESPACE_INFO structures copied into
' lpnspBuffer. Otherwise, the value SOCKET_ERROR is returned, and a specific error number can be retrieved
' by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAAPI WSAEnumNameSpaceProviders (LPDWORD lpdwBufferLength, LPWSANAMESPACE_INFO lpnspBuffer);
'=============================================================================================================
Public Declare Function WSAEnumNameSpaceProviders Lib "WS2_32.DLL" Alias "WSAEnumNameSpaceProvidersA" (ByRef lpdwBufferLength As Long, ByRef lpnspBuffer As WSANAMESPACE_INFO) As Long


'=============================================================================================================
' WSAEnumNetworkEvents
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAEnumNetworkEvents function discovers occurrences of network events for the
' indicated socket, clear internal network event records, and reset event objects (optional).
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s               [in]  Descriptor identifying the socket.
' hEventObject    [in]  Optional handle identifying an associated event object to be reset.
' lpNetworkEvents [out] Pointer to a WSANETWORKEVENTS structure that is filled with a record of network events that occurred and any associated error codes.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is
' returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAEnumNetworkEvents (SOCKET s, WSAEVENT hEventObject, LPWSANETWORKEVENTS lpNetworkEvents);
'=============================================================================================================
Public Declare Function WSAEnumNetworkEvents Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal hEventObject As Long, ByRef lpNetworkEvents As WSANETWORKEVENTS) As Long


'=============================================================================================================
' WSAEnumProtocols
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAEnumProtocols function retrieves information about available transport protocols.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpiProtocols     [in]      Null-terminated array of iProtocol values. This parameter is optional; if lpiProtocols is NULL, information on all available protocols is returned. Otherwise, information is retrieved only for those protocols listed in the array.
' lpProtocolBuffer [out]     Buffer that is filled with WSAPROTOCOL_INFO structures.
' lpdwBufferLength [in, out] On input, the count of bytes in the lpProtocolBuffer buffer passed to WSAEnumProtocols. On output, the minimum buffer size that can be passed to WSAEnumProtocols to retrieve all the requested information. This routine has no ability to enumerate over multiple calls; the passed-in buffer must be large enough to hold all entries in order for the routine to succeed. This reduces the complexity of the API and should not pose a problem because the number of protocols loaded on a machine is typically small.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAEnumProtocols returns the number of protocols to be reported. Otherwise, a value
' of SOCKET_ERROR is returned and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAEnumProtocols (LPINT lpiProtocols, LPWSAPROTOCOL_INFO lpProtocolBuffer, ILPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function WSAEnumProtocols Lib "WS2_32.DLL" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols As Long, ByRef lpProtocolBuffer As WSAPROTOCOL_INFO, ByRef lpdwBufferLength As Long) As Long


'=============================================================================================================
' WSAEventSelect
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAEventSelect function specifies an event object to be associated with the
' supplied set of FD_XXX network events.
'
' Note:
' ¯¯¯¯¯
' Having successfully recorded the occurrence of the network event (by setting the corresponding bit in
' the internal network event record) and signaled the associated event object, no further actions are taken
' for that network event until the application makes the function call that implicitly reenables the setting
' of that network event and signaling of the associated event object.
'
' Network Event:               Re-enabling Function:
' -------------------------------------------------------
' FD_READ                      recv, recvfrom, WSARecv, or WSARecvFrom.
' FD_WRITE                     send, sendto, WSASend, or WSASendTo.
' FD_OOB                       recv, recvfrom, WSARecv, or WSARecvFrom.
' FD_ACCEPT                    accept or WSAAccept unless the error code returned is WSATRY_AGAIN indicating that the condition function returned CF_DEFER.
' FD_CONNECT                   None
' FD_CLOSE                     None
' FD_QOS                       WSAIoctl with command SIO_GET_QOS.
' FD_GROUP_QOS                 Reserved
' FD_ROUTING_INTERFACE_CHANGE  WSAIoctl with command SIO_ROUTING_INTERFACE_CHANGE.
' FD_ADDRESS_LIST_CHANGE       WSAIoctl with command SIO_ADDRESS_LIST_CHANGE.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s              [in] Descriptor identifying the socket.
' hEventObject   [in] Handle identifying the event object to be associated with the supplied set of FD_XXX network events.
' lNetworkEvents [in] Bitmask that specifies the combination of FD_XXX network events in which the application has interest.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the application's specification of the network events and the associated
' event object was successful. Otherwise, the value SOCKET_ERROR is returned, and a specific error number
' can be retrieved by calling WSAGetLastError.
'
' As in the case of the select and WSAAsyncSelect functions, WSAEventSelect will frequently be used to
' determine when a data transfer operation (send or recv) can be issued with the expectation of immediate
' success. Nevertheless, a robust application must be prepared for the possibility that the event object
' is set and it issues a Windows Sockets call that returns WSAEWOULDBLOCK immediately. For example, the
' following sequence of operations is possible:
'   - Data arrives on socket s; Windows Sockets sets the WSAEventSelect event object.
'   - The application does some other processing.
'   - While processing, the application issues an ioctlsocket(s, FIONREAD...) and notices that there is data ready to be read.
'   - The application issues a recv(s,...) to read the data.
'   - The application eventually waits on the event object specified in WSAEventSelect, which returns immediately indicating that data is ready to read.
'   - The application issues recv(s,...), which fails with the error WSAEWOULDBLOCK.
' ____________________________________________________________________________________________________________
' int WSAEventSelect (SOCKET s, WSAEVENT hEventObject, long lNetworkEvents);
'=============================================================================================================
Public Declare Function WSAEventSelect Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal hEventObject As Long, ByVal lNetworkEvents As Long) As Long


'=============================================================================================================
' WSAGetLastError
'
' Minimum Availability : Requires Windows Sockets 1.1 or later
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAGetLastError function gets the error status for the last operation that failed.
'
' Note:
' ¯¯¯¯¯
' A successful function call, or a call to WSAGetLastError, does not reset the error code. To reset the
' error code, use the WSASetLastError function call with iError set to zero. A getsockopt SO_ERROR also
' resets the error code to zero.
'
' The WSAGetLastError function should not be used to check for an error value on receipt of an asynchronous
' message. In this case, the error value is passed in the lParam parameter of the message, and this can
' differ from the value returned by WSAGetLastError.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' The return value indicates the error code for this thread's last Windows Sockets operation that failed.
' ____________________________________________________________________________________________________________
' int  WSAGetLastError (void);
'=============================================================================================================
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long


'=============================================================================================================
' WSAGetOverlappedResult
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAGetOverlappedResult function returns the results of an overlapped operation on
' the specified socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s            [in]  Descriptor identifying the socket. This is the same socket that was specified when the overlapped operation was started by a call to WSARecv, WSARecvFrom, WSASend, WSASendTo, or WSAIoctl.
' lpOverlapped [in]  Pointer to a WSAOVERLAPPED structure that was specified when the overlapped operation was started.
' lpcbTransfer [out] Pointer to a 32-bit variable that receives the number of bytes that were actually transferred by a send or receive operation, or by WSAIoctl.
' fWait        [in]  Flag that specifies whether the function should wait for the pending overlapped operation to complete. If TRUE, the function does not return until the operation has been completed. If FALSE and the operation is still pending, the function returns FALSE and the WSAGetLastError function returns WSA_IO_INCOMPLETE. The fWait parameter may be set to TRUE only if the overlapped operation selected the event-based completion notification.
' lpdwFlags    [out] Pointer to a 32-bit variable that will receive one or more flags that supplement the completion status. If the overlapped operation was initiated through WSARecv or WSARecvFrom, this parameter will contain the results value for lpFlags parameter.
'
' Return:
' ¯¯¯¯¯¯¯
' If WSAGetOverlappedResult succeeds, the return value is TRUE. This means that the overlapped operation
' has completed successfully and that the value pointed to by lpcbTransfer has been updated. If
' WSAGetOverlappedResult returns FALSE, this means that either the overlapped operation has not completed,
' the overlapped operation completed but with errors, or the overlapped operation's completion status could
' not be determined due to errors in one or more parameters to WSAGetOverlappedResult. On failure, the value
' pointed to by lpcbTransfer will not be updated. Use WSAGetLastError to determine the cause of the failure
' (either of WSAGetOverlappedResult or of the associated overlapped operation).
' ____________________________________________________________________________________________________________
' BOOL WSAGetOverlappedResult (SOCKET s, LPWSAOVERLAPPED lpOverlapped, LPDWORD lpcbTransfer, BOOL fWait, LPDWORD lpdwFlags);
'=============================================================================================================
Public Declare Function WSAGetOverlappedResult Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpcbTransfer As Long, ByVal fWait As Long, ByRef lpdwFlags As Long) As Long


'=============================================================================================================
' WSAGetQOSByName
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAGetQOSByName function initializes a QOS structure based on a named template,
' or it supplies a buffer to retrieve an enumeration of the available template names.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s         [in]     Descriptor identifying a socket.
' lpQOSName [in out] Pointer to a specific quality of service template.
' lpQOS     [out]    Pointer to the QOS structure to be filled.
'
' Return:
' ¯¯¯¯¯¯¯
' If WSAGetQOSByName succeeds, the return value is TRUE. If the function fails, the return value is FALSE.
' To get extended error information, call WSAGetLastError.
' ____________________________________________________________________________________________________________
' BOOL WSAGetQOSByName (SOCKET s, LPWSABUF lpQOSName, lpQOS lpQOS);
'=============================================================================================================
Public Declare Function WSAGetQOSByName Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpQOSName As WSABUF, ByRef lpQOS As QOS) As Long


'=============================================================================================================
' WSAGetServiceClassInfo
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAGetServiceClassInfo function retrieves all of the class information (schema)
' pertaining to a specified service class from a specified name space provider.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpProviderId       [in]      Pointer to a GUID that identifies a specific name space provider.
' lpServiceClassId   [in]      Pointer to a GUID identifying the service class.
' lpdwBufferLength   [in, out] On input, the number of bytes contained in the buffer pointed to by lpServiceClassInfos. On output, if the function fails and the error is WSAEFAULT, then it contains the minimum number of bytes to pass for the lpServiceClassInfo to retrieve the record.
' lpServiceClassInfo [out]     Pointer to the service class information from the indicated name space provider for the specified service class.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the WSAGetServiceClassInfo was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAGetServiceClassInfo (LPGUID lpProviderId, LPGUID lpServiceClassId, LPDWORD lpdwBufferLength, LPWSASERVICECLASSINFO lpServiceClassInfo);
'=============================================================================================================
Public Declare Function WSAGetServiceClassInfo Lib "WS2_32.DLL" Alias "WSAGetServiceClassInfoA" (ByRef lpProviderId As GUID, ByRef lpServiceClassId As GUID, ByRef lpdwBufferLength As Long, ByRef lpServiceClassInfo As WSASERVICECLASSINFO) As Long


'=============================================================================================================
' WSAGetServiceClassNameByClassId
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAGetServiceClassNameByClassId function returns the name of the service associated
' with the given type. This name is the generic service name, like FTP or SNA, and not the name of a
' specific instance of that service.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpServiceClassId     [in]      Pointer to the GUID for the service class.
' lpszServiceClassName [out]     Pointer to the service name.
' lpdwBufferLength     [in, out] On input, the length of the buffer returned by lpszServiceClassName. On output, the length of the service name copied into lpszServiceClassName.
'
' Return:
' ¯¯¯¯¯¯¯
' The WSAGetServiceClassNameByClassId function returns a value of zero if successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAGetServiceClassNameByClassId (LPGUID lpServiceClassId, LPTSTR lpszServiceClassName, LPDWORD lpdwBufferLength);
'=============================================================================================================
Public Declare Function WSAGetServiceClassNameByClassId Lib "WS2_32.DLL" Alias "WSAGetServiceClassNameByClassIdA" (ByRef lpServiceClassId As GUID, ByVal lpszServiceClassName As String, ByRef lpdwBufferLength As Long) As Long


'=============================================================================================================
' WSAHtonl
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAHtonl function converts a u_long from host byte order to network byte order.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s         [in]  Descriptor identifying a socket.
' HostLong  [in]  32-bit number in host byte order.
' lpnetlong [out] Pointer to a 32-bit number in network byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAHtonl returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAHtonl (SOCKET s, u_long hostlong, u_long FAR * lpnetlong);
'=============================================================================================================
Public Declare Function WSAHtonl Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal HostLong As Long, ByRef lpNetLong As Long) As Long


'=============================================================================================================
' WSAHtons
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAHtons function converts a u_short from host byte order to network byte order.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s          [in]  Descriptor identifying a socket.
' HostShort  [in]  16-bit number in host byte order.
' lpnetshort [out] Pointer to a 16-bit number in network byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAHtons returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAHtons (SOCKET s, u_short hostshort, u_short FAR * lpnetshort);
'=============================================================================================================
Public Declare Function WSAHtons Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal HostShort As Integer, ByRef lpNetShort As Integer) As Long


'=============================================================================================================
' WSAInstallServiceClass
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAInstallServiceClass function registers a service class schema within a name space.
' This schema includes the class name, class identifier, and any name space-specific information that is
' common to all instances of the service, such as the SAP identifier or object identifier.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpServiceClassInfo [in] Service class to name space specific–type mapping information. Multiple mappings can be handled at one time.  See the section Service Class Data Structures for a description of pertinent data structures.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is returned,
' and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAInstallServiceClass (LPWSASERVICECLASSINFO lpServiceClassInfo);
'=============================================================================================================
Public Declare Function WSAInstallServiceClass Lib "WS2_32.DLL" Alias "WSAInstallServiceClassA" (ByRef lpServiceClassInfo As WSASERVICECLASSINFO) As Long


'=============================================================================================================
' WSAIoctl
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAIoctl function controls the mode of a socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                   [in]  Descriptor identifying a socket.
' dwIoControlCode     [in]  Control code of operation to perform.
' lpvInBuffer         [in]  Pointer to the input buffer.
' cbInBuffer          [in]  Size of the input buffer.
' lpvOutBuffer        [out] Pointer to the output buffer.
' cbOutBuffer         [in]  Size of the output buffer.
' lpcbBytesReturned   [out] Pointer to actual number of bytes of output.
' lpOverlapped        [in]  Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped sockets).
' lpCompletionRoutine [in]  Pointer to the completion routine called when the operation has been completed (ignored for nonoverlapped sockets).  (See CompletionRoutine callback function below)
'
' Return:
' ¯¯¯¯¯¯¯
' Upon successful completion, the WSAIoctl returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSAIoctl (SOCKET s, DWORD dwIoControlCode, LPVOID lpvInBuffer, DWORD cbInBuffer, LPVOID lpvOutBuffer, DWORD cbOutBuffer, LPDWORD lpcbBytesReturned, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSAIoctl Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal dwIoControlCode As Long, ByRef In_Buffer As Any, ByVal In_BufferLen As Long, ByRef Out_Buffer As Any, ByVal Out_BufferLen As Long, ByRef lpcbBytesReturned As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSAIsBlocking
'
' Minimum Availability : Requires Windows Sockets 1.1 (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' Determines if a blocking call is in progress.
'
' This function has been removed in compliance with the Windows Sockets 2 specification, revision 2.2.0.
'
' The Windows Socket WSAIsBlocking function is not exported directly by the Ws2_32.dll, and Windows Sockets 2
' applications should not use this function. Windows Sockets 1.1 applications that call this function are
' still supported through the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during calls to
' blocking functions. Instead of using blocking hooks, an applications should use a separate thread (separate
' from the main GUI thread) for network activity.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is TRUE if there is an outstanding blocking function awaiting completion in the
' current thread.  Otherwise, it is FALSE.
' ____________________________________________________________________________________________________________
' BOOL WSAIsBlocking (void);
'=============================================================================================================
Public Declare Function WSAIsBlocking Lib "WSOCK32.DLL" (void) As Long


'=============================================================================================================
' WSAJoinLeaf
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAJoinLeaf function joins a leaf node into a multipoint session, exchanges connect
' data, and specifies needed quality of service based on the supplied FLOWSPEC structures.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s            [in]  Descriptor identifying a multipoint socket.
' name         [in]  Name of the peer to which the socket is to be joined.
' NameLen      [in]  Length of name.
' lpCallerData [in]  Pointer to the user data that is to be transferred to the peer during multipoint session establishment.
' lpCalleeData [out] Pointer to the user data that is to be transferred back from the peer during multipoint session establishment.
' lpSQOS       [in]  Pointer to the FLOWSPEC structures for socket s, one for each direction.
' lpGQOS       [in]  Reserved.
' dwFlags      [in]  Flags to indicate that the socket is acting as a sender, receiver, or both.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSAJoinLeaf returns a value of type SOCKET that is a descriptor for the newly created
' multipoint socket. Otherwise, a value of INVALID_SOCKET is returned, and a specific error code can be
' retrieved by calling WSAGetLastError.
'
' On a blocking socket, the return value indicates success or failure of the join operation.
'
' With a nonblocking socket, successful initiation of a join operation is indicated by a return of a valid
' socket descriptor. Subsequently, an FD_CONNECT indication will be given on the original socket s when the
' join operation completes, either successfully or otherwise. The application must use either WSAAsyncSelect
' or WSAEventSelect with interest registered for the FD_CONNECT event in order to determine when the join
' operation has completed and checks the associated error code to determine the success or failure of the
' operation. The select function cannot be used to determine when the join operation completes.
'
' Also, until the multipoint session join attempt completes all subsequent calls to WSAJoinLeaf on the same
' socket will fail with the error code WSAEALREADY. After the WSAJoinLeaf operation completes successfully,
' a subsequent attempt will usually fail with the error code WSAEISCONN. An exception to the WSAEISCONN rule
' occurs for a c_root socket that allows root-initiated joins. In such a case, another join may be initiated
' after a prior WSAJoinLeaf operation completes.
'
' If the return error code indicates the multipoint session join attempt failed (that is, WSAECONNREFUSED,
' WSAENETUNREACH, WSAETIMEDOUT) the application can call WSAJoinLeaf again for the same socket.
' ____________________________________________________________________________________________________________
' SOCKET WSAJoinLeaf (SOCKET s, const struct sockaddr FAR *name, int namelen, LPWSABUF lpCallerData, LPWSABUF lpCalleeData, LPQOS lpSQOS, LPQOS lpGQOS, DWORD dwFlags);
'=============================================================================================================
Public Declare Function WSAJoinLeaf Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef PeerName As SOCKADDR, ByVal PeerNameLength As Long, ByRef lpCallerData As WSABUF, ByRef lpCalleeData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByVal dwFlags As Long) As Long


'=============================================================================================================
' WSALookupServiceBegin
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSALookupServiceBegin function initiates a client query that is constrained by the
' information contained within a WSAQUERYSET structure. WSALookupServiceBegin only returns a handle,
' which should be used by subsequent calls to WSALookupServiceNext to get the actual results.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpqsRestrictions [in]  Pointer to the search criteria. See the following for details.
' dwControlFlags   [in]  Flag that controls the depth of the search: LUP_DEEP, LUP_CONTAINERS, LUP_NOCONTAINERS, LUP_FLUSHCACHE, LUP_FLUSHPREVIOUS, LUP_NEAREST, LUP_RES_SERVICE, LUP_RETURN_ALIASES, LUP_RETURN_NAME, LUP_RETURN_TYPE, LUP_RETURN_VERSION, LUP_RETURN_COMMENT, LUP_RETURN_ADDR, LUP_RETURN_BLOB, LUP_RETURN_ALL,
' lphLookup        [out] Handle to be used when calling WSALookupServiceNext in order to start retrieving the results set.
'
' As mentioned above, a WSAQUERYSET structure is used as an input parameter to WSALookupBegin in order
' to qualify the query. The following table indicates how the WSAQUERYSET is used to construct a query.
' When a parameter is marked as (Optional) a NULL pointer can be supplied, indicating that the parameter
' will not be used as a search criteria. See section Query-Related Data Structures for additional information.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' .dwSize                   Must be set to sizeof(WSAQUERYSET). This is a versioning mechanism.
' .dwOutputflags            Ignored for queries.
' .LpszServiceInstanceName  (Optional) Referenced string contains service name. The semantics for wildcarding within the string are not defined, but can be supported by certain name space providers.
' .LpServiceClassId         (Required) The GUID corresponding to the service class.
' .LpVersion                (Optional) References desired version number and provides version comparison semantics (that is, version must match exactly, or version must be not less than the value supplied).
' .LpszComment              Ignored for queries.
' .DwNameSpace              Identifier of a single name space in which to constrain the search, or NS_ALL to include all name spaces.  >>IMPORTANT<<  In most instances, applications interested in only a particular transport protocol should constrain their query by address family and protocol rather than by name space. This would allow an application that needs to locate a TCP/IP service, for example, to have its query processed by all available name spaces such as the local hosts file, DNS, and NIS.
' .LpNSProviderId           (Optional) References the GUID of a specific name space provider, and limits the query to this provider only.
' .LpszContext              (Optional) Specifies the starting point of the query in a hierarchical name space.
' .DwNumberOfProtocols      Size of the protocol constraint array, can be zero.
' .LpafpProtocols           (Optional) References an array of AFPROTOCOLS structure. Only services that utilize these protocols will be returned.
' .LpszQueryString          (Optional) Some name spaces (such as whois++) support enriched SQL-like queries that are contained in a simple text string. This parameter is used to specify that string.
' .DwNumberOfCsAddrs        Ignored for queries.
' .LpcsaBuffer              Ignored for queries.
' .LpBlob                   (Optional) This is a pointer to a provider-specific entity.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is returned,
' and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSALookupServiceBegin (LPWSAQUERYSET lpqsRestrictions, DWORD dwControlFlags, LPHANDLE lphLookup);
'=============================================================================================================
Public Declare Function WSALookupServiceBegin Lib "WS2_32.DLL" Alias "WSALookupServiceBeginA" (ByRef lpqsRestrictions As WSAQUERYSET, ByVal dwControlFlags As Long, ByRef lphLookup As Long) As Long


'=============================================================================================================
' WSALookupServiceEnd
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSALookupServiceEnd function is called to free the handle after previous calls to
' WSALookupServiceBegin and WSALookupServiceNext.
'
' If you call WSALookupServiceEnd from another thread while an existing WSALookupServiceNext is blocked,
' the end call will have the same effect as a cancel and will cause the WSALookupServiceNext call to return
' immediately.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hLookup [in] Handle previously obtained by calling WSALookupServiceBegin.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is returned,
' and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSALookupServiceEnd (HANDLE hLookup);
'=============================================================================================================
Public Declare Function WSALookupServiceEnd Lib "WS2_32.DLL" (ByVal hLookup As Long) As Long


'=============================================================================================================
' WSALookupServiceNext
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSALookupServiceNext function is called after obtaining a handle from a previous call
' to WSALookupServiceBegin in order to retrieve the requested service information.
'
' The provider will pass back a WSAQUERYSET structure in the lpqsResults buffer. The client should continue
' to call this function until it returns WSA_E_NOMORE, indicating that all of WSAQUERYSET has been returned.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hLookup          [in]      Handle returned from the previous call to WSALookupServiceBegin.
' dwControlFlags   [in]      Flags to control the next operation. Currently only LUP_FLUSHPREVIOUS is defined as a means to cope with a result set which is too large. If an application does not wish to (or cannot) supply a large enough buffer, setting LUP_FLUSHPREVIOUS instructs the provider to discard the last result set—which was too large—and move on to the next set for this call.
' lpdwBufferLength [in, out] On input, the number of bytes contained in the buffer pointed to by lpqsResults. On output, if the function fails and the error is WSAEFAULT, then it contains the minimum number of bytes to pass for the lpqsResults to retrieve the record.
' lpqsResults      [out]     Pointer to a block of memory, which will contain one result set in a WSAQUERYSET structure on return.
'
' The following table describes how the query results are represented in the WSAQUERYSET structure:
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' .dwSize                   Will be set to sizeof(WSAQUERYSET). This is used as a versioning mechanism.
' .dwOuputFlags             RESULT_IS_ALIAS flag indicates this is an alias result.
' .LpszServiceInstanceName  Referenced string contains service name.
' .LpServiceClassId         The GUID corresponding to the service class.
' .LpVersion                References version number of the particular service instance.
' .LpszComment              Optional comment string supplied by service instance.
' .DwNameSpace              Name space in which the service instance was found.
' .LpNSProviderId           Identifies the specific name space provider that supplied this query result.
' .LpszContext              Specifies the context point in a hierarchical name space at which the service is located.
' .DwNumberOfProtocols      Undefined for results.
' .lpafpProtocols           Undefined for results, all needed protocol information is in the CSADDR_INFO structures.
' .LpszQueryString          When dwControlFlags includes LUP_RETURN_QUERY_STRING, this parameter returns the unparsed remainder of the lpszServiceInstanceName specified in the original query. For example, in a name space that identifies services by hierarchical names that specify a host name and a file path within that host, the address returned might be the host address and the unparsed remainder might be the file path. If the lpszServiceInstanceName is fully parsed and LUP_RETURN_QUERY_STRING is used, this parameter is NULL or points to a zero-length string.
' .DwNumberOfCsAddrs        Indicates the number of elements in the array of CSADDR_INFO structures.
' .LpcsaBuffer              A pointer to an array of CSADDR_INFO structures, with one complete transport address contained within each element.
' .LpBlob                   (Optional) This is a pointer to a provider-specific entity.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is returned,
' and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSALookupServiceNext (HANDLE hLookup, DWORD dwControlFlags, LPDWORD lpdwBufferLength, LPWSAQUERYSET lpqsResults);
'=============================================================================================================
Public Declare Function WSALookupServiceNext Lib "WS2_32.DLL" Alias "WSALookupServiceNextA" (ByVal hLookup As Long, ByVal dwControlFlags As Long, ByRef lpdwBufferLength As Long, ByRef lpqsResults As WSAQUERYSET) As Long


'=============================================================================================================
' WSANtohl
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSANtohl function converts a u_long from network byte order to host byte order.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s          [in]  Descriptor identifying a socket.
' netLong    [in]  32-bit number in network byte order.
' lphostlong [out] Pointer to a 32-bit number in host byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSANtohl returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSANtohl (SOCKET s, u_long netlong, u_long FAR * lphostlong);
'=============================================================================================================
Public Declare Function WSANtohl Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal NetLong As Long, ByRef lpHostLong As Long) As Long


'=============================================================================================================
' WSANtohs
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSANtohs function converts a u_short from network byte order to host byte order.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s           [in]  Descriptor identifying a socket.
' NetShort    [in]  16-bit number in network byte order.
' lphostshort [out] Pointer to a 16-bit number in host byte order.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSANtohs returns zero. Otherwise, a value of SOCKET_ERROR is returned, and a
' specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSANtohs (SOCKET s, u_short netshort, u_short FAR * lphostshort);
'=============================================================================================================
Public Declare Function WSANtohs Lib "WS2_32.DLL" (ByVal hSocket As Long, ByVal NetShort As Integer, ByRef lpHostShort As Integer) As Long


'=============================================================================================================
' WSAProviderConfigChange
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAProviderConfigChange function notifies the application when the provider
' configuration is changed.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpNotificationHandle [in, out] Pointer to notification handle. If the notification handle is set to NULL (the handle value not the pointer itself), this function returns a notification handle in the location pointed to by lpNotificationHandle.
' lpOverlapped         [in]      Pointer to a WSAOVERLAPPED structure.
' lpCompletionRoutine  [in]      Pointer to the completion routine called when the provider change notification is received.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs the WSAProviderConfigChange returns 0. Otherwise, a value of SOCKET_ERROR is returned
' and a specific error code may be retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING
' indicates that the overlapped operation has been successfully initiated and that completion (and thus
' change event) will be indicated at a later time.
' ____________________________________________________________________________________________________________
' int WSAProviderConfigChange (LPHANDLE lpNotificationHandle, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSAProviderConfigChange Lib "WS2_32.DLL" (ByRef lpNotificationHandle As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSARecv
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSARecv function receives data from a connected socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                    [in]      Descriptor identifying a connected socket.
' lpBuffers            [in, out] Pointer to an array of WSABUF structures. Each WSABUF structure contains a pointer to a buffer and the length of the buffer.
' dwBufferCount        [in]      Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesRecvd [out]     Pointer to the number of bytes received by this call if the receive operation completes immediately.
' lpFlags              [in, out] Pointer to flags: MSG_PEEK, MSG_OOB, MSG_PARTIAL
' lpOverlapped         [in]      Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped sockets).
' lpCompletionRoutine  [in]      Pointer to the completion routine called when the receive operation has been completed (ignored for nonoverlapped sockets).
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs and the receive operation has completed immediately, WSARecv returns zero. In this case,
' the completion routine will have already been scheduled to be called once the calling thread is in the
' alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved
' by calling WSAGetLastError. The error code WSA_IO_PENDING indicates that the overlapped operation has been
' successfully initiated and that completion will be indicated at a later time. Any other error code indicates
' that the overlapped operation was not successfully initiated and no completion indication will occur.
' ____________________________________________________________________________________________________________
' int WSARecv (SOCKET s, LPWSABUF lpBuffers, DWORD dwBufferCount, LPDWORD lpNumberOfBytesRecvd, LPDWORD lpFlags, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSARecv Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesRecvd As Long, ByRef lpFlags As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSARecvDisconnect
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSARecvDisconnect function terminates reception on a socket, and retrieves the
' disconnect data if the socket is connection oriented.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                       [in]  Descriptor identifying a socket.
' lpInboundDisconnectData [out] Pointer to the incoming disconnect data.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSARecvDisconnect returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSARecvDisconnect (SOCKET s, LPWSABUF lpInboundDisconnectData);
'=============================================================================================================
Public Declare Function WSARecvDisconnect Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpInboundDisconnectData As WSABUF) As Long


'=============================================================================================================
' WSARecvEx
'
' Minimum Availability : Requires Windows Sockets 1.1  (Not supported on Windows 95)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSARecvEx function is identical to the recv function, except that the flags parameter
' is an [in, out] parameter. When a partial message is received while using datagram protocol, the
' MSG_PARTIAL bit is set in the flags parameter on return from the function.
'
' Note :
' ¯¯¯¯¯¯
' The Windows Sockets WSARecvEx function is a Microsoft-specific extension to the Windows Sockets
' specification. For more information, see Microsoft Extensions and Windows Sockets 2.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s     [in]      Descriptor identifying a connected socket.
' buf   [out]     Buffer for the incoming data.
' len   [in]      Length of buf.
' Flags [in, out] Indicator specifying whether the message is fully or partially received for datagram sockets.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSARecvEx returns the number of bytes received. If the connection has been closed,
' it returns zero. Additionally, if a partial message was received, the MSG_PARTIAL bit is set in the
' flags parameter. If a complete message was received, MSG_PARTIAL is not set in flags.
'
' Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be retrieved by calling
' WSAGetLastError.
'
' Important : For a stream oriented-transport protocol, MSG_PARTIAL is never set on return from WSARecvEx.
' This function behaves identically to the Windows Sockets recv function for stream-transport protocols.
' ____________________________________________________________________________________________________________
' int WSARecvEx (SOCKET s, char FAR *buf, int len, int *flags);
'=============================================================================================================
Public Declare Function WSARecvEx Lib "WSOCK32.DLL" (ByVal hSocket As Long, ByRef Buffer As Any, ByVal BufferLength As Long, ByRef Flags As Long) As Long


'=============================================================================================================
' WSARecvFrom
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSARecvFrom function receives a datagram and stores the source address.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                    [in]      Descriptor identifying a socket.
' lpBuffers            [in, out] Pointer to an array of WSABUF structures. Each WSABUF structure contains a pointer to a buffer and the length of the buffer.
' dwBufferCount        [in]      Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesRecvd [out]     Pointer to the number of bytes received by this call if the recv operation completes immediately.
' lpFlags              [in, out] Pointer to flags.
' lpFrom               [out]     Optional pointer to a buffer that will hold the source address upon the completion of the overlapped operation.
' lpFromlen            [in, out] Pointer to the size of the from buffer, required only if lpFrom is specified.
' lpOverlapped         [in]      Pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped sockets).
' lpCompletionRoutine  [in]      Pointer to the completion routine called when the recv operation has been completed (ignored for nonoverlapped sockets).
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs and the receive operation has completed immediately, WSARecvFrom returns zero. In this
' case, the completion routine will have already been scheduled to be called once the calling thread is in
' the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be
' retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates that the overlapped operation
' has been successfully initiated and that completion will be indicated at a later time. Any other error code
' indicates that the overlapped operation was not successfully initiated and no completion indication will
' occur.
' ____________________________________________________________________________________________________________
' int WSARecvFrom (SOCKET s, LPWSABUF lpBuffers, DWORD dwBufferCount, LPDWORD lpNumberOfBytesRecvd, LPDWORD lpFlags, struct sockaddr FAR *lpFrom, LPINT lpFromlen, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSARecvFrom Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesRecvd As Long, ByRef lpFlags As Long, ByRef lpFrom As SOCKADDR, ByRef lpFromlen As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSARemoveServiceClass
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSARemoveServiceClass function permanently removes from the registry service class schema.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpServiceClassId [in] Pointer to the GUID for the service class you want to remove.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is zero if the operation was successful. Otherwise, the value SOCKET_ERROR is returned,
' and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSARemoveServiceClass (LPGUID lpServiceClassId);
'=============================================================================================================
Public Declare Function WSARemoveServiceClass Lib "WS2_32.DLL" (ByRef lpServiceClassId As GUID) As Long


'=============================================================================================================
' WSAResetEvent
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAResetEvent function resets the state of the specified event object to nonsignaled.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEvent [in] Handle that identifies an open event object handle.
'
' Return:
' ¯¯¯¯¯¯¯
' If the WSAResetEvent function succeeds, the return value is TRUE. If the function fails, the return
' value is FALSE. To get extended error information, call WSAGetLastError.
' ____________________________________________________________________________________________________________
' BOOL WSAResetEvent (WSAEVENT hEvent);
'=============================================================================================================
Public Declare Function WSAResetEvent Lib "WS2_32.DLL" (ByRef hEvent As Long) As Long


'=============================================================================================================
' WSASend
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASend function sends data on a connected socket.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                   [in]  Descriptor identifying a connected socket.
' lpBuffers           [in]  Pointer to an array of WSABUF structures. Each WSABUF structure contains a pointer to a buffer and the length of the buffer. This array must remain valid for the duration of the send operation.
' dwBufferCount       [in]  Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesSent [out] Pointer to the number of bytes sent by this call if the I/O operation completes immediately.
' dwFlags             [in]  Flags used to modify the behavior of the WSASend function call. See Using dwFlags in the Remarks section for more information.
' lpOverlapped        [in]  Pointer to a WSAOVERLAPPED structure. This parameter is ignored for nonoverlapped sockets.
' lpCompletionRoutine [in]  Pointer to the completion routine called when the send operation has been completed. This parameter is ignored for nonoverlapped sockets.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs and the send operation has completed immediately, WSASend returns zero. In this case,
' the completion routine will have already been scheduled to be called once the calling thread is in the
' alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be
' retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates that the overlapped
' operation has been successfully initiated and that completion will be indicated at a later time. Any
' other error code indicates that the overlapped operation was not successfully initiated and no completion
' indication will occur.
' ____________________________________________________________________________________________________________
' int WSASend (SOCKET s, LPWSABUF lpBuffers, DWORD dwBufferCount, LPDWORD lpNumberOfBytesSent, DWORD dwFlags, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSASend Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesSent As Long, ByVal dwFlags As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSASendDisconnect
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASendDisconnect function initiates termination of the connection for the socket
' and sends disconnect data.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                        [in] Descriptor identifying a socket.
' lpOutboundDisconnectData [in] Pointer to the outgoing disconnect data.
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSASendDisconnect returns zero. Otherwise, a value of SOCKET_ERROR is returned,
' and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' int WSASendDisconnect (SOCKET s, LPWSABUF lpOutboundDisconnectData);
'=============================================================================================================
Public Declare Function WSASendDisconnect Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpOutboundDisconnectData As WSABUF) As Long


'=============================================================================================================
' WSASendTo
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASendTo function sends data to a specific destination, using overlapped I/O
' where applicable.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' s                   [in]  Descriptor identifying a (possibly connected) socket.
' lpBuffers           [in]  Pointer to an array of WSABUF structures. Each WSABUF structure contains a pointer to a buffer and the length of the buffer. This array must remain valid for the duration of the send operation.
' dwBufferCount       [in]  Number of WSABUF structures in the lpBuffers array.
' lpNumberOfBytesSent [out] Pointer to the number of bytes sent by this call if the I/O operation completes immediately.
' dwFlags             [in]  Indicator specifying the way in which the call is made.
' lpTo                [in]  Optional pointer to the address of the target socket.
' iToLen              [in]  Size of the address in lpTo.
' lpOverlapped        [in]  A pointer to a WSAOVERLAPPED structure (ignored for nonoverlapped sockets).
' lpCompletionRoutine [in]  Pointer to the completion routine called when the send operation has been completed (ignored for nonoverlapped sockets).
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs and the send operation has completed immediately, WSASendTo returns zero. In this
' case, the completion routine will have already been scheduled to be called once the calling thread is in
' the alertable state. Otherwise, a value of SOCKET_ERROR is returned, and a specific error code can be
' retrieved by calling WSAGetLastError. The error code WSA_IO_PENDING indicates that the overlapped operation
' has been successfully initiated and that completion will be indicated at a later time. Any other error code
' indicates that the overlapped operation was not successfully initiated and no completion indication will
' occur.
' ____________________________________________________________________________________________________________
' int WSASendTo (SOCKET s, LPWSABUF lpBuffers, DWORD dwBufferCount, LPDWORD lpNumberOfBytesSent, DWORD dwFlags, const struct sockaddr FAR *lpTo, int iToLen, LPWSAOVERLAPPED lpOverlapped, LPWSAOVERLAPPED_COMPLETION_ROUTINE lpCompletionRoutine);
'=============================================================================================================
Public Declare Function WSASendTo Lib "WS2_32.DLL" (ByVal hSocket As Long, ByRef lpBuffers As WSABUF, ByVal dwBufferCount As Long, ByRef lpNumberOfBytesSent As Long, ByVal dwFlags As Long, ByRef lpTo As SOCKADDR, ByVal iToLen As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal lpCompletionRoutine As Long) As Long


'=============================================================================================================
' WSASetBlockingHook
'
' Minimum Availability : Requires Windows Sockets 1.1 (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' Establish an application-supplied blocking hook function.
'
' This function has been removed in compliance with the Windows Sockets 2 specification, revision 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll, and Windows Sockets 2 applications should not
' use this function. Windows Sockets 1.1 applications that call this function are still supported through
' the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during calls to
' blocking functions. Instead of using blocking hooks, an application should use a separate thread
' separate from the main GUI thread) for network activity.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpBlockFunc [in] A pointer to the procedure instance address of the blocking function to be installed.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is a pointer to the procedure-instance of the previously installed blocking function.
' The application or library that calls the WSASetBlockingHook() function should save this return value
' so that it can be restored if necessary.  (If "nesting" is not important, the application may simply
' discard the value returned by WSASetBlockingHook() and eventually use WSAUnhookBlockingHook() to restore
' the default mechanism.)  If the operation fails, a NULL pointer is returned, and a specific error number
' may be retrieved by calling WSAGetLastError().
' ____________________________________________________________________________________________________________
' FARPROC WSASetBlockingHook (FARPROC lpBlockFunc);
'=============================================================================================================
Public Declare Function WSASetBlockingHook Lib "WSOCK32.DLL" (ByVal lpBlockFunc As Long) As Long


'=============================================================================================================
' WSASetEvent
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASetEvent function sets the state of the specified event object to signaled.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' hEvent [in] Handle that identifies an open event object.
'
' Return:
' ¯¯¯¯¯¯¯
' If the function succeeds, the return value is TRUE.  If the function fails, the return value is FALSE.
' To get extended error information, call WSAGetLastError.
' ____________________________________________________________________________________________________________
' BOOL WSASetEvent (WSAEVENT hEvent);
'=============================================================================================================
Public Declare Function WSASetEvent Lib "WS2_32.DLL" (ByVal hEvent As Long) As Long


'=============================================================================================================
' WSASetLastError
'
' Minimum Availability : Requires Windows Sockets 1.1
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASetLastError function sets the error code that can be retrieved through the
' WSAGetLastError function.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' iError [in] Integer that specifies the error code to be returned by a subsequent WSAGetLastError call.
'
' Return:
' ¯¯¯¯¯¯¯
' None
' ____________________________________________________________________________________________________________
' void WSASetLastError (int iError);
'=============================================================================================================
Public Declare Sub WSASetLastError Lib "WSOCK32.DLL" (ByVal iError As Long)


'=============================================================================================================
' WSASetService
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASetService function registers or removes from the registry a service instance
' within one or more name spaces. This function can be used to affect a specific name space provider,
' all providers associated with a specific name space, or all providers across all name spaces.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' lpqsRegInfo    [in] Pointer to the service information for registration or deregistration.
' essOperation   [in] Enumeration whose values include the following: RNRSERVICE_REGISTER, RNRSERVICE_DEREGISTER, RNRSERVICE_DELETE
' dwControlFlags [in] Meaning of dwControlFlags is dependent on the following values: SERVICE_MULTIPLE
'
' ____________________________________________________________________________________________________________
' The following table describes how service property data is represented in a WSAQUERYSET structure.
' Fields labeled as (Optional) can be supplied with a NULL pointer.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' .dwSize                   Must be set to sizeof (WSAQUERYSET). This is a versioning mechanism.
' .dwOutputFlags            Not applicable and ignored.
' .LpszServiceInstanceName  Referenced string contains the service instance name.
' .LpServiceClassId         The GUID corresponding to this service class.
' .LpVersion                (Optional) Supplies service instance version number.
' .LpszComment              (Optional) An optional comment string.
' .DwNameSpace              See table that follows.
' .LpNSProviderId           See table that follows.
' .LpszContext              (Optional) Specifies the starting point of the query in a hierarchical name space.
' .DwNumberOfProtocols      Ignored.
' .LpafpProtocols           Ignored.
' .LpszQueryString          Ignored.
' .DwNumberOfCsAddrs        The number of elements in the array of CSADDR_INFO structures referenced by lpcsaBuffer.
' .LpcsaBuffer              A pointer to an array of CSADDR_INFO structures that contain the address(es) that the service is listening on.
' .LpBlob                   (Optional) This is a pointer to a provider-specific entity.
'
' ____________________________________________________________________________________________________________
' As illustrated in the following, the combination of the dwNameSpace and lpNSProviderId parameters determine
' that name space providers are affected by this function.
' ————————————————————————————————————————————————————————————————————————————————————————————————————————————
' DwNameSpace                    lpNSProviderId     Scope of impact
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' Ignored                        Non-NULL           The specified name-space provider.
' A valid name space identifier  NULL               All name-space providers that support the indicated name space.
' NS_ALL                         NULL               All name-space providers.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value for WSASetService is zero if the operation was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSASetService (LPWSAQUERYSET lpqsRegInfo, WSAESETSERVICEOP essOperation, DWORD dwControlFlags);
'=============================================================================================================
Public Declare Function WSASetService Lib "WS2_32.DLL" Alias "WSASetServiceA" (ByRef lpqsRegInfo As WSAQUERYSET, ByVal essOperation As Long, ByVal dwControlFlags As Long) As Long


'=============================================================================================================
' WSASocket
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSASocket function creates a socket that is bound to a specific transport-service
' provider.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' af             [in] Address family specification.
' type           [in] Type specification for the new socket.
' Protocol       [in] Protocol to be used with the socket that is specific to the indicated address family.
' lpProtocolInfo [in] Pointer to a WSAPROTOCOL_INFO structure that defines the characteristics of the socket to be created.
' g              [in] Reserved.
' dwFlags        [in] Flag that specifies the socket attribute. (See WSA_FLAG_* Constants)
'
' Return:
' ¯¯¯¯¯¯¯
' If no error occurs, WSASocket returns a descriptor referencing the new socket. Otherwise, a value of
' INVALID_SOCKET is returned, and a specific error code can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' SOCKET WSASocket (int af, int type, int protocol, LPWSAPROTOCOL_INFO lpProtocolInfo, GROUP g, DWORD dwFlags);
'=============================================================================================================
Public Declare Function WSASocket Lib "WS2_32.DLL" Alias "WSASocketA" (ByVal AddressFamily As Long, ByVal SocketType As Long, ByVal Protocol As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFO, ByVal Reserved As Long, ByVal dwFlags As Long) As Long


'=============================================================================================================
' WSAStartup
'
' Minimum Availability : Requires Windows Sockets 1.1
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAStartup function initiates use of Ws2_32.dll by a process.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' wVersionRequested [in]  Highest version of Windows Sockets support that the caller can use. The high-order byte specifies the minor version (revision) number; the low-order byte specifies the major version number.
' lpWSAData         [out] Pointer to the WSADATA data structure that is to receive details of the Windows Sockets implementation.
'
' ____________________________________________________________________________________________________________
' In order to support future Windows Sockets implementations and applications that can have functionality
' differences from the current version of Windows Sockets, a negotiation takes place in WSAStartup. The caller
' of WSAStartup and the Ws2_32.dll indicate to each other the highest version that they can support, and each
' confirms that the other's highest version is acceptable. Upon entry to WSAStartup, the Ws2_32.dll examines
' the version requested by the application. If this version is equal to or higher than the lowest version
' supported by the DLL, the call succeeds and the DLL returns in wHighVersion the highest version it supports
' and in wVersion the minimum of its high version and wVersionRequested. The Ws2_32.dll then assumes that the
' application will use wVersion. If the wVersion parameter of the WSADATA structure is unacceptable to the
' caller, it should call WSACleanup and either search for another Ws2_32.dll or fail to initialize.
'
' It is legal and possible for an application written to this version of the specification to successfully
' negotiate a higher version number version. In that case, the application is only guaranteed access to
' higher-version functionality that fits within the syntax defined in this version, such as new Ioctl codes
' and new behavior of existing functions. New functions may be inaccessible. To get full access to the new
' syntax of a future version, the application must fully conform to that future version, such as compiling
' against a new header file, linking to a new library, or other special cases.
'
' This negotiation allows both a Ws2_32.dll and a Windows Sockets application to support a range of Windows
' Sockets versions. An application can use Ws2_32.dll if there is any overlap in the version ranges. The
' following table shows how WSAStartup works with different applications and Ws2_32.dll versions.
' ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' App         DLL        wVersion              wHigh      End
' Versions    Versions   Requested  wVersion   Version    Result
' ------------------------------------------------------------------------------------------------------------
' 1.1         1.1        1.1        1.1        1.1        Use 1.1
' 1.0, 1.1    1.0        1.1        1.0        1.0        Use 1.0
' 1.0         1.0,1.1    1.0        1.0        1.1        Use 1.0
' 1.1         1.0,1.1    1.1        1.1        1.1        Use 1.1
' 1.1         1.0        1.1        1.0        1.0        Application Fails
' 1.0         1.1        1.0        ---        ---        Error (WSAVERNOTSUPPORTED)
' 1.0, 1.1    1.0, 1.1   1.1        1.1        1.1        Use 1.1
' 1.1, 2.0    1.1        2.0        1.1        1.1        Use 1.1
' 2.0         2.0        2.0        2.0        2.0        Use 2.0
' ————————————————————————————————————————————————————————————————————————————————————————————————————————————
'
' Return:
' ¯¯¯¯¯¯¯
' The WSAStartup function returns zero if successful. Otherwise, it returns one of the error codes listed
' in the following.
'
' An application cannot call WSAGetLastError to determine the error code as is normally done in Windows
' Sockets if WSAStartup fails. The Ws2_32.dll will not have been loaded in the case of a failure so the
' client data area where the last error information is stored could not be established.
' ____________________________________________________________________________________________________________
' int WSAStartup (WORD wVersionRequested, LPWSADATA lpWSAData);
'=============================================================================================================
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA) As Long


'=============================================================================================================
' WSAStringToAddress
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAStringToAddress function converts a numeric string to a SOCKADDR structure,
' suitable for passing to Windows Sockets routines that take such a structure.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' AddressString   [in]      Pointer to the zero-terminated human-readable numeric string to convert.
' AddressFamily   [in]      Address family to which the string belongs.
' lpProtocolInfo  [in]      (optional) The WSAPROTOCOL_INFO structure associated with the provider to be used. If this is NULL, the call is routed to the provider of the first protocol supporting the indicated AddressFamily.
' lpAddress       [out]     Buffer that is filled with a single SOCKADDR.
' lpAddressLength [in, out] Length of the Address buffer. Returns the size of the resultant SOCKADDR structure. If the supplied buffer is not large enough, the function fails with a specific error of WSAEFAULT and this parameter is updated with the required size in bytes.
'
' Return:
' ¯¯¯¯¯¯¯
' The return value for WSAStringToAddress is zero if the operation was successful. Otherwise, the value
' SOCKET_ERROR is returned, and a specific error number can be retrieved by calling WSAGetLastError.
' ____________________________________________________________________________________________________________
' INT WSAStringToAddress (LPTSTR AddressString, INT AddressFamily, LPWSAPROTOCOL_INFO lpProtocolInfo, LPSOCKADDR lpAddress, LPINT lpAddressLength);
'=============================================================================================================
Public Declare Function WSAStringToAddress Lib "WS2_32.DLL" Alias "WSAStringToAddressA" (ByVal AddressString As String, ByVal AddressFamily As Long, ByRef lpProtocolInfo As WSAPROTOCOL_INFO, ByRef lpAddress As SOCKADDR, ByRef lpAddressLength As Long) As Long


'=============================================================================================================
' WSAUnhookBlockingHook
'
' Minimum Availability : Requires Windows Sockets 1.1 (Obsolete for Windows Sockets 2.0)
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' Restores the default blocking hook function.
'
' This function has been removed in compliance with the Windows Sockets 2 specification, revision 2.2.0.
'
' The function is not exported directly by the Ws2_32.dll, and Windows Sockets 2 applications should not
' use this function. Windows Sockets 1.1 applications that call this function are still supported through
' the Winsock.dll and Wsock32.dll.
'
' Blocking hooks are generally used to keep a single-threaded GUI application responsive during calls to
' blocking functions. Instead of using blocking hooks, an application should use a separate thread
' (separate from the main GUI thread) for network activity.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' None
'
' Return:
' ¯¯¯¯¯¯¯
' The return value is 0 if the operation was successful.  Otherwise the value SOCKET_ERROR is returned,
' and a specific error number may be retrieved by calling WSAGetLastError().
' ____________________________________________________________________________________________________________
' int WSAUnhookBlockingHook (void);
'=============================================================================================================
Public Declare Function WSAUnhookBlockingHook Lib "WSOCK32.DLL" () As Long


'=============================================================================================================
' WSAWaitForMultipleEvents
'
' Minimum Availability : Requires Windows Sockets 2.0
'
' Description:
' ¯¯¯¯¯¯¯¯¯¯¯¯
' The Windows Sockets WSAWaitForMultipleEvents function returns either when one or all of the specified
' event objects are in the signaled state, or when the time-out interval expires.
'
' Parameters:
' ¯¯¯¯¯¯¯¯¯¯¯
' cEvents    [in] Indicator specifying the number of event object handles in the array pointed to by lphEvents. The maximum number of event object handles is WSA_MAXIMUM_WAIT_EVENTS. One or more events must be specified.
' lphEvents  [in] Pointer to an array of event object handles.
' fWaitAll   [in] Indicator specifying the wait type. If TRUE, the function returns when all event objects in the lphEvents array are signaled at the same time. If FALSE, the function returns when any one of the event objects is signaled. In the latter case, the return value indicates the event object whose state caused the function to return.
' dwTimeout  [in] Indicator specifying the time-out interval, in milliseconds. The function returns if the interval expires, even if conditions specified by the fWaitAll parameter are not satisfied. If dwTimeout is zero, the function tests the state of the specified event objects and returns immediately. If dwTimeout is WSA_INFINITE, the function's time-out interval never expires.
' fAlertable [in] Indicator specifying whether the function returns when the system queues an I/O completion routine for execution by the calling thread. If TRUE, the completion routine is executed and the function returns. If FALSE, the completion routine is not executed when the function returns.
' ____________________________________________________________________________________________________________
' Return:
' ¯¯¯¯¯¯¯
' If the WSAWaitForMultipleEvents function succeeds, the return value indicates the event object that caused
' the function to return.  If the function fails, the return value is WSA_WAIT_FAILED. To get extended error
' information, call WSAGetLastError.
'
' The return value upon success is one of the following values:
' -------------------------------------------------------------
'  Value:                           Meaning
' -------------------------------------------------------------
'  WSA_WAIT_EVENT_0 to
' (WSA_WAIT_EVENT_0 + cEvents - 1)  If fWaitAll is TRUE, the return value indicates that the state of all
'                                   specified event objects is signaled. If fWaitAll is FALSE, the return
'                                   value minus WSA_WAIT_EVENT_0 indicates the lphEvents array index of the
'                                   object that satisfied the wait.
'
'  WSA_WAIT_IO_COMPLETION           One or more I/O completion routines are queued for execution.
'
'  WSA_WAIT_TIMEOUT                 The time-out interval elapsed and the conditions specified by the
'                                   fWaitAll parameter are not satisfied.
' ____________________________________________________________________________________________________________
' DWORD WSAWaitForMultipleEvents (DWORD cEvents, const WSAEVENT FAR *lphEvents, BOOL fWaitAll, DWORD dwTimeout, BOOL fAlertable);
'=============================================================================================================
Public Declare Function WSAWaitForMultipleEvents Lib "WS2_32.DLL" (ByVal cEvents As Long, ByVal lphEvents As Long, ByVal fWaitAll As Long, ByVal dwTimeout As Long, ByVal fAlertable As Long) As Long





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




'-------------------------------------------------------------------------------------------------------------
' ¤ WSAAccept.lpfnCondition = Procedure instance address of an optional-condition function furnished by
' Windows Sockets. This function is used in the accept or reject decision based on the caller information
' passed in as parameters.
'_____________________________________________________________________________________________________________
'-------------------------------------------------------------------------------------------------------------
' The prototype of the condition function is as follows:
'-------------------------------------------------------------------------------------------------------------
' int CALLBACK ConditionFunc(
'   IN      LPWSABUF   lpCallerId,    // The lpCallerId parameter points to a WSABUF structure that contains the address of the connecting entity, where its len parameter is the length of the buffer in bytes, and its buf parameter is a pointer to the buffer.  The buf portion of the WSABUF pointed to by lpCallerId points to a SOCKADDR. The SOCKADDR structure is interpreted according to its address family (typically by casting the SOCKADDR to some type specific to the address family).
'   IN      LPWSABUF   lpCallerData,  // The lpCallerData is a value parameter that contains any user data. The information in these parameters is sent along with the connection request. If no caller identification or caller data is available, the corresponding parameters will be NULL. Many network protocols do not support connect-time caller data. Most conventional network protocols can be expected to support caller identifier information at connection-request time.
'   IN OUT  LPQOS      lpSQOS,        // The lpSQOS parameter references the FLOWSPEC structures for socket s specified by the caller, one for each direction, followed by any additional provider-specific parameters. The sending or receiving flow specification values will be ignored as appropriate for any unidirectional sockets. A NULL value for indicates that there is no caller supplied QOS and that no negotiation is possible. A non-NULL lpSQOS pointer indicates that a QOS negotiation is to occur or that the provider is prepared to accept the QOS request without negotiation.
'   IN OUT  LPQOS      lpGQOS,        // The lpGQOS parameter is reserved, and should be NULL.
'   IN      LPWSABUF   lpCalleeId,    // The lpCalleeId is a value parameter that contains the local address of the connected entity. The buf portion of the WSABUF pointed to by lpCalleeId points to a SOCKADDR. The SOCKADDR structure is interpreted according to its address family (typically by casting the SOCKADDR to some type specific to the address family).
'   OUT     LPWSABUF   lpCalleeData,  // The lpCalleeData is a result parameter used by the condition function to supply user data back to the connecting entity. The lpCalleeData->len initially contains the length of the buffer allocated by the service provider and pointed to by lpCalleeData->buf. A value of zero means passing user data back to the caller is not supported. The condition function should copy up to lpCalleeData->len bytes of data into lpCalleeData->buf, and then update lpCalleeData->len to indicate the actual number of bytes transferred. If no user data is to be passed back to the caller, the condition function should set lpCalleeData->len to zero. The format of all address and user data is specific to the address family to which the socket belongs.
'   OUT     GROUP FAR  *g,            // (GROUP = unsigned int)
'   IN      DWORD      dwCallbackData // The dwCallbackData parameter value passed to the condition function is the value passed as the dwCallbackData parameter in the original WSAAccept call. This value is interpreted only by the Windows Socket version 2 client. This allows a client to pass some context information from the WSAAccept call site through to the condition function. This also provides the condition function with any additional information required to determine whether to accept the connection or not. A typical usage is to pass a (suitably cast) pointer to a data structure containing references to application-defined objects with which this socket is associated.
' );
'-------------------------------------------------------------------------------------------------------------
Public Function ConditionFunc(ByRef lpCallerId As WSABUF, ByRef lpCallerData As WSABUF, ByRef lpSQOS As FLOWSPEC, ByVal Reserved As Long, ByRef lpCalleeId As WSABUF, ByRef lpCalleeData As WSABUF, ByRef Group As Long, ByVal dwCallbackData As Long) As Long
  
  DoEvents
  
End Function


'-------------------------------------------------------------------------------------------------------------
' ¤ WSAIoctl.lpCompletionRoutine = Pointer to the completion routine called when the operation has been
' completed (ignored for nonoverlapped sockets).
'
' If the lpCompletionRoutine parameter is NULL, the hEvent parameter of lpOverlapped is signaled when the
' overlapped operation completes if it contains a valid event object handle. An application can use
' WSAWaitForMultipleEvents or WSAGetOverlappedResult to wait or poll on the event object.
'
' If lpCompletionRoutine is not NULL, the hEvent parameter is ignored and can be used by the application to
' pass context information to the completion routine. A caller that passes a non-NULL lpCompletionRoutine
' and later calls WSAGetOverlappedResult for the same overlapped I/O request may not set the fWait parameter
' for that invocation of WSAGetOverlappedResult to TRUE. In this case, the usage of the hEvent parameter is
' undefined, and attempting to wait on the hEvent parameter would produce unpredictable results.
'-------------------------------------------------------------------------------------------------------------
' ¤ WSAProviderConfigChange.lpCompletionRoutine = Pointer to the completion routine called when the provider
' change notification is received.
'-------------------------------------------------------------------------------------------------------------
' ¤ WSARecv.lpCompletionRoutine =  Pointer to the completion routine called when the receive operation has
' been completed (ignored for nonoverlapped sockets).
'-------------------------------------------------------------------------------------------------------------
' ¤ WSARecvFrom.lpCompletionRoutine = Pointer to the completion routine called when the recv operation has
' been completed (ignored for nonoverlapped sockets).
'-------------------------------------------------------------------------------------------------------------
' ¤ WSASend.lpCompletionRoutine = Pointer to the completion routine called when the send operation has
' been completed. This parameter is ignored for nonoverlapped sockets.
'-------------------------------------------------------------------------------------------------------------
' ¤ WSASendTo.lpCompletionRoutine = Pointer to the completion routine called when the send operation has
' been completed (ignored for nonoverlapped sockets).
'_____________________________________________________________________________________________________________
'-------------------------------------------------------------------------------------------------------------
' The prototype of the completion routine is as follows::
'-------------------------------------------------------------------------------------------------------------
' void CALLBACK CompletionRoutine(
'   IN  DWORD           dwError,       // The dwError parameter specifies the completion status for the overlapped operation as indicated by lpOverlapped
'   IN  DWORD           cbTransferred, // The cbTransferred parameter specifies the number of bytes returned
'   IN  LPWSAOVERLAPPED lpOverlapped,  //
'   IN  DWORD           dwFlags        // Currently, there are no flag values defined and dwFlags will be zero
' );
'-------------------------------------------------------------------------------------------------------------
Public Sub CompletionRoutine(ByVal dwError As Long, ByVal cbTransferred As Long, ByRef lpOverlapped As WSAOVERLAPPED, ByVal dwFlags As Long)
  
  DoEvents
  
End Sub


'-------------------------------------------------------------------------------------------------------------
' WSASetBlockingHook.lpBlockFunc - A pointer to the procedure instance address of the blocking function to
' be installed.
'
' The simplest such blocking hook function would simply return FALSE.  If a service provider depends on
' messages for internal operation it may execute PeekMessage(hMyWnd...) before executing the application
' blocking hook so it can get its messages without affecting the rest of the system.
'_____________________________________________________________________________________________________________
'-------------------------------------------------------------------------------------------------------------
' The prototype of the completion routine is as follows::
'-------------------------------------------------------------------------------------------------------------
' int BlockingHook (void);
'-------------------------------------------------------------------------------------------------------------
Public Function BlockingHook() As Long
  
  Dim lpMSG     As MSG
  Dim ReturnVal As Long
  
  ' Get the next message (if any)
  ReturnVal = PeekMessage(lpMSG, 0, 0, 0, PM_REMOVE)
  
  ' If we got a message, process it
  If ReturnVal <> 0 Then ' TRUE if we got a message
    TranslateMessage lpMSG
    DispatchMessage lpMSG
  End If
  
  ' ReturnValurn the result
  BlockingHook = ReturnVal
  
End Function




'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX




' This function takes the error number passed to it via the "LastErrorNum" perameter and gets the error message
' that that number represents.  If no error number is specified, then check to see if one occured.
Public Function GetLastWin32Err(Optional ByVal LastErrorNum As Long, _
                                Optional ByVal LastAPICalled As String = "last", _
                                Optional ByRef Return_ErrNum As Long, _
                                Optional ByRef Return_ErrDesc As String, _
                                Optional ByVal ShowErrorMsg As Boolean = True) As Boolean
On Error Resume Next
  
  ' Clear the return values first
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' If no error message is specified then check for one
  If LastErrorNum = 0 Then
    LastErrorNum = GetLastError
    If LastErrorNum = 0 Then
      Exit Function
    End If
  End If
  
  ' Allocate a buffer for the error description
  Return_ErrDesc = String(MAX_PATH, 0)
  
  ' Get the error description
  FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, LastErrorNum, 0, Return_ErrDesc, MAX_PATH, 0
  Return_ErrNum = LastErrorNum
  Return_ErrDesc = Left(Return_ErrDesc, InStr(Return_ErrDesc, Chr(0)) - 1)
  If Right(Return_ErrDesc, Len(vbCrLf)) = vbCrLf Then
    Return_ErrDesc = Left(Return_ErrDesc, Len(Return_ErrDesc) - Len(vbCrLf))
  End If
  
  ' Display the error message
  If ShowErrorMsg = True Then
    MsgBox "An error occured while calling the " & LastAPICalled & " Windows API function." & Chr(13) & "Below is the error information:" & Chr(13) & Chr(13) & "Error Number = " & CStr(LastErrorNum) & Chr(13) & "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Windows API Error"
  End If
  GetLastWin32Err = True
  
  ' Set the last error to 0 (no error) so next time through it doesn't report the same error twice
  SetLastError 0
  
End Function

' This function takes the error number passed to it via the "LastErrorNum" perameter and gets the error message
' that that number represents.  If no error number is specified, then check to see if one occured.
Public Function GetLastWinsockErr(Optional ByVal LastErrorNum As Long, _
                                  Optional ByVal LastAPICalled As String = "last", _
                                  Optional ByRef Return_ErrNum As Long, _
                                  Optional ByRef Return_ErrDesc As String, _
                                  Optional ByVal ShowErrorMsg As Boolean = True) As Boolean
  
  Dim ErrMsg As String
  
  ' Reset the return variables
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' If the user didn't specified an error, check if one happened
  If LastErrorNum = 0 Then
    LastErrorNum = WSAGetLastError
    If LastErrorNum = 0 Then
      Exit Function
    End If
  End If
  
  Select Case LastErrorNum
    Case SOCKET_ERROR:               ErrMsg = "Socket Error"
    Case WSA_INVALID_HANDLE:         ErrMsg = "Specified event object handle is invalid" '(An application attempts to use an event object, but the specified handle is not valid)"
    Case WSA_INVALID_PARAMETER:      ErrMsg = "One or more parameters are invalid" '(An application used a Windows Sockets function which directly maps to a Win32 function. The Win32 function is indicating a problem with one or more parameters)"
    Case WSA_IO_PENDING:             ErrMsg = "Overlapped operations will complete later" '(The application has initiated an overlapped operation that cannot be completed immediately. A completion indication will be given later when the operation has been completed)"
    Case WSA_NOT_ENOUGH_MEMORY:      ErrMsg = "Insufficient memory available" '(An application used a Windows Sockets function that directly maps to a Win32 function. The Win32 function is indicating a lack of required memory resources)
    Case WSA_OPERATION_ABORTED:      ErrMsg = "Overlapped operation aborted" '(An overlapped operation was canceled due to the closure of the socket, or the execution of the SIO_FLUSH command in WSAIoctl)
    Case WSAEINTR:                   ErrMsg = "Interrupted function call"
    Case WSAEBADF:                   ErrMsg = "Bad file number"
    Case WSEACCES:                   ErrMsg = "Permission denied"
    Case WSAEFAULT:                  ErrMsg = "Bad address"
    Case WSAEINVAL:                  ErrMsg = "Invalid argument"
    Case WSAEMFILE:                  ErrMsg = "Too many open files"
    Case WSAEWOULDBLOCK:             ErrMsg = "A non-blocking socket operation could not be completed immediately"
    Case WSAEINPROGRESS:             ErrMsg = "A blocking operation is currently executing"
    Case WSAEALREADY:                ErrMsg = "An operation was attempted on a non-blocking socket that already had an operation in progress"
    Case WSAENOTSOCK:                ErrMsg = "An operation was attempted on something that is not a socket"
    Case WSAEDESTADDRREQ:            ErrMsg = "A required address was omitted from an operation on a socket"
    Case WSAEMSGSIZE:                ErrMsg = "A message sent on a datagram socket was larger than the internal message buffer or some other network limit, or the buffer used to receive a datagram into was smaller than the datagram itself"
    Case WSAEPROTOTYPE:              ErrMsg = "A protocol was specified in the socket function call that does not support the semantics of the socket type requested"
    Case WSAENOPROTOOPT:             ErrMsg = "An unknown, invalid, or unsupported option or level was specified in a getsockopt or setsockopt call"
    Case WSAEPROTONOSUPPORT:         ErrMsg = "The requested protocol has not been configured into the system, or no implementation for it exists"
    Case WSAESOCKTNOSUPPORT:         ErrMsg = "The support for the specified socket type does not exist in this address family"
    Case WSAEOPNOTSUPP:              ErrMsg = "The attempted operation is not supported for the type of object referenced"
    Case WSAEPFNOSUPPORT:            ErrMsg = "The protocol family has not been configured into the system or no implementation for it exists"
    Case WSAEAFNOSUPPORT:            ErrMsg = "An address incompatible with the requested protocol was used"
    Case WSAEADDRINUSE:              ErrMsg = "Address already in use - Only one usage of each socket address (protocol/network address/port) is normally permitted"
    Case WSAEADDRNOTAVAIL:           ErrMsg = "The requested address is not valid in its context"
    Case WSAENETDOWN:                ErrMsg = "Network is down - This error may be reported at any time if the Windows Sockets implementation detects an underlying failure"
    Case WSAENETUNREACH:             ErrMsg = "Network is unreachable - A socket operation encountered a dead network"
    Case WSAENETRESET:               ErrMsg = "The connection has been broken due to keep-alive activity detecting a failure while the operation was in progress"
    Case WSAECONNABORTED:            ErrMsg = "An established connection was aborted by the software in your host machine"
    Case WSAECONNRESET:              ErrMsg = "An existing connection was forcibly closed by the remote host"
    Case WSAENOBUFS:                 ErrMsg = "An operation on a socket could not be performed because the system lacked sufficient buffer space or because a queue was full"
    Case WSAEISCONN:                 ErrMsg = "A connect request was made on an already connected socket"
    Case WSAENOTCONN:                ErrMsg = "Socket is not connected - A request to send or receive data was disallowed because the socket is not connected and (when sending on a datagram socket using a sendto call) no address was supplied"
    Case WSAESHUTDOWN:               ErrMsg = "A request to send or receive data was disallowed because the socket had already been shut down in that direction with a previous shutdown call"
    Case WSAETOOMANYREFS:            ErrMsg = "Too many references to some kernel object"
    Case WSAETIMEDOUT:               ErrMsg = "A connection attempt failed because the connected party did not properly respond after a period of time, or established connection failed because connected host has failed to respond"
    Case WSAECONNREFUSED:            ErrMsg = "No connection could be made because the target machine actively refused it"
    Case WSAELOOP:                   ErrMsg = "Too many levels of symbolic links - Cannot translate name"
    Case WSAENAMETOOLONG:            ErrMsg = "Name component or name was too long"
    Case WSAEHOSTDOWN:               ErrMsg = "A socket operation failed because the destination host was down"
    Case WSAEHOSTUNREACH:            ErrMsg = "A socket operation was attempted to an unreachable host"
    Case WSAENOTEMPTY:               ErrMsg = "Cannot remove a directory that is not empty"
    Case WSAEPROCLIM:                ErrMsg = "A Windows Sockets implementation may have a limit on the number of applications that may use it simultaneously"
    Case WSAEUSERS:                  ErrMsg = "Ran out of quota"
    Case WSAEDQUOT:                  ErrMsg = "Ran out of disk quota"
    Case WSAESTALE:                  ErrMsg = "File handle reference is no longer available"
    Case WSAEREMOTE:                 ErrMsg = "Item is not available locally"
    Case WSASYSNOTREADY:             ErrMsg = "WSAStartup cannot function at this time because the underlying system it uses to provide network services is currently unavailable"
    Case WSAVERNOTSUPPORTED:         ErrMsg = "The Windows Sockets version requested is not supported"
    Case WSANOTINITIALISED:          ErrMsg = "Either the application has not called WSAStartup, or WSAStartup failed"
    Case WSAEDISCON:                 ErrMsg = "Disconnect"
    Case WSAENOMORE:                 ErrMsg = "No more results can be returned by WSALookupServiceNext"
    Case WSAECANCELLED:              ErrMsg = "A call to WSALookupServiceEnd was made while this call was still processing - The call has been canceled"
    Case WSAEINVALIDPROCTABLE:       ErrMsg = "The procedure call table is invalid"
    Case WSAEINVALIDPROVIDER:        ErrMsg = "The requested service provider is invalid"
    Case WSAEPROVIDERFAILEDINIT:     ErrMsg = "The requested service provider could not be loaded or initialized"
    Case WSASYSCALLFAILURE:          ErrMsg = "A system call that should never fail has failed"
    Case WSASERVICE_NOT_FOUND:       ErrMsg = "No such service is known - The service cannot be found in the specified name space"
    Case WSATYPE_NOT_FOUND:          ErrMsg = "The specified class was not found"
    Case WSA_E_NO_MORE:              ErrMsg = "No more results can be returned by WSALookupServiceNext"
    Case WSA_E_CANCELLED:            ErrMsg = "A call to WSALookupServiceEnd was made while this call was still processing - The call has been canceled"
    Case WSAEREFUSED:                ErrMsg = "A database query failed because it was actively refused"
    Case WSAHOST_NOT_FOUND:          ErrMsg = "Host not found - This message indicates that the key (name, address, and so on) was not found"
    Case WSATRY_AGAIN:               ErrMsg = "Nonauthoritative host not found - This error may suggest that the name service itself is not functioning"
    Case WSANO_RECOVERY:             ErrMsg = "Nonrecoverable error - This error may suggest that the name service itself is not functioning"
    Case WSANO_DATA:                 ErrMsg = "Valid name, no data record of requested type - This error indicates that the key (name, address, and so on) was not found"
    Case WSA_QOS_RECEIVERS:          ErrMsg = "At least one Reserve has arrived"
    Case WSA_QOS_SENDERS:            ErrMsg = "At least one Path has arrived"
    Case WSA_QOS_NO_SENDERS:         ErrMsg = "There are no senders"
    Case WSA_QOS_NO_RECEIVERS:       ErrMsg = "There are no receivers"
    Case WSA_QOS_REQUEST_CONFIRMED:  ErrMsg = "Reserve has been confirmed"
    Case WSA_QOS_ADMISSION_FAILURE:  ErrMsg = "Error due to lack of resources"
    Case WSA_QOS_POLICY_FAILURE:     ErrMsg = "Rejected for administrative reasons - bad credentials"
    Case WSA_QOS_BAD_STYLE:          ErrMsg = "Unknown or conflicting style"
    Case WSA_QOS_BAD_OBJECT:         ErrMsg = "Problem with some part of the filterspec or providerspecific buffer in general"
    Case WSA_QOS_TRAFFIC_CTRL_ERROR: ErrMsg = "Problem with some part of the flowspec"
    Case WSA_QOS_GENERIC_ERROR:      ErrMsg = "General QOS error"
    Case WSA_QOS_ESERVICETYPE:       ErrMsg = "An invalid or unrecognized service type was found in the flowspec"
    Case WSA_QOS_EFLOWSPEC:          ErrMsg = "An invalid or inconsistent flowspec was found in the QOS structure"
    Case WSA_QOS_EPROVSPECBUF:       ErrMsg = "Invalid QOS provider-specific buffer"
    Case WSA_QOS_EFILTERSTYLE:       ErrMsg = "An invalid QOS filter style was used"
    Case WSA_QOS_EFILTERTYPE:        ErrMsg = "An invalid QOS filter type was used"
    Case WSA_QOS_EFILTERCOUNT:       ErrMsg = "An incorrect number of QOS FILTERSPECs were specified in the FLOWDESCRIPTOR"
    Case WSA_QOS_EOBJLENGTH:         ErrMsg = "An object with an invalid ObjectLength field was specified in the QOS provider-specific buffer"
    Case WSA_QOS_EFLOWCOUNT:         ErrMsg = "An incorrect number of flow descriptors was specified in the QOS structure"
    Case WSA_QOS_EUNKOWNPSOBJ:       ErrMsg = "An unrecognized object was found in the QOS provider-specific buffer"
    Case WSA_QOS_EPOLICYOBJ:         ErrMsg = "An invalid policy object was found in the QOS provider-specific buffer"
    Case WSA_QOS_EFLOWDESC:          ErrMsg = "An invalid QOS flow descriptor was found in the flow descriptor list"
    Case WSA_QOS_EPSFLOWSPEC:        ErrMsg = "An invalid or inconsistent flowspec was found in the QOS provider-specific buffer"
    Case WSA_QOS_EPSFILTERSPEC:      ErrMsg = "An invalid FILTERSPEC was found in the QOS provider-specific buffer"
    Case WSA_QOS_ESDMODEOBJ:         ErrMsg = "An invalid shape discard mode object was found in the QOS provider-specific buffer"
    Case WSA_QOS_ESHAPERATEOBJ:      ErrMsg = "An invalid shaping rate object was found in the QOS provider-specific buffer"
    Case WSA_QOS_RESERVED_PETYPE:    ErrMsg = "A reserved policy element was found in the QOS provider-specific buffer"
    Case Else:                       ErrMsg = "Unknown Error"
  End Select
  
  ' Return the error information
  Return_ErrNum = LastErrorNum
  Return_ErrDesc = ErrMsg
  
  ' If the user specified to, show an error message
  If ShowErrorMsg = True Then
    MsgBox "The following Winsock error occured during a call to the " & LastAPICalled & " API:" & Chr(13) & Chr(13) & "Error Number = " & CStr(LastErrorNum) & Chr(13) & "Error Description = " & ErrMsg, vbOKOnly + vbExclamation, "  Winsock Error"
  End If
  
  ' Return that an error occured
  GetLastWinsockErr = True
  
End Function
