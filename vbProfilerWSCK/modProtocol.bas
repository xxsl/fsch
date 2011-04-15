Attribute VB_Name = "modProtocol"
Option Explicit

Public Const C_BEGIN                        As Byte = 1

Public Const C_INIT                         As Byte = 2

Public Const C_RUNNING                      As Byte = 4

Public Const C_SAMPLE                       As Byte = 0

Public Const C_NEW_OBJECT_SAMPLE            As Byte = 1

Public Const C_DELETE_OBJECT_SAMPLE         As Byte = 2

Public Const C_MEMBERS                      As Byte = 3

Public Const C_SWF                          As Byte = 4

Public Const C_GC                           As Byte = 5

Public Const C_STRING                       As Byte = 6

Public Const C_COUNTS                       As Byte = 7

Public Const C_START_SAMPLING               As Byte = 0

Public Const C_STOP_SAMPLING                As Byte = 1

Public Const C_GET_MEMBERS                  As Byte = 2

Public Const C_PAUSE                        As Byte = 3

Public Const C_RESUME                       As Byte = 4

Public Const C_EXIT                         As Byte = 5

Public Const C_FORCE_GC                     As Byte = 6

Public Const C_MEMORY_PROFILING             As Byte = 7

Public Const C_MEMORY_PROFILING_STACKTRACES As Byte = 8

Public Const C_PERFORMANCE_PROFILING        As Byte = 9

Public Const C_GET_COUNTS                   As Byte = 10
