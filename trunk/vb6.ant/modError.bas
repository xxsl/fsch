Attribute VB_Name = "modConst"
Public Enum FCSHErrors
    FCSH_NOT_STATRED = 1&
    FCSH_BUSY = 2&
    FCSH_START_FAILURE = 3&
    FCSH_ALREADY_STATRED = 4&
End Enum

Public Const FCSH_STOPPED As String = "fcsh_stop"
Public Const FCSH_STARTED As String = "fcsh_start"
Public Const FCSH_DATA As String = "fcsh_data"
Public Const FCSH_STATE  As String = "fcsh_getstate"
Public Const FCSH_BUILD_SUCCESSFULL As String = "fcsh_build_success"
Public Const FCSH_BUILD_WARNING As String = "fcsh_build_warn"
Public Const FCSH_BUILD_ERROR As String = "fcsh_build_error"

Public Enum Build
    BUILD_ERROR = 10&
    BUILD_SUCCESSFULL = 11&
    BUILD_WARNING = 12&
End Enum


