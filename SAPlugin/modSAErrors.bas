Attribute VB_Name = "modSAErrors"
 ' modSAErrors
Option Explicit

Public Enum SAErrors
    EX_USAGE = 64
    EX_DATAERR
    EX_NOINPUT
    EX_NOUSER
    EX_NOHOST
    EX_UNAVAILABLE
    EX_SOFTWARE
    EX_OSERR
    EX_OSFILE
    EX_CANTCREAT
    EX_IOERR
    EX_TEMPFAIL
    EX_PROTOCOL
    EX_NOPERM
    EX_CONFIG
    
    
End Enum


Public Function ReturnCodeToString(rc As Integer)
    
    Dim s As String
    s = "Unknown error " & CStr(rc)
    Select Case rc
    
    Case EX_USAGE: s = "64 command line usage error"
    Case EX_DATAERR: s = "65 data format error"
    Case EX_NOINPUT: s = "66 cannot open input"
    Case EX_NOUSER: s = "67 addressee unknown"
    Case EX_NOHOST: s = "68 host name unknown"
    Case EX_UNAVAILABLE: s = "69 service unavailable"
    Case EX_SOFTWARE: s = "70  internal software error"
    Case EX_OSERR: s = "71  system error (e.g., can't fork)"
    Case EX_OSFILE: s = "72  critical OS file missing"
    Case EX_CANTCREAT: s = "73  can't create (user) output file"
    Case EX_IOERR: s = "74  input/output error"
    Case EX_TEMPFAIL: s = "75  temp failure; user is invited to retry"
    Case EX_PROTOCOL: s = "76  remote error in protocol"
    Case EX_NOPERM: s = "77  permission denied"
    Case EX_CONFIG:      s = "78  configuration error"
    End Select
    
    ReturnCodeToString = s

End Function

