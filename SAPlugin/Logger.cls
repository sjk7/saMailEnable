VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim m_Path As String
Dim FSO As New FileSystemObject
Dim m_ts As TextStream

Const LOG_ROTATE_SIZE_BYTES As Long = 1024& * 1024& * 200&

Private Sub Create()


    Debug.Assert LenB(App.ProductName)
    m_Path = App.Path & "\" & App.ProductName & ".log"
   Set m_ts = FSO.OpenTextFile(m_Path, ForAppending, Not FSO.FileExists(m_Path))
   Log vbNewLine, True
   Log "----------------------------------------------------", True
    Log App.ProductName & " Started"
   
    
End Sub

Public Sub Log(s As String, Optional exclusiveText As Boolean)



        Dim sz As Long: sz = FileLen(m_Path)
        If (sz > LOG_ROTATE_SIZE_BYTES) Then
            m_ts.Close
            Kill m_Path
            Create
            Log "Old log file too big, so I created a new one."
        End If
        


    If Not exclusiveText Then
        Dim output As String
        output = Now & vbTab & s
        m_ts.WriteLine (output)
        Debug.Print output
    Else
        m_ts.WriteLine s
        Debug.Print s
    End If
    

    
    
End Sub


Private Sub Class_Initialize()
    Create
End Sub

Private Sub Class_Terminate()
    On Local Error Resume Next
       Log vbNewLine, True
   Log "----------------------------------------------------", True
    Log App.ProductName & " Closed."
    m_ts.Close
    
End Sub
