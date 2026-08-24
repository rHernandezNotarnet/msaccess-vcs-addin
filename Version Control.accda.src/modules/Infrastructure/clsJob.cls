VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'@Folder("Infrastructure")

Public KeyHash As String
Public Action As String
Public Start As Double
Public Timeout As Double

' File the worker writes its return value to, for a job that reports through the
' file channel instead of a COM callback. (See clsWorker.PollJobResults.)
Public ResultFile As String
