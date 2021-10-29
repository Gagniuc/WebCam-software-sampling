VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "timTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : timTimer
' Author    : James Arnold ("Rezzie")
' Purpose   : Timer from a class module, without the need for a form
'---------------------------------------------------------------------------------------

'We'll use the GetTickCount API, and use a currency to store tick count
Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim timStart As Currency, timElapsed As Double, Running As Boolean

Public Sub StartTimer()
  timElapsed = 0 'Reset timer
  Running = True 'Now running
  timStart = GetTickCount 'Assign start time
End Sub

Public Property Get Elapsed() As Double
'If timer is running then read API, otherwise read variable
  If Running Then Elapsed = GetElapsedTime Else Elapsed = timElapsed
End Property

Public Sub EndTimer()
  timElapsed = GetElapsedTime 'Assign end time
  Running = False 'No longer running
End Sub

Private Function GetElapsedTime() As Double
  GetElapsedTime = GetTickCount - timStart 'Read API
End Function


