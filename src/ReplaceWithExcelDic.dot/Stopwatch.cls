VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private StartTime As Date
Private EndTime As Date
Private AccumulatedTime As Date

'@return: 現在のインスタンスで計測された経過時間の合計を取得します。
Property Get Elapsed() As Date
    Elapsed = AccumulatedTime
End Property


Public Sub Start()
    StartTime = Now
End Sub

Public Sub StartNew()
    StartTime = Now
    
    AccumulatedTime = 0
End Sub

Public Sub Stop_()
    EndTime = Now
    
    AccumulatedTime = AccumulatedTime + EndTime - StartTime
End Sub

Public Sub Reset()
    AccumulatedTime = 0
End Sub

