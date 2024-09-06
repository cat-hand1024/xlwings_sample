Attribute VB_Name = "modMain"
Option Explicit

Sub 엑셀버전()

    Dim 버전 As String
    Dim 메시지 As String

    Select Case Application.Version
    
        Case "11.0": 버전 = "2003"
        Case "12.0": 버전 = "2007"
        Case "14.0": 버전 = "2010"
        Case "15.0": 버전 = "2013"
        Case "16.0": 버전 = "2016, 2019, 365 중 하나"
        
    End Select
    
    If Len(버전) > 0 Then
    
        메시지 = "현재 사용 중인 버전은 엑셀 " & 버전 & " 입니다."
        
    Else
    
        메시지 = "엑셀 버전을 알 수 없습니다."
    
    End If

    MsgBox Prompt:=메시지, Title:="버전 체크"

End Sub
