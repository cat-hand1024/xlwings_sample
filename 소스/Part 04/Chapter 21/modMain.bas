Attribute VB_Name = "modMain"
Option Explicit

Sub ��������()

    Dim ���� As String
    Dim �޽��� As String

    Select Case Application.Version
    
        Case "11.0": ���� = "2003"
        Case "12.0": ���� = "2007"
        Case "14.0": ���� = "2010"
        Case "15.0": ���� = "2013"
        Case "16.0": ���� = "2016, 2019, 365 �� �ϳ�"
        
    End Select
    
    If Len(����) > 0 Then
    
        �޽��� = "���� ��� ���� ������ ���� " & ���� & " �Դϴ�."
        
    Else
    
        �޽��� = "���� ������ �� �� �����ϴ�."
    
    End If

    MsgBox Prompt:=�޽���, Title:="���� üũ"

End Sub
