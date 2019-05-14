Option Compare Database
Option Explicit

Dim String_Everything As String

Dim String_Main_A As String, String_Main_B As String, String_Running As String
'需要提前储存修饰符长度，要不然光标不好定位
Dim Modifiers_A As Long, Modifiers_B As Long

Dim LastPrintTimer As Double

Public Sub WriteNewLine(In_Type As String, In_Txt As String)

    Dim NewLine_String As String, FontColor_String As String
    Dim Warning_Line As String
    Dim Len_Modifers As Long
    
    Len_Modifers = 0
    
    NewLine_String = Format(Now(), "hh:mm:ss") & "." & Format((Timer * 1000) Mod 1000, "000") & " | "
    
    NewLine_String = NewLine_String & Left(In_Type & "________", 7) & " | "
    NewLine_String = NewLine_String & In_Txt

    FontColor_String = "#000000"
    If In_Type = "Warning" Then FontColor_String = "#FF0000"
    If In_Type = "Command" Then FontColor_String = "#00BB00"

    Warning_Line = "########################"
    If In_Type = "Warning" Then
        NewLine_String = Warning_Line & "<br>" & NewLine_String & "<br>" & Warning_Line
        '<br>格式化后还剩一个换行符
        Len_Modifers = Len_Modifers + 4 + 4 - 2
    End If
    
    If FontColor_String <> "#000000" Then
        NewLine_String = "<font color=" & FontColor_String & ">" & NewLine_String & "</font color>"
        Len_Modifers = Len_Modifers + Len("<font color=" & FontColor_String & ">" & "</font color>")
    End If
    'NewLine_String = NewLine_String & "<br>"

    String_Everything = String_Everything & NewLine_String & "<br>"
    
    If In_Type <> "Running" Then
        String_Main_B = String_Main_B & NewLine_String & "<br>"
        Len_Modifers = Len_Modifers + 4 - 1
        '如果不是Running就不加Modifiers
        Modifiers_B = Modifiers_B + Len_Modifers
        String_Running = ""
        If Len(String_Main_B) > 5000 Then

            String_Main_A = String_Main_B
            Modifiers_A = Modifiers_B

            String_Main_B = ""
            Modifiers_B = 0

        End If
    Else
        String_Running = NewLine_String
    End If
    
    Me.Message_Switch.Value = False
    Call Message_Switch_AfterUpdate
End Sub

Private Sub Message_Switch_AfterUpdate()
    Static Last_Mode As Boolean
    
    '记录上一次执行模式，模式变动时重新设置焦点
    If Me.Message_Switch.Value <> Last_Mode Then
        Me.Textbox_Message.SetFocus
        Last_Mode = Me.Message_Switch.Value
    End If
    
    On Error GoTo Error_To_Select
    
    Do While Timer() - LastPrintTimer < 0.05
        DoEvents
        DoEvents
        DoEvents
    Loop
   
    'Text有长度限制 ， 用Value可以 ： 而且不用 Focus
    
    If Me.Message_Switch.Value = True Then
        Me.Textbox_Message.Value = String_Everything
        Me.Textbox_Message.SetFocus
        Me.Textbox_Message.SelStart = 0
    Else
        Me.Textbox_Message.Value = String_Main_A & String_Main_B & String_Running
        
        If Len(Me.Textbox_Message.Value) - Modifiers_A - Modifiers_B >= 0 Then
            Me.Textbox_Message.SelStart = Len(Me.Textbox_Message.Value) - Modifiers_A - Modifiers_B
        End If
        
    End If

    LastPrintTimer = Timer()
    
    DoEvents
    
    Exit Sub
    
    
Error_To_Select:
    
    '修正后重试
    Me.Textbox_Message.SetFocus
    Modifiers_B = Modifiers_B + 1
    '增加标记
    String_Running = String_Running & "_Retry"
    
    Call Message_Switch_AfterUpdate
    
End Sub

