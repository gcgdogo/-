Option Compare Database
Option Explicit

Public Sub WriteNewLine(In_Type As String, In_Txt As String)

    Dim NewLine_String As String, FontColor_String As String
    Dim Warning_Line As String

    NewLine_String = Format(Now(), "hh:mm:ss") & "." & Format((Timer * 1000) Mod 1000, "000") & " | "
    
    NewLine_String = NewLine_String & Left(In_Type & "          ", 10) & " | "
    NewLine_String = NewLine_String & In_Txt

    FontColor_String = "#000000"
    If In_Type = "Warning" Then FontColor_String = "#FF0000"
    If In_Type = "Command" Then FontColor_String = "#00FF00"

    Warning_Line = "#################################################################################################################"
    If In_Type = "Warning" Then NewLine_String = Warning_Line & vbCrLf & NewLine_String & vbCrLf & Warning_Line

    If FontColor_String <> "#000000" Then NewLine_String = "<font color=" & FontColor_String & ">" & NewLine_String & "</font color>"

    NewLine_String = "<div>" & NewLine_String & "</div>"

    'Text有长度限制 ， 用Value可以 ： 而且不用 Focus
    Me.Textbox_Everything.Value = NewLine_String & Me.Textbox_Everything.Value
    
    If In_Type = "Running" Then Exit Sub

    Me.Textbox_Message.Value = NewLine_String & Me.Textbox_Everything.Value

    
End Sub