Attribute VB_Name = "display_window_features"
Sub Main()

    display_window_features (False) ' Turn features off
    display_window_features (True) ' Turns features on

End Sub

Sub display_window_features(display As Boolean)
    
    ' Hides/shows formula bar, scroll bars, status bar & ribbon
    With Application
        .ExecuteExcel4Macro "show.toolbar(""Ribbon""," + CStr(display) + ")"
        .DisplayFormulaBar = display
        .DisplayScrollBars = display
        .DisplayStatusBar = display
    End With

End Sub
