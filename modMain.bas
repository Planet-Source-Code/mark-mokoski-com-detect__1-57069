Attribute VB_Name = "modMain"
Option Explicit
    'To use Balloon tips,Int Common Controls Lib
    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Sub Main()

    'Int Common Controls Lib for Balloon Tip use
    InitCommonControls

    ' * Test to see if App is already running
    ' * If App is running, terminate copy

        If App.PrevInstance Then
            MsgBox "COM Detect application is already running" & vbCrLf & vbCrLf & _
            "Only one instance (copy) of program this can be running" & vbCrLf & _
            "for proper operation", vbInformation, "Application ERROR"
            End
        Else
            '  MsgBox "This is the first instance of your application"
            'Make main form visible
            Load frmMain
            frmMain.Visible = True
            
        End If

End Sub


