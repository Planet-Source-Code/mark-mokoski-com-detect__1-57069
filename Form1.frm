VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COM Port Detect"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Detect Progress"
      Height          =   1455
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   6015
      Begin VB.TextBox InfoText 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Text            =   "Form1.frx":1272
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Active COM Ports"
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   6015
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   7
         Left            =   4680
         MouseIcon       =   "Form1.frx":127D
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1587
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   6
         Left            =   3120
         MouseIcon       =   "Form1.frx":19C9
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":1CD3
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   5
         Left            =   1680
         MouseIcon       =   "Form1.frx":2115
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":241F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   4
         Left            =   240
         MouseIcon       =   "Form1.frx":2861
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2B6B
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   3
         Left            =   4680
         MouseIcon       =   "Form1.frx":2FAD
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":32B7
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   2
         Left            =   3120
         MouseIcon       =   "Form1.frx":36F9
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":3A03
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   1
         Left            =   1680
         MouseIcon       =   "Form1.frx":3E45
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":414F
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton COMicon 
         Caption         =   "COM x"
         Height          =   855
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form1.frx":4591
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":489B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   20
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   19
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Label3"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   7
         Left            =   4680
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   12
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   2
         Left            =   3120
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortSettings 
         Alignment       =   2  'Center
         Caption         =   "Settings"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label PortType 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Serial Port"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      MouseIcon       =   "Form1.frx":4CDD
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":4FE7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ComboBox txtPortNumber 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2250
      MouseIcon       =   "Form1.frx":52F1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Text            =   "8"
      Top             =   5400
      Width           =   645
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   6360
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      MouseIcon       =   "Form1.frx":55FB
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":5905
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COM ports to test: 1 thru "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5445
      Width           =   2130
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
Dim buffer As String
Dim myport As Integer
Dim TimeElapsed As Integer
Dim PortInfo(7) As String

Dim Command1Tip As New clsTooltips
Dim Command2Tip As New clsTooltips
Dim Frame1Tip As New clsTooltips
Dim InfoTextTip As New clsTooltips
Dim COMiconTip(7) As New clsTooltips

Private Sub CheckPort(X As Integer)
    
    Timer1.Enabled = False
    
        If X <> 1 Then PrintText "---------------------"
   
    PrintText "Checking COM" & Trim(Str(X)) & "..."
   
    'Check for port status

        If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
   
    'Start error handling
    On Error GoTo ErrorHandler
    MSComm1.CommPort = X
    MSComm1.Settings = "19200,N,8,1"
    MSComm1.InputLen = 0
    MSComm1.PortOpen = True
    On Error GoTo 0
   
    MSComm1.Output = "ATI1" & Chr$(13)
        
    'Wait for a response for 2 seconds. If nothing returns then exit

        If WaitForResponse(1) = False Then GoTo NothingReturned
        
    MSComm1.Output = "ATI4" & Chr$(13)
        
    'Wait for response for 2 seconds. If nothing returns then exit

        If WaitForResponse(1) = False Then GoTo NothingReturned
    
    'If something returned...
    PrintText ParseBuffer(buffer)
    PrintText "COM" & Trim(Str(X)) & " is a modem."
    MSComm1.PortOpen = False
    COMicon(X - 1).Picture = LoadResPicture(102, vbResIcon)
    COMicon(X - 1).Caption = "COM " + Str(X)
    COMicon(X - 1).Visible = True
    PortType(X - 1).Caption = "Modem"
    PortInfo(X - 1) = ParseBuffer(buffer)
    PortType(X - 1).Visible = True
    PortSettings(X - 1).Caption = UCase(MSComm1.Settings)
    PortSettings(X - 1).Visible = True
    COMiconTip(X - 1).Title = "Modem"
    COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & PortInfo(X - 1) & vbCrLf & "Click for Details"

    Exit Sub
   
NothingReturned:
    PrintText "Is an installed COM Port"
    MSComm1.PortOpen = False
    COMicon(X - 1).Picture = LoadResPicture(101, vbResIcon)
    COMicon(X - 1).Caption = "COM" + Str(X)
    COMicon(X - 1).Visible = True
    PortType(X - 1).Caption = "Serial Port"
    PortInfo(X - 1) = ""
    PortType(X - 1).Visible = True
    PortSettings(X - 1).Caption = UCase(MSComm1.Settings)
    PortSettings(X - 1).Visible = True
    COMiconTip(X - 1).Title = "Serial Port"
    COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"
    
    Exit Sub


ErrorHandler:

    PrintText ErrorString(Err.Number, X)

End Sub

Private Sub COMicon_Click(Index As Integer)
Dim Handshake As String

'Get some COM Port info and display in listbox
InfoText.Text = ""
MSComm1.CommPort = (Index + 1)
PrintText "COM Port: " + Str(Index + 1)
PrintText "Port Type: " + PortType(Index).Caption
PrintText "Modem ID: " + PortInfo(Index)
PrintText "Settings: " + PortSettings(Index)
'Skip if port is in use bt another App
Select Case PortSettings(Index)
    Case "Port In Use"
    
    Case Else
    
        PrintText "DTR: " + Str(MSComm1.DTREnable)
        PrintText "RTS: " + Str(MSComm1.RTSEnable)

        Select Case MSComm1.Handshaking
            Case 0
                Handshake = "NONE"
            Case 1
                Handshake = "Xon/Xoff"
            Case 2
                Handshake = "RTS"
            Case 3
                Handshake = "RTS & Xon/Xoff"
        End Select
        
        PrintText "Handshaking: " + Handshake
        
End Select
InfoText.SelStart = 1

End Sub

Private Sub Command1_Click()

Frame1Tip.Active = True

Command1.Enabled = False
Command1.BackColor = vbButtonFace
Command2.Enabled = False
Command2.BackColor = vbButtonFace

Dim c As Integer

If (Len(txtPortNumber.Text) = 0 Or Val(txtPortNumber.Text) = 0) Then
    MsgBox "Please enter a valid integer.", vbInformation
    Exit Sub
End If

Dim i As Integer

For c = 0 To 7
    COMicon(c).Visible = False
    PortType(c).Visible = False
    PortSettings(c).Visible = False
Next c

InfoText.Text = ""

For i = 1 To Int(Val(txtPortNumber.Text))
    Call CheckPort(i)
Next

Command1.Enabled = True
Command1.BackColor = &HC0C0C0
Command2.Enabled = True
Command2.BackColor = &HC0C0C0

PrintText "---------------------"
PrintText "COM Port search complete"
PrintText "Click START to rescan COM Ports"

Frame1Tip.Active = False

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Y As Integer
Dim z As String


Command1Tip.CreateBalloon Command1, "Click to Start Testing", "Start COM Test", 1
Command2Tip.CreateBalloon Command2, "Click to end Program", "End Application", 3
InfoTextTip.CreateBalloon InfoText, "Test Status Messages and COM Port info", "Status Messages", 1
Frame1Tip.CreateBalloon Frame1, "COM Port Status Shown Here", "COM Port Test Results", 1

Timer1.Enabled = False

'Load Combo Box with Port number list
For Y = 0 To 7
    z = Str(Y + 1)
    txtPortNumber.AddItem z, Y
    COMiconTip(Y).CreateBalloon COMicon(Y), "COM" & z & " Installed" & " ", 1
Next Y

InfoText.Text = ""
PrintText "Click START to begin COM Port detect"

End Sub

Private Sub Timer1_Timer()

    TimeElapsed = TimeElapsed + 1

End Sub

Private Function WaitForResponse(X As Integer) As Boolean
    
    buffer = ""
    WaitForResponse = False
    TimeElapsed = 0
    Timer1.Enabled = True
   
        Do
      
            DoEvents
      
            buffer = buffer & MSComm1.Input
      
                If Len(buffer) <> 0 Then

                        If InStr(1, buffer, "OK") <> 0 Then
            
                            WaitForResponse = True
                            Timer1.Enabled = False
                            TimeElapsed = 0
                            
                            Exit Function
                        End If

                End If
   
                If TimeElapsed > X Then
                    Timer1.Enabled = False
                    Exit Function
                End If
   
        Loop
   
End Function
Private Sub PrintText(X As String)
InfoText.Text = InfoText.Text + X + vbCrLf
InfoText.SelStart = Len(InfoText.Text) + 1
End Sub
Private Function ParseBuffer(X As String) As String
Dim i As Integer
Dim Splitter As String
Dim Pos1, Pos2 As Integer

Splitter = Chr(13) & Chr(10)
Pos1 = InStr(1, X, Splitter)
Pos2 = InStr(Pos1 + 2, X, Splitter)

ParseBuffer = Mid(X, Pos1 + 2, Pos2 - Pos1 - 2)

End Function

Private Function ErrorString(ERROR As Long, X As Integer) As String

    Dim tmp            As String

        Select Case ERROR
            Case 8021
                tmp = "Internal error retrieving device control block for the port"
                COMicon(X - 1).Picture = LoadResPicture(104, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 394
                tmp = "Property is write-only"
            Case 380
                tmp = "Invalid property value"
            Case 8012
                'tmp = "The device is not open"
                tmp = "Port In Use"
                COMicon(X - 1).Picture = LoadResPicture(103, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 8005
                tmp = "Port In Use"
                COMicon(X - 1).Picture = LoadResPicture(103, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 8002
                tmp = "Invalid port number"
            Case 8018
                tmp = "Operation valid only when the port is open"
                COMicon(X - 1).Picture = LoadResPicture(104, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 8000
                tmp = "Operation not valid while the port is opened"
                COMicon(X - 1).Picture = LoadResPicture(104, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 8020 & 8015
                'tmp = "Error reading comm device"
                tmp = "Port ERROR"
                COMicon(X - 1).Picture = LoadResPicture(104, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

            Case 383
                tmp = "Property is read-only"
            Case Else
                tmp = "Other error..."
                COMicon(X - 1).Picture = LoadResPicture(104, vbResIcon)
                COMicon(X - 1).Caption = "COM " + Str(X)
                COMicon(X - 1).Visible = True
                PortType(X - 1).Caption = "Serial Port"
                PortInfo(X - 1) = ""
                PortType(X - 1).Visible = True
                PortSettings(X - 1).Caption = tmp
                PortSettings(X - 1).Visible = True
                COMiconTip(X - 1).Title = "Serial Port"
                COMiconTip(X - 1).TipText = "COM" & Str(X) & " Installed" & vbCrLf & "Click for Details"

        End Select

    ErrorString = tmp

End Function
