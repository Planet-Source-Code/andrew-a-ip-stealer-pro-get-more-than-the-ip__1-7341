VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IP Stealer Pro - By Plasma"
   ClientHeight    =   5025
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   6915
      Begin MSComDlg.CommonDialog dlgCommon 
         Left            =   4260
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Script"
         Filter          =   "*.ips"
      End
      Begin MSWinsockLib.Winsock sckServer 
         Index           =   0
         Left            =   5100
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock sckListen 
         Left            =   4680
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "http://127.0.0.1:1234"
         Top             =   240
         Width           =   1995
      End
      Begin VB.CommandButton cmdOffline 
         Caption         =   "Offline"
         Default         =   -1  'True
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Hide Options"
         Height          =   315
         Left            =   5580
         TabIndex        =   11
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdOnline 
         Caption         =   "Online"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Collected IP's"
      Height          =   2175
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   6915
      Begin ComctlLib.ListView lvIPs 
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   3201
         View            =   3
         Arrange         =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "IP"
            Object.Width           =   2379
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Time"
            Object.Width           =   2302
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Browser"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "System OS"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Config"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   6915
      Begin VB.CommandButton cmdScript 
         Caption         =   "Script"
         Height          =   315
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   660
         Width           =   795
      End
      Begin VB.CheckBox chkDupe 
         Caption         =   "Do not add IP's if they are already in the list."
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   3435
      End
      Begin VB.TextBox txtErrorText 
         Height          =   315
         Left            =   900
         TabIndex        =   6
         Text            =   "Sorry, the requested file could not be found."
         Top             =   660
         Width           =   3915
      End
      Begin VB.CheckBox chkBeep 
         Caption         =   "Beep when an IP address has been added."
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   3375
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "80"
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label lblHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label lblHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   6180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   13
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Error text:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   795
      End
      Begin VB.Label lblHelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6180
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   3
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Port to listen on:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuCopyIP 
         Caption         =   "Copy IP"
      End
      Begin VB.Menu sepbar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyIPAddy 
         Caption         =   "Copy IP Address"
      End
      Begin VB.Menu mnuCopyTime 
         Caption         =   "Copy Time Info"
      End
      Begin VB.Menu mnuCopyBrowser 
         Caption         =   "Copy Browser Info"
      End
      Begin VB.Menu mnuCopyOS 
         Caption         =   "Copy Operating System Info"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOffline_Click()
cmdOffline.Enabled = False
Me.txtPort.Enabled = True
For i = 0 To 99
IPData(i).FreeSocket = True
sckServer(i).Close
Next i
sckListen.Close
Me.Caption = "IP Stealer Pro - Offline"
cmdOnline.Enabled = True
cmdOnline.Default = True
End Sub

Private Sub cmdOnline_Click()
On Error GoTo Online_Err
cmdOnline.Enabled = False
For i = 0 To 99
IPData(i).FreeSocket = True
Next i
sckListen.LocalPort = txtPort
sckListen.Listen
If txtPort = 80 Then
txtAddress = "http://" & sckListen.LocalIP
Else
txtAddress = "http://" & sckListen.LocalIP & ":" & txtPort
End If
Me.txtPort.Enabled = False
Me.Caption = "IP Stealer Pro - Online"
cmdOffline.Enabled = True
cmdOffline.Default = True
Exit Sub
Online_Err:
If Err = 13 Then 'Type mismatch
    MsgBox "Error: The listen port must only contain numbers." & vbCrLf & vbCrLf & "Please correct the error and try again.", vbExclamation, "Port Error"
    cmdOptions.Caption = "Hide Options"
    Me.Height = 5850
    cmdOffline_Click
    txtPort.Enabled = True
    txtPort.SetFocus
Else
    MsgBox "Unknown error: Err.Number-" & Err.Number & vbCrLf & "Err.Description-" & Err.Description, vbExclamation, "Unknown error"
End If
End Sub

Private Sub cmdOptions_Click()
If cmdOptions.Caption = "Show Options" Then
cmdOptions.Caption = "Hide Options"
Me.Height = 5850
Else
cmdOptions.Caption = "Show Options"
Me.Height = 3630
End If
End Sub

Private Sub cmdScript_Click()
On Error GoTo Script_err
If Not Left(txtErrorText, 7) = "SCRIPT:" Then
ErrorText = txtErrorText
End If

If cmdScript.BackColor = &H8000000F Then 'Grey colour
dlgCommon.CancelError = True 'This will make it so if the user selects
'cancel button, it will cause an error (VB Runtime error)
dlgCommon.Filter = "IP Stealer scripts (*.ips)|*.ips|Text Documents (*.txt)|*.txt|Webpages (*.htm; *.html)|*.htm; *.html"
dlgCommon.InitDir = App.Path
dlgCommon.ShowOpen
txtErrorText.Locked = True
txtErrorText.ForeColor = vbBlue
ScriptPath = dlgCommon.FileName
txtErrorText = "SCRIPT:" & dlgCommon.FileName
cmdScript.BackColor = vbBlue
Else
txtErrorText = ErrorText
txtErrorText.ForeColor = vbBlack
cmdScript.BackColor = &H8000000F
txtErrorText.Locked = False
End If
'End If
Exit Sub
Script_err:
If Err = 32755 Then 'User pressed 'CANCEL' button for the dialog box
txtErrorText = ErrorText
txtErrorText.ForeColor = vbBlack
cmdScript.BackColor = &H8000000F
txtErrorText.Locked = False
End If
End Sub

Private Sub Form_Load()
Options = GetSetting(App.EXEName, "Config", "Options", True)
FirstTime = GetSetting(App.EXEName, "StartUp", "FirstTime", True)
ErrorText = GetSetting(App.EXEName, "Config", "ErrorText", "Error 404 - File not found")
ScriptPath = GetSetting(App.EXEName, "Config", "ScriptPath")
Me.txtPort = GetSetting(App.EXEName, "Config", "Port", "80")
Me.chkBeep = GetSetting(App.EXEName, "Config", "Beep", 0)
Me.chkDupe = GetSetting(App.EXEName, "Config", "Dupe", 0)
cmdScript.BackColor = GetSetting(App.EXEName, "Config", "UseScript", &H8000000F)
cmdOnline_Click
If cmdScript.BackColor = vbBlue Then
txtErrorText = "SCRIPT:" & ScriptPath
txtErrorText.ForeColor = vbBlue
Else
txtErrorText = ErrorText
End If

If Options = True Then
cmdOptions.Caption = "Show Options"
cmdOptions_Click
Else
cmdOptions.Caption = "Hide Options"
cmdOptions_Click
End If


If FirstTime = True Then 'First time running the program
frmMain.Show
frmAbout.Show vbModal  'vbModal - User must respond to this form
'before doing anything else in the app. Like a MSGBOX, user must press OK to continue...
End If


For i = 1 To 99
Load sckServer(i)
IPData(i).FreeSocket = True
Next i
IPData(0).FreeSocket = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Options As Boolean

If cmdOptions.Caption = "Hide Options" Then
Options = True
Else
Options = False
End If

If cmdScript.BackColor = &H8000000F Then
ErrorText = txtErrorText
Else
ScriptPath = Mid(txtErrorText, 8)
End If
SaveSetting App.EXEName, "Config", "Options", Options
SaveSetting App.EXEName, "Startup", "FirstTime", False
SaveSetting App.EXEName, "Config", "ErrorText", ErrorText
SaveSetting App.EXEName, "Config", "ScriptPath", ScriptPath
SaveSetting App.EXEName, "Config", "Port", txtPort
SaveSetting App.EXEName, "Config", "Beep", chkBeep
SaveSetting App.EXEName, "Config", "Dupe", chkDupe
SaveSetting App.EXEName, "Config", "UseScript", cmdScript.BackColor

Unload frmAbout
Unload Me
End
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To 3
lblHelp(i).ForeColor = vbBlack
Next i
End Sub

Private Sub lblHelp_Click(Index As Integer)
Select Case Index
    Case 0
        MsgBox "This is the operating port for your IP Stealer. If you enter port 80 (Default), you will only need your friends to enter 'http://MyIP'. Otherwise, if you entered a different port, for instance 123, they would need to type 'http://MyIP:123'", vbInformation, "Help - Port"
    Case 1
        MsgBox "When the user arrives at your fake website, their IP will be added, but they will wonder why the page is not loading. Send them some fake error message like a '404 File Not Found' error. That way, your friends will just think that it no longer exists." & vbCrLf & vbCrLf & _
        "If the SCRIPT button is coloured BLUE, then the script you selected will be used. If it is grey, then the error text will be used as the error message." & vbCrLf & _
        "*** Scripts are just HTML files, but small ones. They are used for bigger amounts of data to be sent.", vbInformation, "Help - Error Text"
    Case 2
        MsgBox "This will cause your computer to play a 'beep' sound when a new IP is stolen.", vbInformation, "Help - Beep"
    Case 3
        MsgBox "This will check for duplicate IP's. For instance, if your friend visits your IP Stealer address once, and then again, if this option is selected, it will not add his IP again.", vbInformation, "Help - Duplicate"
End Select
End Sub

Private Sub lblHelp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblHelp(Index).ForeColor = vbBlue
End Sub

Private Sub lvIPs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lvIPs.ListItems.Count = 0 Then Exit Sub 'Dont popup the menu if no IP's are there
If Button = vbRightButton Then
PopupMenu mnuPopUp
End If

End Sub

Private Sub mnuAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuCopyBrowser_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lvIPs.ListItems.Item(lvIPs.SelectedItem.Index).SubItems(2)
End Sub

Private Sub mnuCopyIP_Click()
txtAddress_DblClick 'Run that sub's code...
End Sub

Private Sub mnuCopyIPAddy_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lvIPs.SelectedItem
End Sub

Private Sub mnuCopyOS_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lvIPs.ListItems.Item(lvIPs.SelectedItem.Index).SubItems(3)
End Sub

Private Sub mnuCopyTime_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lvIPs.ListItems.Item(lvIPs.SelectedItem.Index).SubItems(1)
End Sub

Private Sub mnuExit_Click()
Unload frmAbout
Unload Me
End
End Sub

Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)
For i = 0 To 99
If IPData(i).FreeSocket = True Then
IPData(i).FreeSocket = False
sckServer(i).Accept requestID
Exit Sub
End If
Next i
End Sub

Private Sub sckServer_Close(Index As Integer)
IPData(Index).FreeSocket = True
sckServer(Index).Close
End Sub

Private Sub Timer1_Timer()
'lvIPs.ColumnHeaders.Item(1).Left = 1349
'lvIPs.ColumnHeaders.Item(2).Width = 1305
'lvIPs.ColumnHeaders.Item(3).Width = 1305
'lvIPs.ColumnHeaders.Item(4).Width = 1440

Me.Caption = " 1-" & lvIPs.ColumnHeaders.Item(1).Width & _
" 2-" & lvIPs.ColumnHeaders.Item(2).Width & _
 " 3-" & lvIPs.ColumnHeaders.Item(3).Width & _
 " 4-" & lvIPs.ColumnHeaders.Item(4).Width

End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strScript As Variant
On Error GoTo SendData_Err
Dim strData As String, OSData, OSData1, OSData2
Dim strBrowser As String
Dim SplitData
sckServer(Index).GetData strData
SplitData = Split(strData, ";")
OSData1 = Split(strData, vbCrLf)
For i = LBound(OSData1) To UBound(OSData1)
If Left(OSData1(i), 11) = "User-Agent:" Then 'Found the sysinfo line
OSData2 = Split(OSData1(i), ";")
Exit For
End If
Next i

If Right(OSData2(2), 1) = ")" Then
OSData = Left(OSData2(2), Len(OSData2(2)) - 1)
Else
OSData = OSData2(2)
End If

strBrowser = OSData2(1)

Open App.Path & "\debug.txt" For Output As #1
Print #1, strData & vbCrLf
Close #1

If chkDupe Then
For i = 0 To Me.lvIPs.ListItems.Count - 1
If sckServer(Index).RemoteHostIP = lvIPs.ListItems.Item(1) Then
End If
Next i
Else
AddNewIP sckServer(Index).RemoteHostIP, OSData, strBrowser
If chkBeep Then
Beep
End If
End If


If cmdScript.BackColor = vbBlue Then
Open ScriptPath For Input As #1
strScript = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1
sckServer(Index).SendData strScript & Chr(10)
Else
sckServer(Index).SendData txtErrorText & Chr(10)
End If
DoEvents 'Waits for the data to be sent before closing the connection
sckServer(Index).Close

Exit Sub
SendData_Err:
MsgBox "ERROR! Err-" & Err.Number & " - Err.Description-" & Err.Description
'MsgBox "Error, the SCRIPT file has been moved, deleted, or no longer exists. Correct the problem, and then go online.", vbExclamation, "Missing Script!"
cmdOffline_Click
Exit Sub
End Sub

Private Sub txtAddress_DblClick()
'txtAddress.BackColor = vbRed
txtAddress.ForeColor = vbRed
Clipboard.Clear
Clipboard.SetText txtAddress
MsgBox "IP Copied to clipboard.", , "IP Copied"
lvIPs.SetFocus
txtAddress.ForeColor = vbBlue
End Sub
