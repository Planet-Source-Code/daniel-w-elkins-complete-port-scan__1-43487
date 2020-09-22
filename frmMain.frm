VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{9B1E48ED-8018-11D3-B75D-006097A1EBF0}#1.0#0"; "DNS.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PortScan"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   7290
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status : Idle."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   375
      Left            =   120
      TabIndex        =   26
      Top             =   6840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top"
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   375
      Left            =   3840
      TabIndex        =   24
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Caption         =   " Results "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   4935
      Begin DNSControl.DNS DNS 
         Left            =   1200
         Top             =   720
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin MSWinsockLib.Winsock Socket 
         Left            =   600
         Top             =   720
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ILRes 
         Left            =   2520
         Top             =   1200
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0ECA
               Key             =   "IP"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":131C
               Key             =   "Host"
            EndProperty
         EndProperty
      End
      Begin VB.CheckBox chkClear 
         Caption         =   "Clear on Session Start"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         ToolTipText     =   " If this box is checked, the results will be cleared everytime a scan is started "
         Top             =   2400
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   3600
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVRes 
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "ILRes"
         SmallIcons      =   "ILRes"
         ColHdrIcons     =   "ILRes"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IP Address"
            Object.Width           =   3493
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hostname"
            Object.Width           =   4630
            ImageIndex      =   2
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Enter Session Information "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   4935
      Begin VB.CommandButton cmdTimeout 
         Caption         =   "Timeout"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   29
         ToolTipText     =   " Change the Timeout Interval "
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkHost 
         Caption         =   "Resolve Hostnames"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enter Port :"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Enter an IP Range to Scan "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   4935
      Begin VB.TextBox txtE1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtE2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtE3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtE4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtS4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   3
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtS3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtS2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtS1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   3
         TabIndex        =   0
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "             .              .               ."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   650
         Width           =   3495
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "             .              .               ."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   13
         Top             =   290
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "End IP :"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Start IP :"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   0
      TabIndex        =   9
      Top             =   1440
      Width           =   5175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      Picture         =   "frmMain.frx":176E
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   8
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblProg 
      Alignment       =   2  'Center
      Caption         =   "0 / 0"
      Height          =   255
      Left            =   1440
      TabIndex        =   28
      Top             =   6480
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bResHost As Boolean
Dim bClearRes As Boolean
Dim lIPCnt As Long
Dim bKeepGoing As Boolean
Dim lTimeout As Long
Dim iResponse As Integer
Dim iCnt As Long
Sub NumOnly(KeyAscii As Integer)
KeyAscii = IIf(Not KeyAscii = 8 And Not IsNumeric(Chr(KeyAscii)), 0, KeyAscii)
End Sub
Sub SelAll(oTxtBox As TextBox)
oTxtBox.SelStart = 0
oTxtBox.SelLength = Len(oTxtBox.Text)
End Sub
Sub SaveSettings()
Dim sIP1 As String, sIP2 As String: sIP1 = Empty: sIP2 = Empty
sIP1 = txtS1.Text & "." & txtS2.Text & "." & txtS3.Text & "." & txtS4.Text
sIP2 = txtE1.Text & "." & txtE2.Text & "." & txtE3.Text & "." & txtE4.Text
SaveSetting App.ProductName, "Main", "StartIP", sIP1
SaveSetting App.ProductName, "Main", "EndIP", sIP2
sIP1 = Empty: sIP2 = Empty
SaveSetting App.ProductName, "Main", "StartPort", txtPort.Text
SaveSetting App.ProductName, "Main", "Hostnames", chkHost.Value
SaveSetting App.ProductName, "Main", "ClearRes", chkClear.Value
SaveSetting App.ProductName, "Main", "Timeout", Str$(lTimeout)
End Sub
Sub ReadSettings()
Dim sTmpIP As String, sTmp() As String: sTmpIP = Empty
sTmpIP = GetSetting(App.ProductName, "Main", "StartIP", "")
If Len(sTmpIP) = 0 Then
txtS1.Text = "": txtS2.Text = "": txtS3.Text = "": txtS4.Text = ""
Else
sTmp() = Split(sTmpIP, ".")
txtS1.Text = sTmp(0): txtS2.Text = sTmp(1): txtS3.Text = sTmp(2): txtS4.Text = sTmp(3)
End If
sTmpIP = Empty
sTmpIP = GetSetting(App.ProductName, "Main", "EndIP", "")
If Len(sTmpIP) = 0 Then
txtE1.Text = "": txtE2.Text = "": txtE3.Text = "": txtE4.Text = ""
Else
sTmp() = Split(sTmpIP, ".")
txtE1.Text = sTmp(0): txtE2.Text = sTmp(1): txtE3.Text = sTmp(2): txtE4.Text = sTmp(3)
End If
txtPort.Text = GetSetting(App.ProductName, "Main", "StartPort", "")
chkHost.Value = GetSetting(App.ProductName, "Main", "Hostnames", 0)
chkClear.Value = GetSetting(App.ProductName, "Main", "ClearRes", 0)
lTimeout = Val(GetSetting(App.ProductName, "Main", "Timeout", 50000))
End Sub
Function ValClass(sStartClass As String, sEndClass As String) As Boolean
ValClass = False
If Val(sStartClass) > 255 Then
ValClass = False
ElseIf Val(sEndClass) > 255 Then
ValClass = False
ElseIf Val(sStartClass) > Val(sEndClass) Then
ValClass = False
Else
ValClass = True
End If
End Function
Private Sub cmdStart_Click()
If Len(txtS1.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtS1.SetFocus
Exit Sub

ElseIf Len(txtS2.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtS2.SetFocus
Exit Sub

ElseIf Len(txtS3.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtS3.SetFocus
Exit Sub

ElseIf Len(txtS4.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtS4.SetFocus
Exit Sub

ElseIf Len(txtE1.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtE1.SetFocus
Exit Sub

ElseIf Len(txtE2.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtE2.SetFocus
Exit Sub

ElseIf Len(txtE3.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtE3.SetFocus
Exit Sub

ElseIf Len(txtE4.Text) = 0 Then
MsgBox "Enter a complete IP range", vbCritical, "IP Range Required"
txtE4.SetFocus
Exit Sub

ElseIf Not ValClass(txtS1.Text, txtE1.Text) Then
MsgBox "Invalid IP range", vbCritical, "IP Range Required"
txtS1.SetFocus
Exit Sub

ElseIf Not ValClass(txtS2.Text, txtE2.Text) Then
MsgBox "Invalid IP range", vbCritical, "IP Range Required"
txtS2.SetFocus
Exit Sub

ElseIf Not ValClass(txtS3.Text, txtE3.Text) Then
MsgBox "Invalid IP range", vbCritical, "IP Range Required"
txtS3.SetFocus
Exit Sub

ElseIf Not ValClass(txtS4.Text, txtE4.Text) Then
MsgBox "Invalid IP range", vbCritical, "IP Range Required"
txtS4.SetFocus
Exit Sub

ElseIf Len(txtPort.Text) = 0 Then
MsgBox "Enter a port to scan", vbCritical, "Port Required"
txtPort.SetFocus
Exit Sub
Else

bResHost = chkHost.Value
bClearRes = chkClear.Value
If bClearRes Then
LVRes.ListItems.Clear
End If
bKeepGoing = True
Call ResetFields

Call SaveSettings

lIPCnt = Abs(Val(txtE3.Text) - Val(txtS3.Text) + 1) * Abs(Val(txtE4.Text) - Val(txtS4.Text) + 1)
Call ScanIPs

End If
End Sub

Private Sub cmdStop_Click()
bKeepGoing = False
End Sub

Private Sub cmdTimeout_Click()
Dim sTO As String: sTO = Empty
If lTimeout = 0 Then
lTimeout = 50000
End If
sTO = InputBox("Enter a new timeout interval in number of CPU cycles : ", "Current Is (" & lTimeout & ")")
If Len(sTO) = 0 Then
Exit Sub
ElseIf Val(sTO) = 0 Then
MsgBox "Interval must be between 1 and 200,000", vbCritical, "Invalid Interval"
Exit Sub
ElseIf Val(sTO) > 200000 Then
MsgBox "Interval must be between 1 and 200,000", vbCritical, "Invalid Interval"
Exit Sub
End If
lTimeout = Val(sTO)
sTO = Empty
End Sub

Private Sub Form_Load()
Call ReadSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bKeepGoing = False
End
End Sub

Private Sub Socket_Connect()
iResponse = 1
Dim iNum As Integer: iNum = 0
Dim sTmpDomain As String: sTmpDomain = Empty
iNum = LVRes.ListItems.Count + 1
LVRes.ListItems.Add , , Socket.RemoteHostIP, , "IP"
If bResHost Then
sTmpDomain = DNS.AddressToName(Socket.RemoteHostIP)
LVRes.ListItems(iNum).ListSubItems.Add , , sTmpDomain, "Host", " Domain Name : " & sTmpDomain & Space$(1)
End If
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
iResponse = 2
End Sub

Private Sub txtE1_Change()
txtS1.Text = txtE1.Text
End Sub

Private Sub txtE1_GotFocus()
Call SelAll(txtE1)
End Sub

Private Sub txtE1_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
End Sub

Private Sub txtE2_Change()
txtS2.Text = txtE2.Text
End Sub

Private Sub txtE2_GotFocus()
Call SelAll(txtE2)
End Sub

Private Sub txtE2_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtE2.Text) = 0 Then
txtE1.SetFocus
End If

End Sub

Private Sub txtE3_GotFocus()
Call SelAll(txtE3)
End Sub

Private Sub txtE3_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtE3.Text) = 0 Then
txtE2.SetFocus
End If

End Sub

Private Sub txtE4_GotFocus()
Call SelAll(txtE4)
End Sub

Private Sub txtE4_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtE4.Text) = 0 Then
txtE3.SetFocus
End If
End Sub

Private Sub txtS1_Change()
txtE1.Text = txtS1.Text
End Sub

Private Sub txtS1_GotFocus()
Call SelAll(txtS1)
End Sub

Private Sub txtS1_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
End Sub

Private Sub txtS2_Change()
txtE2.Text = txtS2.Text
End Sub

Private Sub txtS2_GotFocus()
Call SelAll(txtS2)
End Sub

Private Sub txtS2_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtS2.Text) = 0 Then
txtS1.SetFocus
End If
End Sub

Private Sub txtS3_GotFocus()
Call SelAll(txtS3)
End Sub

Private Sub txtS3_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtS3.Text) = 0 Then
txtS2.SetFocus
End If

End Sub

Private Sub txtS4_GotFocus()
Call SelAll(txtS4)
End Sub

Private Sub txtS4_KeyPress(KeyAscii As Integer)
Call NumOnly(KeyAscii)
If KeyAscii = 8 And Len(txtS4.Text) = 0 Then
txtS3.SetFocus
End If

End Sub

Sub ResetFields()
Bar.Value = 0
lblProg.Caption = "0 / 0"
End Sub

Function SecTime(ByVal Seconds As Single) As String
Dim sHours As Single, sMins As Single, sSecs As Single
sHours = Seconds \ (60 * 60)
sMins = (Seconds - sHours * (60 * 60)) \ 60
sSecs = Seconds Mod 60
SecTime = Format(sHours, "00") & ":" & Format(sMins, "00") & ":" & Format(sSecs, "00")
End Function

Sub ScanIPs()
Dim A As Integer, B As Integer, C As Integer, sTO As Long: iCnt = 0: sTO = 3000
Dim iSPort As Long, iEPort As Long
Dim sTimeStart As Single
Dim sCurIP As String, sTimeLeft As String: sCurIP = Empty: sTimeLeft = Empty
sTimeStart = Timer
iSPort = Val(txtPort.Text)

For A = Val(txtS3.Text) To Val(txtE3.Text)
    For B = Val(txtS4.Text) To Val(txtE4.Text)

iCnt = iCnt + 1
If Not bKeepGoing Then
StatusBar.SimpleText = "Status : Session Cancelled."
Call ResetFields
Exit For
End If

sCurIP = txtS1.Text & "." & txtS2.Text & "." & A & "." & B
sTimeLeft = SecTime((Timer - sTimeStart) * (lIPCnt - iCnt) / iCnt)
StatusBar.SimpleText = "Status : Current IP - " & sCurIP & " (" & sTimeLeft & ") [" & Int(Bar.Value) & " %] . . ."
lblProg.Caption = iCnt & " / " & lIPCnt
Socket.Close
Socket.Connect sCurIP, txtPort.Text
iResponse = 0
sTO = 0
On Error Resume Next
Bar.Value = Bar.Value + 100 / lIPCnt
Do
DoEvents
sTO = sTO + 1
Loop Until sTO = lTimeout Or iResponse > 0
DoEvents
DoEvents
Next
Next
Bar.Value = 100
lblProg.Caption = lIPCnt & " / " & lIPCnt
StatusBar.SimpleText = "Status : Session Complete."
End Sub


