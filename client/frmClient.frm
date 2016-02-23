VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yo Momma"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstUsers 
      Height          =   3375
      ItemData        =   "frmClient.frx":0000
      Left            =   4680
      List            =   "frmClient.frx":0002
      TabIndex        =   8
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Text            =   "NuckFuggets"
      Top             =   3480
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock sockMain 
      Left            =   4200
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Hello ;)"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "send"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtStatus 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "nerd33.com"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Users"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblNick 
      Caption         =   "Nick"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Shape Indicator 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Host"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ## enumbs have to be up top
Private Const seper = "#|"
Private Const weper = "|#"

Public Enum eSpecialFolders
    SpecialFolder_AppData = &H1A    ' current widnows user on computer or network (98 or later)
    SpecialFolder_CommonAppData = &H23 ' for all widnows users on this comp (2000 or later)
    SpecialFolder_LocalAppData = &H1C ' current user on this comp only (2000 or later)
    SpecialFolder_Documents = &H5  ' current widnows user docments
End Enum
Private gLoadedNick As String
Private gLoadedHost As String
Private gConfigFile As String

Public Function SpecialFolder(pFolder As eSpecialFolders) As String
    Dim objShell As Object
    Dim objFolder As Object
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.namespace(CLng(pFolder))
    If (Not objFolder Is Nothing) Then SpecialFolder = objFolder.Self.Path
    Set objFolder = Nothing
    Set objShell = Nothing
    If SpecialFolder = "" Then Err.Raise 513, "SpecialFolder", "The folder path could not be detected"
End Function

Private Sub cmdConnect_Click()
    Dim anArray As Variant
    Dim port As String
    Dim host As String
    anArray = Split(txtHost.Text, ":")
    If UBound(anArray) < 2 Then
        host = anArray(0)
        port = "25565"
    Else
        host = anArray(0)
        port = anArray(1)
    End If
    sockMain.RemoteHost = host
    sockMain.RemotePort = port
    sockMain.Connect
End Sub

Private Sub cmdSend_Click()
    If sockMain.State = sckConnected Then
        sockMain.SendData "MSG:" & txtSend.Text
        ' comment out below as server sends msg back
  '      txtStatus.Text = txtStatus.Text & txtNick.Text & ":" & txtSend.Text & vbCrLf
        txtSend.Text = ""
    End If
End Sub

Private Sub Form_Load()
 Dim strConfigFile As String
 Dim strConfigFolder As String
 
 strConfigFolder = SpecialFolder(SpecialFolder_LocalAppData) & "\FN33"
 strConfigFile = strConfigFolder & "\config.ini"
 Dim sFileText As String
 Dim sFinal As String
 Dim iReadFileNo As Integer
 Dim iWriteFileNo As Integer
 Dim erno As Integer
 Dim infinateLoop As Integer
 Dim anArray As Variant
 
letsReadAgain:
 iReadFileNo = FreeFile
 On Error GoTo readError
 Open strConfigFile For Input As iReadFileNo
    Do While Not EOF(iReadFileNo)
        Input #iReadFileNo, sFileText
        sFinal = sFinal & sFileText & vbCrLf
    Loop
 Close #iReadFileNo
 anArray = Split(sFinal, vbCrLf)
 Dim Itum As Variant
 gLoadedNick = "-"
 gLoadedHost = "-"
 For Each Itum In anArray
    Dim cnArray As Variant
    cnArray = Split(Itum, seper)
    If UBound(cnArray) > 0 Then
        If cnArray(0) = "nick" Then gLoadedNick = cnArray(1)
        If cnArray(0) = "host" Then gLoadedHost = cnArray(1)
    End If
 Next
 txtNick.Text = gLoadedNick
 txtHost.Text = gLoadedHost
 gConfigFile = strConfigFile
 Exit Sub
readError:
erno = Err.Number
On Error GoTo 0
If erno = 76 Then
    ' path not found so create ???
    On Error GoTo createFolderError
    MkDir strConfigFolder
    GoTo createFile ' obviously if the folder doesnt exist how can the file, so go straight to create it
 ElseIf erno = 53 Then
    ' file not found so create it????
    Stop
 End If
 GoTo exitThis
 
createFolderError:
    MsgBox (Err.Number)
    GoTo exitThis
 
createFileError:
    MsgBox (Err.Number)
    GoTo exitThis
    
createFile:
    Dim status As Boolean
    status = writeSettingsToFile
    
 GoTo exitThis
 
exitThis:
End Sub

Private Sub sockMain_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim anArray As Variant
    
    sockMain.GetData strData, vbString
    Debug.Print "[" & strData & "]" '10053
    If InStr(1, strData, "CMD:") = 1 Then
        anArray = Split(strData, "CMD:")
        If anArray(1) = "CONN" Then
'            txtSend.Locked = False
            Indicator.FillColor = &H80FF&
            Dim usrNick As String
            usrNick = "USR:" & txtNick.Text
            sockMain.SendData usrNick
            Debug.Print usrNick
        ElseIf anArray(1) = "10053" Then
            txtSend.Locked = True
            Indicator.FillColor = &HFF&
            txtStatus.Text = txtStatus.Text & "CONNECTION ABORTED (NICK IN USE)" & vbCrLf
        ElseIf anArray(1) = "CONN_SUCC" Then
            txtSend.Locked = False
            Indicator.FillColor = &HFF00&
            Dim usrNaick As String
            usrNick = "USR:" & txtNick.Text
            sockMain.SendData usrNick
            Debug.Print usrNick
        ElseIf anArray(1) = "QUIT" Then
            txtSend.Locked = True
            Indicator.FillColor = &HFF&
            txtStatus.Text = txtStatus.Text & "CONNECTION QUIT BY SERVER" & vbCrLf
        End If
        Debug.Print anArray(1)
    ElseIf InStr(1, strData, "LST:") = 1 Then
        anArray = Split(strData, "LST:")
        Dim bnArray As Variant
        bnArray = Split(anArray(1), "`~`")
        lstUsers.Clear
        Dim x As Integer
        Dim Item As Variant
        For Each Item In bnArray
            lstUsers.AddItem Item
        Next
    ElseIf InStr(1, strData, "MSG:") = 1 Then
        anArray = Split(strData, "MSG:")
        txtStatus.Text = txtStatus.Text & anArray(1) & vbCrLf
    End If
End Sub

Private Sub txtHost_LostFocus()
    If gLoadedHost <> txtHost.Text Then writeSettingsToFile
End Sub

Private Sub txtNick_LostFocus()
    If gLoadedNick <> txtNick.Text Then writeSettingsToFile
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSend_Click
End Sub

'Private Sub writeNewNickToFile(ByVal nick As String)
Private Function writeSettingsToFile()
    Dim iWriteFileNo As Integer
    iWriteFileNo = FreeFile
    On Error GoTo createFileError2
    Open gConfigFile For Output As #iWriteFileNo
        Print #iWriteFileNo, "host" & seper & txtHost.Text
        Print #iWriteFileNo, "nick" & seper & txtNick.Text
    Close #iWriteFileNo
    
    writeSettingsToFile = True
    
    Exit Function
    
createFileError2:

    writeSettingsToFile = False
End Function
