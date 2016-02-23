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
      TabIndex        =   10
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
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
      TabIndex        =   7
      Text            =   "Hello ;)"
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "send"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtStatus 
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "25565"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "nerd33.com"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Users"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblNick 
      Caption         =   "Nick"
      Height          =   255
      Left            =   240
      TabIndex        =   8
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
   Begin VB.Label Label2 
      Caption         =   "Port"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Host"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ## enumbs have to be up top
Public Enum eSpecialFolders
    SpecialFolder_AppData = &H1A    ' current widnows user on computer or network (98 or later)
    SpecialFolder_CommonAppData = &H23 ' for all widnows users on this comp (2000 or later)
    SpecialFolder_LocalAppData = &H1C ' current user on this comp only (2000 or later)
    SpecialFolder_Documents = &H5  ' current widnows user docments
End Enum
Private gLoadedNick As String
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
    sockMain.RemoteHost = txtHost.Text
    sockMain.RemotePort = txtPort.Text
    sockMain.Connect
End Sub

Private Sub cmdSend_Click()
    If sockMain.State = sckConnected Then
        sockMain.SendData "MSG:" & txtSend.Text
        txtStatus.Text = txtStatus.Text & txtNick.Text & ":" & txtSend.Text & vbCrLf
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
 For Each Itum In anArray
    If gLoadedNick = "-" Then gLoadedNick = Itum
    Debug.Print "{" & Itum & "}"
 Next
 txtNick.Text = gLoadedNick
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
    iWriteFileNo = FreeFile
    On Error GoTo createFileError
    Open strConfigFile For Output As #iWriteFileNo
        Write #iWriteFileNo, "erm" & vbCrLf
    Close #iWriteFileNo
    infinateLoop = infinateLoop + 1
    If infinateLoop < 3 Then
        GoTo letsReadAgain
    Else
        MsgBox ("uh oh we tried thrice!")
    End If
    
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
        If anArray(1) = "Connecting" Then
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
        ElseIf anArray(1) = "Connection Successful" Then
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

Private Sub txtNick_LostFocus()
    If gLoadedNick <> txtNick.Text Then writeNewNickToFile txtNick.Text
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdSend_Click
End Sub

Private Sub writeNewNickToFile(ByVal nick As String)
    Dim iWriteFileNo As Integer
    iWriteFileNo = FreeFile
    On Error GoTo createFileError2
    Open gConfigFile For Output As #iWriteFileNo
        Write #iWriteFileNo, nick & vbCrLf
    Close #iWriteFileNo
    Exit Sub
createFileError2:
        MsgBox ("error writing to file")
    
End Sub
