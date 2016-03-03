VERSION 5.00
Begin VB.UserControl ChatMsg 
   BackColor       =   &H8000000C&
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ScaleHeight     =   1500
   ScaleWidth      =   7155
   ToolboxBitmap   =   "ChatMsg.ctx":0000
   Begin VB.Label Nick 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nick"
      Height          =   195
      Left            =   435
      TabIndex        =   1
      Top             =   30
      Width           =   330
      WordWrap        =   -1  'True
   End
   Begin VB.Label Message 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label1"
      Height          =   195
      Left            =   870
      TabIndex        =   0
      Top             =   30
      Width           =   480
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "ChatMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private p_Testindex As Integer
Private p_Margin As Long

Const minWidth As Long = 480
Const minHeight As Long = 480

Private iCount As Integer
Private rMessages(4) As String
Private rNicks(4) As String

Public Event msgClicky()

Public Property Let TestIndex(ByVal newVal As Integer)
    p_Testindex = newVal
    
    Message.Caption = "[" & CStr(p_Testindex) & "] " & randomMessage
    Message.ToolTipText = Message.Caption
    Nick.Caption = randomNick
    Nick.ToolTipText = Nick.Caption
    
   ' Debug.Print Nick.Caption & " : " & Message.Caption & vbCrLf
End Property

Public Property Get TestIndex() As Integer
    TestIndex = p_Testindex
End Property

Public Property Let Margin(ByVal newVal As Long)
    If newVal <> p_Margin Then
        p_Margin = newVal
        UserControl_Resize
    End If
End Property

Public Property Get Margin() As Long
    Margin = p_Margin
End Property

Private Sub Message_Click()
    RaiseEvent msgClicky
End Sub

Private Sub Nick_Click()
    RaiseEvent msgClicky
End Sub

Private Sub UserControl_Click()
    RaiseEvent msgClicky
End Sub

' ////////

Private Function randomMessage() As String
    Randomize
    Dim r As Long
    r = 4 * Rnd
    randomMessage = rMessages(r)
End Function

Private Function randomNick() As String
    Randomize
    Dim r As Long
    r = 4 * Rnd
    randomNick = rNicks(r)
End Function

Private Sub UserControl_Initialize()
    rMessages(0) = "This is a teset"
    rMessages(1) = "FFS"
    rMessages(2) = "The quick brown fox jumped over the lazy dogs"
    rMessages(3) = "Now is the time for all good men and true to come to the aid of the party. "
    rMessages(4) = "Pack my box with five dozen liquor jugs"
    rNicks(0) = "Jane"
    rNicks(1) = "Jonathan"
    rNicks(2) = "Skooperdoopereffingdoo"
    rNicks(3) = "sky"
    rNicks(4) = "raspberry"
    
   ' iCount = iCount + 1
   ' Me.TestIndex = iCount
End Sub

Private Function maxiHeight(ByVal a1 As Long, ByVal a2 As Long) As Long
    If a1 > a2 Then maxiHeight = a1 Else maxiHeight = a2
End Function

Private Sub UserControl_Resize()
    Dim mw As Long
    Dim mh As Long
    If UserControl.Width < minWidth Then UserControl.Width = minWidth
    mw = UserControl.Width - (90 + Nick.Width)
    Message.Width = mw
    mh = maxiHeight(Message.Height, Nick.Height) + 60
'    Debug.Print mh
    UserControl.Height = mh
'    If UserControl.Height < minHeight Then UserControl.Height = minHeight
'    Debug.Print UserControl.Width & "," & UserControl.Height

    
    Nick.Left = 30
    Message.Left = 60 + Nick.Width
    
  '  Message.Height = UserControl.Height - 60
 '  Nick.Height = UserControl.Height - 60
End Sub

