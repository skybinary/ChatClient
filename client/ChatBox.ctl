VERSION 5.00
Begin VB.UserControl ChatBox 
   BackStyle       =   0  'Transparent
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9000
   ScaleHeight     =   4290
   ScaleWidth      =   9000
   ToolboxBitmap   =   "ChatBox.ctx":0000
   Begin VB.VScrollBar VScroll1 
      Height          =   3855
      Left            =   6480
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Wrapper 
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      ScaleHeight     =   3915
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin prjFN33Client.ChatMsg ChatMsgList 
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   -615
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   450
      End
   End
End
Attribute VB_Name = "ChatBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const minWidth As Long = 480
Const minHeight As Long = 480

Public Event clicky(ByVal Index As Integer)

Private topOffset As Integer
Private itemOffset As Integer
Private itemsOnScreen As Integer

Private chatCombinedHeight As Long

Private Sub ChatMsg1_GotFocus(Index As Integer)

End Sub

Private Sub ChatMsgList_msgClicky(Index As Integer)
    RaiseEvent clicky(ByVal Index)
End Sub

Private Sub UserControl_Initialize()
    itemOffset = 1
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width < minWidth Then UserControl.Width = minWidth
    If UserControl.Height < minHeight Then UserControl.Height = minHeight
    Wrapper.Width = UserControl.Width - 240
    Wrapper.Height = UserControl.Height - 240
    
    VScroll1.Left = (Wrapper.Width + 240) - VScroll1.Width
    VScroll1.Height = Wrapper.Height
End Sub

Private Sub VScroll1_Change()
    itemOffset = (VScroll1.Value)
   ' Debug.Print itemOffset
    Wrapper_Resize
End Sub

Private Sub Wrapper_Resize()
    Dim thisWidth As Long
    Dim vw As Integer
    Dim msg As ChatMsg
    Dim x As Integer
    Dim y As Integer
    Dim mypos As Integer
    mypos = 0
    ChatMsgList(0).Top = 0 - ChatMsgList(0).Height
    ChatMsgList(0).Visible = False
    itemsOnScreen = 0
    If itemOffset > -1 Then
        For x = 0 To itemOffset - 1
            ChatMsgList(x).Visible = False
            ChatMsgList(x).Top = 0 - ChatMsgList(x).Height
        Next
        For y = itemOffset To ChatMsgList.Count - 1
            ChatMsgList(y).Width = thisWidth
            ChatMsgList(y).Top = mypos
            ChatMsgList(y).Visible = True
        
            If mypos + ChatMsgList(y).Height > UserControl.Height Then Exit For
            itemsOnScreen = itemsOnScreen + 1
            mypos = mypos + ChatMsgList(y).Height
        Next
        For x = y To ChatMsgList.Count - 1
            ChatMsgList(x).Visible = False
        Next
    End If
    Debug.Print itemsOnScreen, ChatMsgList.Count
    If VScroll1.Visible Then
        vw = VScroll1.Width - 60
    Else
        vw = 60
        topOffset = 0
    End If
    thisWidth = Wrapper.Width - vw
End Sub

'////////// functions ?

Public Sub addNew(ByVal contents As String)
    Load ChatMsgList(ChatMsgList.UBound + 1)
    ChatMsgList(ChatMsgList.UBound).TestIndex = ChatMsgList.UBound
    Wrapper_Resize
End Sub

Public Sub removeMsg(ByVal Index As Integer)
    Dim mcount As Integer
    Dim a As Integer
    For mcount = Index To ChatMsgList.UBound - 1
        swapEm mcount, mcount + 1
        'Set ChatMsgList(mcount) = ChatMsgList(mcount + 1)
    Next
    Unload ChatMsgList(ChatMsgList.UBound)
    Wrapper_Resize
End Sub

Private Function calcMsgsHeight() As Long
    Dim totty As Long
    Dim x As Integer
    For x = 0 To ChatMsgList.Count - 1
        If ChatMsgList(x).Visible Then totty = totty + ChatMsgList(x).Height
    Next
    calcMsgsHeight = totty
End Function

Private Sub swapEm(ByVal dest As Integer, src As Integer)
    ChatMsgList(dest).TestIndex = ChatMsgList(src).TestIndex
    
End Sub

Sub showHideScroll(ByVal tf As Boolean)
    VScroll1.Visible = tf
    
End Sub
