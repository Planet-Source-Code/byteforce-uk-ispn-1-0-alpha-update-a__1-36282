VERSION 5.00
Begin VB.Form frmIM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(Unknown) - Conversation"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmIM.frx":26D0
   ScaleHeight     =   4635
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   675
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "localuser@127.0.0.1:210"
      Top             =   60
      Width           =   3840
   End
   Begin VB.TextBox txtIMLine 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   3915
      Width           =   3300
   End
   Begin VB.TextBox txtContact 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   675
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Unknown@127.0.0.1:210"
      Top             =   285
      Width           =   3840
   End
   Begin VB.ListBox lstConversation 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000C0&
      Height          =   3390
      ItemData        =   "frmIM.frx":48BEA
      Left            =   45
      List            =   "frmIM.frx":48C18
      TabIndex        =   2
      Top             =   540
      Width           =   4500
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   120
      Left            =   30
      TabIndex        =   6
      Top             =   660
      Width           =   4530
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Local:"
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   90
      TabIndex        =   5
      Top             =   60
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remote:"
      ForeColor       =   &H00C00000&
      Height          =   270
      Left            =   60
      TabIndex        =   3
      Top             =   285
      Width           =   615
   End
   Begin VB.Image imgPress 
      Height          =   690
      Left            =   2505
      Picture         =   "frmIM.frx":48C88
      Top             =   5490
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgHot 
      Height          =   690
      Left            =   1290
      Picture         =   "frmIM.frx":494D9
      Top             =   5490
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgNorm 
      Height          =   690
      Left            =   75
      Picture         =   "frmIM.frx":49D5D
      Top             =   5490
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgSend 
      Height          =   720
      Left            =   3345
      Picture         =   "frmIM.frx":4A5A4
      Stretch         =   -1  'True
      Top             =   3930
      Width           =   1215
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Public Sub SetState(DisableControls As Boolean)
'Enables \ Disables Instant Messenging controls
Select Case DisableControls

    Case True
    
        'Disable Controls
        imgSend.Enabled = False
        txtIMLine = " User unavailable"
        'txtIMLine.MultiLine = True
        txtIMLine.Locked = True
        txtIMLine.ForeColor = vbRed
        
    Case False
        
        'Enable Controls
        imgSend.Enabled = True
        txtIMLine = ""
        'txtIMLine.MultiLine = False
        txtIMLine.ForeColor = &H80000008
        txtIMLine.Locked = False
        
End Select

End Sub

Public Sub DisplayIM(UserHandle As String, AttachedServer As String, Optional TextToAdd As String)
'Check for null handle
If UserHandle = "" Then MsgBox "Unspecified user handle. Please specify a handle to connect to.", 16: Exit Sub
If AttachedServer = "" Then MsgBox "Unspecified server. Please specify a server to point to.", 16: Exit Sub

'Show Window
Me.Show

'Set Foreground window
PROBas.SetForegroundWindow Me.hWnd


'Check to see if there was a texttoadd parameter specified
If Not TextToAdd = "" Then
    
    'text is specified so add to conversation listbox
    lstConversation.AddItem TrimHandle(UserHandle) & ": " & TextToAdd
    lstConversation.ListIndex = lstConversation.ListCount - 1

    'play im sound
    If Right(App.Path, 1) = "\" Then pfile = App.Path Else pfile = App.Path & "\"
    PROBas.sndPlaySound pfile & "im.wav", PROBas.SND_ASYNC
    
End If

'Set the IM window caption
If Not Me.Caption = UserHandle & " - Conversation" Then Me.Caption = UserHandle & " - Conversation"

'The local user handle label
If Not txtFrom = TrimHandle(modISPN.ISPN_LocalHandle) & GetAttachedServer(True) Then txtFrom = TrimHandle(modISPN.ISPN_LocalHandle) & GetAttachedServer(True)

'The contact to send IMs to is already set to the correct destination
If txtContact = UserHandle & "@" & AttachedServer Then Exit Sub

'Update contact
txtContact = UserHandle & "@" & AttachedServer

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSend.Picture = imgNorm
End Sub

Private Sub imgSend_Click()
'Check for null inputs
If txtIMLine = "" Then
    
    Beep
    Exit Sub

End If

If txtContact = "" Then
    
    MsgBox "Contact is not specified. Please specify a contact to send a message to.", 16
    Exit Sub

End If

'Add what is being sent to the clients own display..
lstConversation.AddItem modISPN.ISPN_LocalHandle & ": " & txtIMLine
lstConversation.ListIndex = lstConversation.ListCount - 1

'Send the IM to the server for casting
modIMHandler.SendIMToHandle txtContact, txtIMLine

On Error Resume Next
'Select the text just typed..
txtIMLine.SelStart = 0
txtIMLine.SelLength = Len(txtIMLine)
txtIMLine.SetFocus
End Sub

Private Sub imgSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSend.Picture = imgPress
End Sub

Private Sub imgSend_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Exit Sub
If Button = 2 Then Exit Sub
imgSend.Picture = imgHot
End Sub


Private Sub lstConversation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSend.Picture = imgNorm
End Sub

Private Sub txtIMLine_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then imgSend_Click
End Sub

Private Sub txtIMLine_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSend.Picture = imgNorm
End Sub

