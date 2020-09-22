VERSION 5.00
Begin VB.Form frmUsers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ISPN Server User Manager"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Users"
      Height          =   4935
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   4605
      Begin VB.CommandButton cmdPolicy 
         Caption         =   "User Policy"
         Enabled         =   0   'False
         Height          =   840
         Left            =   3023
         Picture         =   "frmUsers.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3855
         Width           =   1110
      End
      Begin VB.CommandButton cmdInvestigate 
         Caption         =   "Investigate"
         Enabled         =   0   'False
         Height          =   840
         Left            =   1793
         Picture         =   "frmUsers.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3855
         Width           =   1110
      End
      Begin VB.CommandButton cmdKick 
         Caption         =   "Logoff ISPN"
         Enabled         =   0   'False
         Height          =   840
         Left            =   548
         Picture         =   "frmUsers.frx":1994
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3855
         Width           =   1110
      End
      Begin VB.ListBox lstUsers 
         ForeColor       =   &H000040C0&
         Height          =   1740
         Left            =   105
         TabIndex        =   3
         Top             =   840
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Below you can choose tasks to carry out with the selected user (above)."
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   720
         TabIndex        =   6
         Top             =   3105
         Width           =   3795
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmUsers.frx":265E
         Top             =   3060
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   105
         X2              =   4500
         Y1              =   2925
         Y2              =   2925
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   105
         X2              =   4500
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Label lblCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a list of users currently signed into ISPN Server. Hit 'Refresh' to update the list."
         ForeColor       =   &H00800000&
         Height          =   510
         Left            =   735
         TabIndex        =   5
         Top             =   270
         Width           =   3795
      End
      Begin VB.Image imgCaption 
         Height          =   480
         Left            =   120
         Picture         =   "frmUsers.frx":3328
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblLoggedOn 
         BackStyle       =   0  'Transparent
         Caption         =   "0 User(s) logged on to 127.0.0.1"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   135
         TabIndex        =   4
         Top             =   2625
         Width           =   4320
      End
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2025
      TabIndex        =   1
      Top             =   5025
      Width           =   1230
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   3390
      TabIndex        =   0
      Top             =   5025
      Width           =   1230
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdKick_Click()
Dim tmpVarConv As Integer

tmpVarConv = lstUsers.ItemData(lstUsers.ListIndex)

Call frmServer.Server_Close(tmpVarConv)

Form_Load 'Refresh the form
lstUsers_Click 'Reset task control buttons
End Sub

Private Sub cmdRefresh_Click()
lstUsers.Clear
tmpNumberOfUsers = 0

'Get logged on users and add to list
For findusers = 0 To ISPN_TopLevelCtlId
    
    If Not ISPNUSER_MemberHandle(findusers) = "" Then
        
        'Add user to string that will be sent out to the client (contact list)
        lstUsers.AddItem ISPNUSER_MemberHandle(findusers) & " (" & frmServer.Server(findusers).RemoteHostIP & " on socket " & findusers & ")"
        lstUsers.ItemData(tmpNumberOfUsers) = findusers 'Set a tag so we can kick the user if we want..
        
        tmpNumberOfUsers = tmpNumberOfUsers + 1 'Update counter
    
    End If

Next findusers

lblLoggedOn = tmpNumberOfUsers & " User(s) logged on to " & frmServer.Server(0).LocalIP
End Sub

Private Sub Form_Load()
cmdRefresh_Click

End Sub

Private Sub lstUsers_Click()
If lstUsers = "" Then
    
    cmdKick.Enabled = False

Else
    
    cmdKick.Enabled = True

End If
End Sub
