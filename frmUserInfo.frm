VERSION 5.00
Begin VB.Form frmUserInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "(Unknown) - Information"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   420
      Left            =   4200
      TabIndex        =   4
      ToolTipText     =   "Close this window"
      Top             =   6045
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   4050
      Left            =   165
      TabIndex        =   3
      Top             =   1890
      Width           =   5205
      Begin VB.CommandButton cmdViewProfile 
         Enabled         =   0   'False
         Height          =   255
         Left            =   2310
         MouseIcon       =   "frmUserInfo.frx":038A
         MousePointer    =   99  'Custom
         Picture         =   "frmUserInfo.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3630
         Width           =   270
      End
      Begin VB.CommandButton cmdSendiMail 
         Enabled         =   0   'False
         Height          =   255
         Left            =   2310
         MouseIcon       =   "frmUserInfo.frx":0A1E
         MousePointer    =   99  'Custom
         Picture         =   "frmUserInfo.frx":0D28
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2970
         Width           =   270
      End
      Begin VB.CommandButton cmdStartIM 
         Enabled         =   0   'False
         Height          =   255
         Left            =   2310
         MouseIcon       =   "frmUserInfo.frx":10B2
         MousePointer    =   99  'Custom
         Picture         =   "frmUserInfo.frx":13BC
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2355
         Width           =   270
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   4815
         Picture         =   "frmUserInfo.frx":1746
         Top             =   2505
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   4815
         Picture         =   "frmUserInfo.frx":1AD0
         Top             =   2790
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label hlnk_profile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View users profile"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2655
         MouseIcon       =   "frmUserInfo.frx":1E5A
         MousePointer    =   99  'Custom
         TabIndex        =   26
         ToolTipText     =   "View this users personal information and photo [if given]"
         Top             =   3645
         Width           =   1335
      End
      Begin VB.Label hlnk_imail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send iMail to user"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2655
         MouseIcon       =   "frmUserInfo.frx":2164
         MousePointer    =   99  'Custom
         TabIndex        =   25
         ToolTipText     =   "Send ISPN Mail to this user"
         Top             =   2985
         Width           =   1305
      End
      Begin VB.Label hlnk_sendIM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send this user an IM"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2655
         MouseIcon       =   "frmUserInfo.frx":246E
         MousePointer    =   99  'Custom
         TabIndex        =   24
         ToolTipText     =   "Send an Instant Message to this user"
         Top             =   2355
         Width           =   1500
      End
      Begin VB.Image imgServiceStat 
         Height          =   240
         Index           =   2
         Left            =   420
         Picture         =   "frmUserInfo.frx":2778
         Top             =   3660
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblServiceStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   705
         TabIndex        =   20
         Top             =   3660
         Width           =   1080
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   135
         Picture         =   "frmUserInfo.frx":2B02
         Stretch         =   -1  'True
         Top             =   3375
         Width           =   240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Profile:"
         Height          =   240
         Index           =   8
         Left            =   450
         TabIndex        =   19
         Top             =   3390
         Width           =   930
      End
      Begin VB.Image imgServiceStat 
         Height          =   240
         Index           =   1
         Left            =   420
         Picture         =   "frmUserInfo.frx":55B2
         Top             =   3000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblServiceStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   705
         TabIndex        =   18
         Top             =   3000
         Width           =   1080
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   135
         Picture         =   "frmUserInfo.frx":593C
         Top             =   2715
         Width           =   240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "iMail:"
         Height          =   240
         Index           =   7
         Left            =   450
         TabIndex        =   17
         Top             =   2730
         Width           =   390
      End
      Begin VB.Image imgServiceStat 
         Height          =   240
         Index           =   0
         Left            =   420
         Picture         =   "frmUserInfo.frx":5CC6
         Top             =   2370
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgServiceStatLoading 
         Height          =   240
         Left            =   4815
         Picture         =   "frmUserInfo.frx":6050
         Top             =   3390
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgServiceStatOn 
         Height          =   240
         Left            =   4815
         Picture         =   "frmUserInfo.frx":63DA
         Top             =   3135
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgServiceStatOff 
         Height          =   240
         Left            =   4815
         Picture         =   "frmUserInfo.frx":6764
         Top             =   3660
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblServiceStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   705
         TabIndex        =   16
         Top             =   2370
         Width           =   1080
      End
      Begin VB.Image imgService 
         Height          =   240
         Left            =   135
         Picture         =   "frmUserInfo.frx":6AEE
         Top             =   2085
         Width           =   240
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instant Messaging:"
         Height          =   240
         Index           =   6
         Left            =   450
         TabIndex        =   15
         Top             =   2100
         Width           =   1425
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "General Information"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enabled Services (Access was denied)"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1785
         Width           =   2790
      End
      Begin VB.Label lblLogonCount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   12
         Top             =   1425
         Width           =   1080
      End
      Begin VB.Label lblAccountCreated 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   11
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label lblFullHandle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   10
         Top             =   885
         Width           =   1080
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please wait..."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1455
         TabIndex        =   9
         Top             =   615
         Width           =   1080
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Logon Count:"
         Height          =   240
         Index           =   3
         Left            =   195
         TabIndex        =   8
         Top             =   1425
         Width           =   1005
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Made:"
         Height          =   240
         Index           =   2
         Left            =   195
         TabIndex        =   7
         Top             =   1155
         Width           =   1110
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full handle:"
         Height          =   240
         Index           =   1
         Left            =   195
         TabIndex        =   6
         Top             =   885
         Width           =   825
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User:"
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   615
         Width           =   405
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   315
      Picture         =   "frmUserInfo.frx":91BE
      Top             =   1140
      Width           =   360
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000017&
      Height          =   825
      Left            =   165
      TabIndex        =   2
      Top             =   930
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmUserInfo.frx":BC6E
      ForeColor       =   &H80000017&
      Height          =   825
      Left            =   810
      TabIndex        =   1
      Top             =   930
      Width           =   4560
   End
   Begin VB.Line Line1 
      X1              =   -195
      X2              =   5805
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Image imgCaption 
      Height          =   480
      Left            =   120
      Picture         =   "frmUserInfo.frx":BD05
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   705
      TabIndex        =   0
      Top             =   165
      Width           =   2850
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -435
      Picture         =   "frmUserInfo.frx":C9CF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Dim local_HANDLE As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub UpdateGeneralInfo(sAccMade As String, sLgnCount As String, Optional sTargetHandle As String)
'**Update the general information**

'Bring to front\show
Me.Show

If Not sTargetHandle = "" Then local_HANDLE = sTargetHandle

If Not Me.Caption = local_HANDLE & " - Information" Then Me.Caption = local_HANDLE & " - Information"

lblUser = TrimHandle(local_HANDLE)
If CheckNull(lblUser, True) = True Then lblUser = "Access was denied or unavailable.": lblUser.ForeColor = vbRed

lblFullHandle = TrimHandle(local_HANDLE) & GetAttachedServer(True)
If CheckNull(lblFullHandle, True) = True Then lblFullHandle = "Access was denied or unavailable.": lblFullHandle.ForeColor = vbRed

lblAccountCreated = sAccMade
If CheckNull(lblAccountCreated, True) = True Then lblAccountCreated = "Access was denied or unavailable.": lblAccountCreated.ForeColor = vbRed

lblLogonCount = sLgnCount
If CheckNull(lblLogonCount, True) = True Then lblUser = "Access was denied or unavailable.": lblUser.ForeColor = vbRed


End Sub

Private Sub cmdSendiMail_Click()
hlnk_imail_Click
End Sub

Private Sub cmdStartIM_Click()
hlnk_sendIM_Click
End Sub

Private Sub cmdViewProfile_Click()
hlnk_profile_Click
End Sub

Public Sub GetUserInfo(sTargetHandle As String)
'**Gets information about a specifed user (sTargethandle) from the server**

'Check for null input
If CheckNull(sTargetHandle, True) = True Then MsgBox "Please select a valid user. The action was cancelled.", 48, "User Information": Exit Sub

local_HANDLE = sTargetHandle

'Set window caption
If Not Me.Caption = local_HANDLE & " - Information" Then Me.Caption = local_HANDLE & " - Information"

'Show form
Me.Show

'Request information
Call modProfileHandler.RequestGeneralInfo(sTargetHandle)

End Sub

Private Sub hlnk_imail_Click()
ShowToDo
End Sub

Private Sub hlnk_profile_Click()
ShowToDo
End Sub

Private Sub hlnk_sendIM_Click()
ShowToDo
End Sub
