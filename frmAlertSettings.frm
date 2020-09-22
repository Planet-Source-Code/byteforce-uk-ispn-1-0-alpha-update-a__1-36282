VERSION 5.00
Begin VB.Form frmAlertSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ISPN Alert Settings"
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
   Icon            =   "frmAlertSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAlertSettings.frx":038A
   ScaleHeight     =   6555
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   420
      Left            =   2865
      TabIndex        =   5
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4170
      TabIndex        =   6
      Top             =   6045
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   4050
      Left            =   150
      TabIndex        =   3
      Top             =   1890
      Width           =   5205
      Begin VB.CheckBox chkProfileViewAlert 
         Caption         =   "Profile view alert"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   19
         Top             =   2865
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.CommandButton cmdSubscribers 
         Caption         =   "Veiw Subscriber Alerts"
         Enabled         =   0   'False
         Height          =   345
         Left            =   510
         TabIndex        =   17
         Top             =   3540
         Width           =   2550
      End
      Begin VB.CheckBox chkIMAlert 
         Caption         =   "IM alert"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   14
         Top             =   2625
         Value           =   1  'Checked
         Width           =   930
      End
      Begin VB.CheckBox chkiMailAlert 
         Caption         =   "New iMail alert"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   13
         Top             =   2370
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkLogoffAlert 
         Caption         =   "User logoff alert"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   12
         Top             =   2115
         Width           =   1515
      End
      Begin VB.CheckBox chkLogonAlert 
         Caption         =   "User logon alert"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   11
         Top             =   1875
         Value           =   1  'Checked
         Width           =   1515
      End
      Begin VB.CheckBox chkNoNetworkAlerts 
         Caption         =   "I do not want to recieve network alerts."
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   495
         TabIndex        =   8
         Top             =   765
         Width           =   3750
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "The server you are attached to does not have this service!"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         Left            =   3135
         TabIndex        =   18
         Top             =   3480
         Width           =   1965
      End
      Begin VB.Label Label7 
         Caption         =   "Subscriber Alerts"
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
         Height          =   270
         Left            =   480
         TabIndex        =   16
         Top             =   3180
         Width           =   1320
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   2
         Left            =   210
         Picture         =   "frmAlertSettings.frx":0714
         Top             =   3165
         Width           =   240
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "N.B: You will still recieve IM's with this option unchecked"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   225
         Left            =   1425
         TabIndex        =   15
         Top             =   2640
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "Local Alerts"
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
         Height          =   270
         Left            =   480
         TabIndex        =   10
         Top             =   1515
         Width           =   1185
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   1
         Left            =   210
         Picture         =   "frmAlertSettings.frx":0A9E
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "This includes: New client version update available from server, inactivity warnings and other related alerts."
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   450
         Left            =   810
         TabIndex        =   9
         Top             =   1020
         Width           =   4275
      End
      Begin VB.Label Label3 
         Caption         =   "Network Alerts"
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
         Height          =   270
         Left            =   480
         TabIndex        =   7
         Top             =   405
         Width           =   1185
      End
      Begin VB.Image Image3 
         Height          =   240
         Index           =   0
         Left            =   210
         Picture         =   "frmAlertSettings.frx":0E28
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   315
      Picture         =   "frmAlertSettings.frx":11B2
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
      Caption         =   $"frmAlertSettings.frx":3C62
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
      Picture         =   "frmAlertSettings.frx":3D0A
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alert Settings"
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
      Width           =   2280
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -435
      Picture         =   "frmAlertSettings.frx":49D4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmAlertSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Private Sub cmdApply_Click()
'Save settings with SaveASetting Sub
SaveASetting chkNoNetworkAlerts

SaveASetting chkLogonAlert
SaveASetting chkLogoffAlert
SaveASetting chkiMailAlert
SaveASetting chkIMAlert
SaveASetting chkProfileViewAlert
End Sub

Private Sub cmdCancel_Click()
'Close
Unload Me
End Sub

Private Sub cmdSave_Click()
'Save & close
cmdApply_Click
cmdCancel_Click
End Sub

Private Sub GetASetting(CheckCtl As CheckBox)
'Retireves a setting from the registry (uses the checkbox caption property as the keyname
CheckCtl.Value = GetSetting("Nexis Software Technologies", "ISPN_Server", "Alerts_" & CheckCtl.Caption, CheckCtl.Value)
End Sub

Private Sub SaveASetting(CheckCtl As CheckBox)
'Retireves a setting from the registry (uses the checkbox caption property as the keyname
SaveSetting "Nexis Software Technologies", "ISPN_Server", "Alerts_" & CheckCtl.Caption, CheckCtl.Value
End Sub

Private Sub Form_Load()
'Load values from specific registry key with GetASetting
GetASetting chkNoNetworkAlerts

GetASetting chkLogonAlert
GetASetting chkLogoffAlert
GetASetting chkiMailAlert
GetASetting chkIMAlert
GetASetting chkProfileViewAlert
End Sub
