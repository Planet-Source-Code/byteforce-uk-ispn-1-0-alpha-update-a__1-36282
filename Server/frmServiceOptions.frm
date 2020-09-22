VERSION 5.00
Begin VB.Form frmServiceOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ISPN Services Options"
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
   Icon            =   "frmServiceOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmServiceOptions.frx":038A
   ScaleHeight     =   6555
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Left            =   1575
      TabIndex        =   4
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Height          =   420
      Left            =   2880
      TabIndex        =   5
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4185
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
      Begin VB.CheckBox chkiMailNoRTF 
         Caption         =   "Disable rich text formatting"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   18
         Top             =   3510
         Width           =   2400
      End
      Begin VB.CheckBox chkiMailNoFileAttachments 
         Caption         =   "Disable file attachments"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   17
         Top             =   3255
         Width           =   2115
      End
      Begin VB.CheckBox chkiMailFilter 
         Caption         =   "Server side word filter"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   16
         Top             =   2985
         Width           =   1995
      End
      Begin VB.CheckBox chkServiceiMail 
         Caption         =   "Allow iMail (Private ISPN Mail)"
         Height          =   240
         Left            =   405
         TabIndex        =   15
         Top             =   2685
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CheckBox chkProfilesNoUserStats 
         Caption         =   "Disable user statistic info"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   14
         Top             =   2385
         Width           =   2235
      End
      Begin VB.CheckBox chkProfilesNoPicture 
         Caption         =   "Disable profile picture"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   13
         Top             =   2115
         Width           =   1965
      End
      Begin VB.CheckBox chkServiceProfiles 
         Caption         =   "Allow User Profiles"
         Height          =   240
         Left            =   405
         TabIndex        =   12
         Top             =   1815
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkIMUsePicture 
         Caption         =   "Use profile picture"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   645
         TabIndex        =   11
         Top             =   1485
         Value           =   1  'Checked
         Width           =   1725
      End
      Begin VB.CheckBox chkIMFloodGuard 
         Caption         =   "Protect against flooding"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   10
         Top             =   1215
         Width           =   2130
      End
      Begin VB.CheckBox chkIMFilter 
         Caption         =   "Server side word filter"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   645
         TabIndex        =   9
         Top             =   945
         Width           =   2010
      End
      Begin VB.CheckBox chkServiceIM 
         Caption         =   "Allow Instant Messages"
         Height          =   240
         Left            =   405
         TabIndex        =   8
         Top             =   630
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "These settings do not have any effect yet!"
         ForeColor       =   &H0000FFFF&
         Height          =   570
         Left            =   3135
         TabIndex        =   19
         Top             =   1710
         Width           =   1830
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   105
         Picture         =   "frmServiceOptions.frx":06CC
         Top             =   2655
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   105
         Picture         =   "frmServiceOptions.frx":0A56
         Top             =   1800
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   105
         Picture         =   "frmServiceOptions.frx":0DE0
         Top             =   615
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Service Options"
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
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   315
      Picture         =   "frmServiceOptions.frx":34B0
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
      Caption         =   $"frmServiceOptions.frx":5F60
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
      Picture         =   "frmServiceOptions.frx":6001
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Options: Services"
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
      Width           =   2910
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -435
      Picture         =   "frmServiceOptions.frx":6CCB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmServiceOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkServiceIM_Click()
'Enable\Disable Sub Items
If chkServiceIM.Value = 1 Then
    
    'Enable
    chkIMFilter.Enabled = True
    chkIMFloodGuard.Enabled = True
    chkIMUsePicture.Enabled = True
    
Else
    
    'Disable
    chkIMFilter.Enabled = False
    chkIMFloodGuard.Enabled = False
    chkIMUsePicture.Enabled = False

End If
End Sub

Private Sub chkServiceiMail_Click()
'Enable\Disable Sub Items
If chkServiceiMail.Value = 1 Then

    'Enable
    chkiMailFilter.Enabled = True
    chkiMailNoRTF.Enabled = True
    chkiMailNoFileAttachments.Enabled = True

Else
    
    'Disable
    chkiMailFilter.Enabled = False
    chkiMailNoRTF.Enabled = False
    chkiMailNoFileAttachments.Enabled = False

End If
End Sub

Private Sub chkServiceProfiles_Click()
'Enable\Disable Sub Items
If chkServiceProfiles.Value = 1 Then
    
    'Enable
    chkProfilesNoPicture.Enabled = True
    chkProfilesNoUserStats.Enabled = True
        
Else
    
    'Disable
    chkProfilesNoPicture.Enabled = False
    chkProfilesNoUserStats.Enabled = False
    
End If
End Sub

Private Sub Command1_Click()
'Cancel
Unload Me
End Sub

Private Sub Command2_Click()
'Apply

'Use SaveASetting to save options
SaveASetting chkServiceIM
    SaveASetting chkIMFilter
    SaveASetting chkIMFloodGuard
    SaveASetting chkIMUsePicture

SaveASetting chkServiceiMail
    SaveASetting chkiMailFilter
    SaveASetting chkiMailNoRTF
    SaveASetting chkiMailNoFileAttachments

SaveASetting chkServiceProfiles
    SaveASetting chkProfilesNoPicture
    SaveASetting chkProfilesNoUserStats

End Sub

Private Sub Form_Load()
'Load

'Use GetASetting to load options
GetASetting chkServiceIM
    GetASetting chkIMFilter
    GetASetting chkIMFloodGuard
    GetASetting chkIMUsePicture

GetASetting chkServiceiMail
    GetASetting chkiMailFilter
    GetASetting chkiMailNoRTF
    GetASetting chkiMailNoFileAttachments

GetASetting chkServiceProfiles
    GetASetting chkProfilesNoPicture
    GetASetting chkProfilesNoUserStats

End Sub

Private Sub Command3_Click()
'Save & Close
Command2_Click
Command1_Click
End Sub

Private Sub GetASetting(CheckCtl As CheckBox)
'Retireves a setting from the registry (uses the checkbox caption property as the keyname
CheckCtl.Value = GetSetting("Nexis Software Technologies", "ISPN_Server", "Services_" & CheckCtl.Caption, CheckCtl.Value)
End Sub

Private Sub SaveASetting(CheckCtl As CheckBox)
'Retireves a setting from the registry (uses the checkbox caption property as the keyname
SaveSetting "Nexis Software Technologies", "ISPN_Server", "Services_" & CheckCtl.Caption, CheckCtl.Value
End Sub


