VERSION 5.00
Begin VB.Form frmIMServiceSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Instant Messaging Service Settings"
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
   Icon            =   "frmIMServiceSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmIMServiceSettings.frx":038A
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
      Left            =   1515
      TabIndex        =   4
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Apply"
      Height          =   420
      Left            =   2820
      TabIndex        =   5
      Top             =   6045
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   4125
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
      Begin VB.CheckBox Check2 
         Caption         =   "Play &sound on reciept of message"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   645
         Value           =   1  'Checked
         Width           =   3105
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enable &Word Filtering"
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   3225
      End
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   315
      Picture         =   "frmIMServiceSettings.frx":06CC
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
      Caption         =   $"frmIMServiceSettings.frx":317C
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
      Picture         =   "frmIMServiceSettings.frx":3218
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IM Service Settings"
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
      Width           =   3150
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   -435
      Picture         =   "frmIMServiceSettings.frx":6228
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5985
   End
End
Attribute VB_Name = "frmIMServiceSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

