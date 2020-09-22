VERSION 5.00
Begin VB.Form frmLogonStatus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Logging On..."
   ClientHeight    =   900
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   5985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogonStatus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "frmLogonStatus.frx":000C
   ScaleHeight     =   900
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5415
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      MousePointer    =   11  'Hourglass
      Picture         =   "frmLogonStatus.frx":0ADB
      Top             =   165
      Width           =   480
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Finding host: ispn://127.0.0.1:210..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   555
      Left            =   795
      MousePointer    =   11  'Hourglass
      TabIndex        =   0
      Top             =   300
      Width           =   4170
   End
End
Attribute VB_Name = "frmLogonStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The content contained and \ or presented in this product is subject to copyright
'and should not be re-distributed unless the product is not modified in ANY WAY.
'For more information please direct queries to matt@andrews-computers.com

Private Sub lblVersion_Click()
End Sub

Private Sub tmrClose_Timer()
Unload Me
End Sub
