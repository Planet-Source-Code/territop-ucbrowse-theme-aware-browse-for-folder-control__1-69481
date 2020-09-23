VERSION 5.00
Begin VB.Form frmMultiples 
   Caption         =   "Multiple Hosting"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   8070
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Multiple Hosting:"
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   7335
      Begin Project1.ucBrowse ucBrowse4 
         Height          =   3255
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7435
         Appearance      =   0
         Path            =   "C:\Documents and Settings\c011292\My Documents\Integrative Biology\Science\Programming\VB6\Example Code\ucBrowse\"
      End
      Begin Project1.ucBrowse ucBrowse3 
         Height          =   3255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   7435
         Appearance      =   0
         Path            =   "C:\Documents and Settings\c011292\My Documents\Integrative Biology\Science\Programming\VB6\Example Code\ucBrowse\"
      End
   End
   Begin Project1.ucBrowse ucBrowse2 
      Height          =   3015
      Left            =   4080
      TabIndex        =   1
      Top             =   480
      Width           =   3255
      _ExtentX        =   6376
      _ExtentY        =   6800
      Appearance      =   0
      Path            =   "C:\Documents and Settings\c011292\My Documents\Integrative Biology\Science\Programming\VB6\Example Code\ucBrowse\"
   End
   Begin Project1.ucBrowse ucBrowse1 
      Height          =   3015
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      _ExtentX        =   6376
      _ExtentY        =   6800
      Appearance      =   0
      Path            =   "C:\Documents and Settings\c011292\My Documents\Integrative Biology\Science\Programming\VB6\Example Code\ucBrowse\"
   End
   Begin VB.Label Label1 
      Caption         =   "Symultaneous Object Instances Hosted by Container Objects!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Supports Symultaneous Independant Objects!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   240
      Width           =   6735
   End
End
Attribute VB_Name = "frmMultiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

