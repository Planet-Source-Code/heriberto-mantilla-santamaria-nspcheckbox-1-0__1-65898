VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo nspCheckBox"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4350
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   0
      Left            =   405
      TabIndex        =   0
      Top             =   825
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   397
      BackColor       =   13126400
      BorderColor     =   33023
      Caption         =   "  Checkbox 1"
      FocusColor      =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   1
      Left            =   405
      TabIndex        =   1
      Top             =   1440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   397
      BackColor       =   13126400
      BorderColor     =   33023
      Caption         =   "  Checkbox 2"
      FocusColor      =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   2
      Left            =   2460
      TabIndex        =   2
      Top             =   1095
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   397
      BackColor       =   16762279
      BorderColor     =   13804169
      Caption         =   "Hardware"
      FocusColor      =   49344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   3
      Left            =   2460
      TabIndex        =   3
      Top             =   1575
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   397
      BackColor       =   16762279
      BorderColor     =   13804169
      Caption         =   "Software"
      FocusColor      =   12583104
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   4
      Left            =   2460
      TabIndex        =   4
      Top             =   2055
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   397
      BackColor       =   16762279
      BorderColor     =   13804169
      Caption         =   "VB Controls"
      FocusColor      =   8438015
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin Checkbox.nspCheckBox nspCheckBox1 
      Height          =   225
      Index           =   5
      Left            =   2460
      TabIndex        =   5
      Top             =   2550
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   397
      BackColor       =   16762279
      BorderColor     =   13804169
      Caption         =   "Contest Winner"
      FocusColor      =   13804169
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by Steppenwolfe and  edit by Heriberto Mantilla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   90
      TabIndex        =   6
      Top             =   2550
      Width           =   2130
   End
   Begin VB.Image imgBorder 
      Height          =   3510
      Index           =   1
      Left            =   2265
      Picture         =   "frmTest.frx":058A
      Top             =   75
      Width           =   2040
   End
   Begin VB.Image imgBorder 
      Height          =   2295
      Index           =   0
      Left            =   60
      Picture         =   "frmTest.frx":1A30
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private DummyToKeepDecCommentsInDeclarations As Boolean


