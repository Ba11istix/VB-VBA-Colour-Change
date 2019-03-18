VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Colour"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox slider 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   3120
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   41
      Top             =   3720
      Width           =   495
      Begin VB.Shape slide 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   120
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox wheel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   40
      Top             =   3720
      Width           =   2895
      Begin VB.Shape selector 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   90
         Left            =   2520
         Shape           =   3  'Circle
         Top             =   120
         Width           =   90
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
      Height          =   615
      Left            =   3720
      TabIndex        =   38
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   3720
      List            =   "frmMain.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   615
      Left            =   3720
      TabIndex        =   39
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   16
      Left            =   3720
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   15
      Left            =   4320
      Top             =   4680
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   14
      Left            =   3720
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   13
      Left            =   4320
      Top             =   4080
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   12
      Left            =   3720
      Top             =   3480
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   11
      Left            =   4320
      Top             =   3480
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   10
      Left            =   3720
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   9
      Left            =   4320
      Top             =   2880
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   8
      Left            =   3720
      Top             =   2280
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   7
      Left            =   4320
      Top             =   2280
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   6
      Left            =   3720
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   5
      Left            =   4320
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   4
      Left            =   4320
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   3
      Left            =   3720
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   2
      Left            =   3720
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   1
      Left            =   4320
      Top             =   480
      Width           =   495
   End
   Begin VB.Shape border 
      Height          =   495
      Index           =   0
      Left            =   -500
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   -500
      TabIndex        =   25
      Tag             =   "Not Used"
      Top             =   3720
      Width           =   495
   End
   Begin VB.Shape Shape 
      Height          =   3495
      Left            =   120
      Top             =   120
      Width           =   3495
   End
   Begin VB.Line Line 
      BorderColor     =   &H80000000&
      X1              =   120
      X2              =   3600
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label callreturn 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Call Return"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label bookmark 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bookmark"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   28
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   3720
      TabIndex        =   37
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   4320
      TabIndex        =   36
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   3720
      TabIndex        =   35
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   4320
      TabIndex        =   34
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   3720
      TabIndex        =   33
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   4320
      TabIndex        =   32
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   3720
      TabIndex        =   30
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   4320
      TabIndex        =   29
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label norm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " = ""Test"""
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   2
      Left            =   2520
      TabIndex        =   19
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label norm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   1
      Left            =   1920
      TabIndex        =   16
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label norm 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "()"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   960
      Width           =   255
   End
   Begin VB.Label selected 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loop W"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   720
      TabIndex        =   22
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label breakpoint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "        Do Events 'Breakpoint"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label error 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "        I am an error"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   14
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label identifier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   2
      Left            =   2040
      TabIndex        =   17
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label identifier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "txtTest"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   3
      Left            =   1200
      TabIndex        =   18
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label exepoint 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "End Sub 'Running here"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   24
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label keyword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " While True"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   4
      Left            =   1200
      TabIndex        =   23
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label keyword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Do"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   3
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label identifier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " Test"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   1
      Left            =   1440
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.Label keyword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Private Sub"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label keyword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " as Long"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label identifier 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   " i"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   255
   End
   Begin VB.Label keyword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Private"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label comment 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "'This is a comment"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   233
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   3720
      TabIndex        =   27
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   4320
      TabIndex        =   26
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   3720
      TabIndex        =   21
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   4320
      TabIndex        =   20
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   4320
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3720
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label colour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label general 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
   Begin VB.Menu mLoad 
      Caption         =   "Load"
   End
   Begin VB.Menu mSave 
      Caption         =   "Save"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const def01 As String = "ff ff ff 00"
Const def02 As String = "c0 c0 c0 00"
Const def03 As String = "80 80 80 00"
Const def04 As String = "00 00 00 00"
Const def05 As String = "ff 00 00 00"
Const def06 As String = "80 00 00 00"
Const def07 As String = "ff ff 00 00"
Const def08 As String = "80 80 00 00"
Const def09 As String = "00 ff 00 00"
Const def10 As String = "00 80 00 00"
Const def11 As String = "00 ff ff 00"
Const def12 As String = "00 80 80 00"
Const def13 As String = "00 00 ff 00"
Const def14 As String = "00 00 80 00"
Const def15 As String = "ff 00 ff 00"
Const def16 As String = "80 00 80 00"

Const VB6 As Long = 926453      'VB6.exe, reversed style
Const VBA6 As Long = 727109     'VBA6.DLL, Normal style
Const VBA6Rev As Long = 1596905 'VBA6.DLL, Reversed style
Const VBE7 As Long = 2305245    'VBE7.DLL, Normal style
Const VBE7Rev As Long = 2301309 'VBE7.DLL, Reversed style
Const VBE6 As Long = 902573     'VBE6.DLL, Normal style
Const VBE6Rev As Long = 1889505 'VBE6.dll, Reversed style

Const ByteCount = 64

Private Type OpenFileName
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileNameA Lib "comdlg32.dll" (pOpenfilename As OpenFileName) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hFile As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Const PI As Double = 3.14159265358979

Private Current(1 To 2, 1 To 10) As Integer '(1,x) = Back colour, (2,x) = Forecolour
Private Paths(0 To 3) As String '0=VB6, 1=VBA6, 2=VBE7, 3=VBE6

Private CurrentColour As Integer

Function BrowseForFile(sInitDir As String, Optional ByVal sFileFilters As String, Optional sTitle As String = "Open File") As String
    Dim tFileBrowse As OpenFileName
    Const clMaxLen As Long = 254
    
    tFileBrowse.lStructSize = Len(tFileBrowse)
    
    sFileFilters = Replace(sFileFilters, "|", vbNullChar)
    sFileFilters = Replace(sFileFilters, ";", vbNullChar)
    If Right$(sFileFilters, 1) <> vbNullChar Then
        sFileFilters = sFileFilters & vbNullChar
    End If
    
    tFileBrowse.flags = 0
    tFileBrowse.hwndOwner = Me.hWnd
    tFileBrowse.hInstance = App.hInstance
    tFileBrowse.lpstrFilter = sFileFilters & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
    tFileBrowse.lpstrFile = String(clMaxLen, " ")
    tFileBrowse.nMaxFile = clMaxLen + 1
    tFileBrowse.lpstrFileTitle = Space$(clMaxLen)
    tFileBrowse.nMaxFileTitle = clMaxLen + 1
    tFileBrowse.lpstrInitialDir = sInitDir
    tFileBrowse.lpstrTitle = sTitle

    If GetOpenFileName(tFileBrowse) Then
        BrowseForFile = Trim$(tFileBrowse.lpstrFile)
        If Right$(BrowseForFile, 1) = vbNullChar Then
            BrowseForFile = left$(BrowseForFile, Len(BrowseForFile) - 1)
        End If
    End If
End Function

Function SaveAsCommonDialog(Optional sTitle = "Save File", Optional sFilter As String, Optional sDefaultDir As String) As String
    Const clBufferLen As Long = 255
    Dim OFName As OpenFileName, sBuffer As String * clBufferLen
    
    OFName.lStructSize = Len(OFName)
    OFName.hwndOwner = Me.hWnd
    OFName.hInstance = App.hInstance
    If Len(sFilter) Then
        OFName.lpstrFilter = sFilter
    Else
        OFName.lpstrFilter = "Text Files (*.txt)" & Chr$(0) & "*.txt" & Chr$(0) & "All Files (*.*)" & Chr$(0) & "*.*" & Chr$(0)
    End If
    OFName.lpstrFile = sBuffer
    OFName.nMaxFile = clBufferLen
    OFName.lpstrFileTitle = sBuffer
    OFName.nMaxFileTitle = clBufferLen
    If Len(sDefaultDir) Then
        OFName.lpstrInitialDir = sDefaultDir
    Else
        OFName.lpstrInitialDir = CurDir$
    End If
    OFName.lpstrTitle = sTitle
    OFName.flags = 0

    If GetSaveFileNameA(OFName) Then
        SaveAsCommonDialog = left$(OFName.lpstrFile, InStr(1, OFName.lpstrFile, Chr(0)) - 1)
    Else
        SaveAsCommonDialog = ""
    End If
End Function

Private Function GetKey(path As String, key As String) As String
    Dim hKey As Long
    Dim sValue As String
    Dim lValue As Long
    Dim ret As Long
    
    RegOpenKey &H80000001, path, hKey
    ret = RegQueryValueEx(hKey, key, 0&, 1&, vbNullString, lValue)
    If ret = 2 Then 'didnt find key
        If key = "CodeBackColors" Then
            GetKey = "1 14 1 7 6 1 1 1 1 1 0 0 0 0 0 0 " 'Default back colours (highlight colour does not match)
        ElseIf key = "CodeForeColors" Then
            GetKey = "4 1 5 4 1 10 14 4 4 4 0 0 0 0 0 0 " 'Default fore colours
        ElseIf key = "FontFace" Then
            GetKey = "Consolas" 'font we want...
        End If
    Else
        sValue = Space(lValue)
        RegQueryValueEx hKey, key, 0&, 1&, ByVal sValue, Len(sValue)
        GetKey = left(sValue, Len(sValue) - 1)
    End If
    RegCloseKey hKey
End Function

Private Sub SetKey(path As String, key As String, value As String)
    Dim hKey As Long

    RegOpenKey &H80000001, path, hKey
    RegSetValueExString hKey, key, 0&, 1, value, Len(value)
    RegCloseKey hKey
End Sub

Private Function GetBackColours() As String
    If cmbType.text = "VB6" Then
        GetBackColours = GetKey("Software\Microsoft\VBA\Microsoft Visual Basic", "CodeBackColors")
    ElseIf cmbType.text = "VBE7" Then
        GetBackColours = GetKey("Software\Microsoft\VBA\7.1\Common", "CodeBackColors")
    ElseIf cmbType.text = "VBE6" Then
        GetBackColours = GetKey("Software\Microsoft\VBA\6.0\Common", "CodeBackColors")
    End If
End Function

Private Function GetForeColours() As String
    If cmbType.text = "VB6" Then
        GetForeColours = GetKey("Software\Microsoft\VBA\Microsoft Visual Basic", "CodeForeColors")
    ElseIf cmbType.text = "VBE7" Then
        GetForeColours = GetKey("Software\Microsoft\VBA\7.1\Common", "CodeForeColors")
    ElseIf cmbType.text = "VBE6" Then
        GetForeColours = GetKey("Software\Microsoft\VBA\6.0\Common", "CodeForeColors")
    End If
End Function

Private Sub ReadCurrent()
    Dim startByte As Long
    Dim f As Long
    Dim i As Long
    
    Dim B As Byte
    Dim strText As String
    Dim colours() As String
    
    f = FreeFile
    
    If cmbType.text = "VB6" Then
        startByte = VB6
        Open Paths(0) For Binary As #f
    ElseIf cmbType.text = "VBE7" Then
        startByte = VBE7Rev
        Open Paths(2) For Binary As #f
    ElseIf cmbType.text = "VBE6" Then
        startByte = VBE6Rev
        Open Paths(3) For Binary As #f
    Else
        Exit Sub
    End If
    
    For i = startByte To startByte + ByteCount - 1
        Get #f, i, B
        strText = strText & Right("0" & Hex$(B), 2) & " "
    Next i
    strText = Trim$(strText)
    
    colour(1).BackColor = ColourFromByte(Mid(strText, 15 * 12 + 1, 11))
    colour(2).BackColor = ColourFromByte(Mid(strText, 7 * 12 + 1, 11))
    colour(3).BackColor = ColourFromByte(Mid(strText, 8 * 12 + 1, 11))
    colour(4).BackColor = ColourFromByte(Mid(strText, 0 * 12 + 1, 11))
    colour(5).BackColor = ColourFromByte(Mid(strText, 12 * 12 + 1, 11))
    colour(6).BackColor = ColourFromByte(Mid(strText, 4 * 12 + 1, 11))
    colour(7).BackColor = ColourFromByte(Mid(strText, 14 * 12 + 1, 11))
    colour(8).BackColor = ColourFromByte(Mid(strText, 6 * 12 + 1, 11))
    colour(9).BackColor = ColourFromByte(Mid(strText, 10 * 12 + 1, 11))
    colour(10).BackColor = ColourFromByte(Mid(strText, 2 * 12 + 1, 11))
    colour(11).BackColor = ColourFromByte(Mid(strText, 11 * 12 + 1, 11))
    colour(12).BackColor = ColourFromByte(Mid(strText, 3 * 12 + 1, 11))
    colour(13).BackColor = ColourFromByte(Mid(strText, 9 * 12 + 1, 11))
    colour(14).BackColor = ColourFromByte(Mid(strText, 1 * 12 + 1, 11))
    colour(15).BackColor = ColourFromByte(Mid(strText, 13 * 12 + 1, 11))
    colour(16).BackColor = ColourFromByte(Mid(strText, 5 * 12 + 1, 11))
    
    Close #f
    
    colours = Split(GetBackColours, " ")
    For i = 1 To 10
        Current(1, i) = CInt(colours(i - 1))
        If Current(1, i) = 0 Then Current(1, i) = 1
    Next i
    
    colours = Split(GetForeColours, " ")
    For i = 1 To 10
        Current(2, i) = CInt(colours(i - 1))
        If Current(2, i) = 0 Then Current(2, i) = 1
    Next i
End Sub

Private Function ColourFromByte(sIn As String) As Long
    ColourFromByte = CLng("&H" & Mid(sIn, 7, 2) & Mid(sIn, 4, 2) & Mid(sIn, 1, 2))
End Function

Private Function ColourNormal() As String
    ColourNormal = Trim(ColorToHex(colour(1).BackColor) & ColorToHex(colour(2).BackColor) _
                 & ColorToHex(colour(3).BackColor) & ColorToHex(colour(4).BackColor) _
                 & ColorToHex(colour(5).BackColor) & ColorToHex(colour(6).BackColor) _
                 & ColorToHex(colour(7).BackColor) & ColorToHex(colour(8).BackColor) _
                 & ColorToHex(colour(9).BackColor) & ColorToHex(colour(10).BackColor) _
                 & ColorToHex(colour(11).BackColor) & ColorToHex(colour(12).BackColor) _
                 & ColorToHex(colour(13).BackColor) & ColorToHex(colour(14).BackColor) _
                 & ColorToHex(colour(15).BackColor) & ColorToHex(colour(16).BackColor))
End Function

Private Function ColourReversed() As String
    ColourReversed = Trim(ColorToHex(colour(4).BackColor) & ColorToHex(colour(14).BackColor) _
                        & ColorToHex(colour(10).BackColor) & ColorToHex(colour(12).BackColor) _
                        & ColorToHex(colour(6).BackColor) & ColorToHex(colour(16).BackColor) _
                        & ColorToHex(colour(8).BackColor) & ColorToHex(colour(2).BackColor) _
                        & ColorToHex(colour(3).BackColor) & ColorToHex(colour(13).BackColor) _
                        & ColorToHex(colour(9).BackColor) & ColorToHex(colour(11).BackColor) _
                        & ColorToHex(colour(5).BackColor) & ColorToHex(colour(15).BackColor) _
                        & ColorToHex(colour(7).BackColor) & ColorToHex(colour(1).BackColor))
End Function

Private Function ColorToHex(ByVal Color As Long) As String 'returns 4 byte string with spaces
    Dim bytOut(11) As Byte

    bytOut(0) = &H30& Or ((Color And &HF0&) \ &H10&)
    bytOut(2) = &H30& Or (Color And &HF&)
    bytOut(4) = &H30& Or ((Color And &HF000&) \ &H1000&)
    bytOut(6) = &H30& Or ((Color And &HF00&) \ &H100&)
    bytOut(8) = &H30& Or ((Color And &HF00000) \ &H100000)
    bytOut(10) = &H30& Or ((Color And &HF0000) \ &H10000)

    If bytOut(0) > &H39 Then bytOut(0) = bytOut(0) + 7
    If bytOut(2) > &H39 Then bytOut(2) = bytOut(2) + 7
    If bytOut(4) > &H39 Then bytOut(4) = bytOut(4) + 7
    If bytOut(6) > &H39 Then bytOut(6) = bytOut(6) + 7
    If bytOut(8) > &H39 Then bytOut(8) = bytOut(8) + 7
    If bytOut(10) > &H39 Then bytOut(10) = bytOut(10) + 7

    ColorToHex = CStr(bytOut)
    ColorToHex = left(ColorToHex, 2) & " " & Mid(ColorToHex, 3, 2) & " " & Mid(ColorToHex, 5, 2) & " 00 "
End Function

Private Sub cmbType_Click()
    ReadCurrent
    If Current(1, 1) <> 0 Then Update
End Sub

Private Sub cmdWrite_Click()
    Dim hexBytes() As String
    If cmbType.text = "VB6" Then
        If FileInUse(Paths(0)) = False And FileInUse(Paths(1)) = False Then
            hexBytes = Split(ColourReversed, " ")
            WriteFile VB6, hexBytes, Paths(0)
            WriteFile VBA6Rev, hexBytes, Paths(1)
            hexBytes = Split(ColourNormal, " ")
            WriteFile VBA6, hexBytes, Paths(1)
            UpdateReg
        Else
            MsgBox "Please close VB6!", vbExclamation, "Error"
        End If
    ElseIf cmbType.text = "VBE7" Then
        If FileInUse(Paths(2)) = False Then
            hexBytes = Split(ColourReversed, " ")
            WriteFile VBE7Rev, hexBytes, Paths(2)
            hexBytes = Split(ColourNormal, " ")
            WriteFile VBE7, hexBytes, Paths(2)
            UpdateReg
        Else
            MsgBox "Please close all open Office applications!", vbExclamation, "Error"
        End If
    ElseIf cmbType.text = "VBE6" Then
        If FileInUse(Paths(3)) = False Then
            hexBytes = Split(ColourReversed, " ")
            WriteFile VBE6Rev, hexBytes, Paths(3)
            hexBytes = Split(ColourNormal, " ")
            WriteFile VBE6, hexBytes, Paths(3)
            UpdateReg
        Else
            MsgBox "Please close all open Office applications!", vbExclamation, "Error"
        End If
    End If
End Sub

Private Sub WriteFile(startByte As Long, hexBytes() As String, file As String)
    Dim i As Long
    Dim f As Long
    Dim B As Byte
    
    f = FreeFile
    Open file For Binary As #f
    For i = 0 To ByteCount - 1
        B = "&h" & hexBytes(i)
        Put #f, startByte + i, B
    Next i
    Close #f
End Sub

Private Function FileInUse(sFile As String) As Boolean
    Dim hFile As Long
    
    hFile = CreateFile(sFile, &H80000000, 0, 0, 3, &H80, 0&)
    FileInUse = hFile = -1&
    CloseHandle hFile
End Function

Private Sub UpdateReg()
    Dim path As String
    If cmbType.text = "VB6" Then
        path = "Software\Microsoft\VBA\Microsoft Visual Basic"
    ElseIf cmbType.text = "VBE7" Then
        path = "Software\Microsoft\VBA\7.1\Common"
    ElseIf cmbType.text = "VBE6" Then
        path = "Software\Microsoft\VBA\6.0\Common"
    End If
    
    SetKey path, "CodeBackColors", GetBackColourReg
    SetKey path, "CodeForeColors", GetForeColourReg
    SetKey path, "FontFace", "Consolas"
End Sub

Private Function GetBackColourReg() As String
    GetBackColourReg = Current(1, 1) & " " & Current(1, 2) & " " & Current(1, 3) & " " & Current(1, 4) & " " & _
                       Current(1, 5) & " " & Current(1, 6) & " " & Current(1, 7) & " " & Current(1, 8) & " " & _
                       Current(1, 9) & " " & Current(1, 10) & " "
End Function

Private Function GetForeColourReg() As String
    GetForeColourReg = Current(2, 1) & " " & Current(2, 2) & " " & Current(2, 3) & " " & Current(2, 4) & " " & _
                       Current(2, 5) & " " & Current(2, 6) & " " & Current(2, 7) & " " & Current(2, 8) & " " & _
                       Current(2, 9) & " " & Current(2, 10) & " "
End Function

Private Sub cmdDefault_Click()
    colour(1).BackColor = ColourFromByte(def01)
    colour(2).BackColor = ColourFromByte(def02)
    colour(3).BackColor = ColourFromByte(def03)
    colour(4).BackColor = ColourFromByte(def04)
    colour(5).BackColor = ColourFromByte(def05)
    colour(6).BackColor = ColourFromByte(def06)
    colour(7).BackColor = ColourFromByte(def07)
    colour(8).BackColor = ColourFromByte(def08)
    colour(9).BackColor = ColourFromByte(def09)
    colour(10).BackColor = ColourFromByte(def10)
    colour(11).BackColor = ColourFromByte(def11)
    colour(12).BackColor = ColourFromByte(def12)
    colour(13).BackColor = ColourFromByte(def13)
    colour(14).BackColor = ColourFromByte(def14)
    colour(15).BackColor = ColourFromByte(def15)
    colour(16).BackColor = ColourFromByte(def16)
    
    Current(1, 1) = 1
    Current(1, 2) = 14
    Current(1, 3) = 1
    Current(1, 4) = 7
    Current(1, 5) = 6
    Current(1, 6) = 1
    Current(1, 7) = 1
    Current(1, 8) = 1
    Current(1, 9) = 1
    Current(1, 10) = 1
    
    Current(2, 1) = 4
    Current(2, 2) = 1
    Current(2, 3) = 5
    Current(2, 4) = 4
    Current(2, 5) = 1
    Current(2, 6) = 10
    Current(2, 7) = 14
    Current(2, 8) = 4
    Current(2, 9) = 4
    Current(2, 10) = 4
    
    Update
End Sub

Private Sub Update()
    general.BackColor = colour(Current(1, 1)).BackColor
    norm(0).BackColor = colour(Current(1, 1)).BackColor
    norm(0).ForeColor = colour(Current(2, 1)).BackColor
    norm(1).BackColor = colour(Current(1, 1)).BackColor
    norm(1).ForeColor = colour(Current(2, 1)).BackColor
    norm(2).BackColor = colour(Current(1, 1)).BackColor
    norm(2).ForeColor = colour(Current(2, 1)).BackColor
    
    selected.BackColor = colour(Current(1, 2)).BackColor
    selected.ForeColor = colour(Current(2, 2)).BackColor

    error.BackColor = colour(Current(1, 3)).BackColor
    error.ForeColor = colour(Current(2, 3)).BackColor
    
    exepoint.BackColor = colour(Current(1, 4)).BackColor
    exepoint.ForeColor = colour(Current(2, 4)).BackColor

    breakpoint.BackColor = colour(Current(1, 5)).BackColor
    breakpoint.ForeColor = colour(Current(2, 5)).BackColor
    
    comment.BackColor = colour(Current(1, 6)).BackColor
    comment.ForeColor = colour(Current(2, 6)).BackColor
    
    keyword(0).BackColor = colour(Current(1, 7)).BackColor
    keyword(0).ForeColor = colour(Current(2, 7)).BackColor
    keyword(1).BackColor = colour(Current(1, 7)).BackColor
    keyword(1).ForeColor = colour(Current(2, 7)).BackColor
    keyword(2).BackColor = colour(Current(1, 7)).BackColor
    keyword(2).ForeColor = colour(Current(2, 7)).BackColor
    keyword(3).BackColor = colour(Current(1, 7)).BackColor
    keyword(3).ForeColor = colour(Current(2, 7)).BackColor
    keyword(4).BackColor = colour(Current(1, 7)).BackColor
    keyword(4).ForeColor = colour(Current(2, 7)).BackColor
    
    identifier(0).BackColor = colour(Current(1, 8)).BackColor
    identifier(0).ForeColor = colour(Current(2, 8)).BackColor
    identifier(1).BackColor = colour(Current(1, 8)).BackColor
    identifier(1).ForeColor = colour(Current(2, 8)).BackColor
    identifier(2).BackColor = colour(Current(1, 8)).BackColor
    identifier(2).ForeColor = colour(Current(2, 8)).BackColor
    identifier(3).BackColor = colour(Current(1, 8)).BackColor
    identifier(3).ForeColor = colour(Current(2, 8)).BackColor

    bookmark.BackColor = colour(Current(1, 9)).BackColor
    bookmark.ForeColor = colour(Current(2, 9)).BackColor
    
    callreturn.BackColor = colour(Current(1, 10)).BackColor
    callreturn.ForeColor = colour(Current(2, 10)).BackColor
End Sub

Private Sub colour_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    border(CurrentColour).BorderColor = 0
    border(CurrentColour).BorderWidth = 1
    border(Index).BorderColor = 255
    border(Index).BorderWidth = 2
    CurrentColour = Index
    MoveColour colour(Index).BackColor
End Sub

Private Sub Form_Load()
    If Dir("C:\Program Files\VB6\VB6.exe") <> vbNullString Then
        If Dir("C:\Program Files\VB6\VBA6.dll") <> vbNullString Then
            Paths(0) = "C:\Program Files\VB6\VB6.exe"
            Paths(1) = "C:\Program Files\VB6\VBA6.dll"
        End If
    ElseIf Dir("C:\Program Files (x86)\VB6\VB6.exe") <> vbNullString Then
        If Dir("C:\Program Files (x86)\VB6\VBA6.dll") <> vbNullString Then
            Paths(0) = "C:\Program Files (x86)\VB6\VB6.exe"
            Paths(1) = "C:\Program Files (x86)\VB6\VBA6.dll"
        End If
    ElseIf Dir("C:\Program Files\VB98\VB6.exe") <> vbNullString Then
        If Dir("C:\Program Files\VB98\VBA6.dll") <> vbNullString Then
            Paths(0) = "C:\Program Files\VB98\VB6.exe"
            Paths(1) = "C:\Program Files\VB98\VBA6.dll"
        End If
    ElseIf Dir("C:\Program Files (x86)\VB98\VB6.exe") <> vbNullString Then
        If Dir("C:\Program Files (x86)\VB98\VBA6.dll") <> vbNullString Then
            Paths(0) = "C:\Program Files (x86)\VB98\VB6.exe"
            Paths(1) = "C:\Program Files (x86)\VB98\VBA6.dll"
        End If
    ElseIf Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe") <> vbNullString Then
        If Dir("C:\Program Files (x86)\Microsoft Visual Studio\VB98\VBA6.dll") <> vbNullString Then
            Paths(0) = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe"
            Paths(1) = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VBA6.dll"
        End If
    End If
    If Paths(0) <> vbNullString Then cmbType.AddItem "VB6"
    
    If Dir("C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7.1\VBE7.dll") <> vbNullString Then
        Paths(2) = "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA7.1\VBE7.dll"
    ElseIf Dir("C:\Program Files\Common Files\microsoft shared\VBA\VBA7.1\VBE7.dll") <> vbNullString Then
        Paths(2) = "C:\Program Files\Common Files\microsoft shared\VBA\VBA7.1\VBE7.dll"
    End If
    If Paths(2) <> vbNullString Then cmbType.AddItem "VBE7"
    
    If Dir("C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6.dll") <> vbNullString Then
        Paths(3) = "C:\Program Files (x86)\Common Files\microsoft shared\VBA\VBA6\VBE6.dll"
    ElseIf Dir("C:\Program Files\Common Files\microsoft shared\VBA\VBA6\VBE6.dll") <> vbNullString Then
        Paths(3) = "C:\Program Files\Common Files\microsoft shared\VBA\VBA6\VBE6.dll"
    End If
    If Paths(3) <> vbNullString Then cmbType.AddItem "VBE6"
    
    UpdateWheel
    UpdateSlider
    MoveSelector wheel.ScaleWidth / 2, wheel.ScaleHeight / 2
    MoveSlider 1000
    
    cmdDefault_Click
    'colour_MouseUp 2, 1, 1, 1, 1
End Sub

Private Sub mLoad_Click()
    Dim file As String
    Dim f As Long
    Dim i As Integer
    Dim text(1 To 18) As String
    Dim back() As String
    Dim fore() As String
    
    file = BrowseForFile(App.path, "Style File (*.ini);*.ini", "Open Style")
    If file <> vbNullString Then
        f = FreeFile
        Open file For Input As #f
            Do While Not EOF(f) And i < UBound(text)
                Line Input #f, text(i + 1)
                i = i + 1
            Loop
        Close #f
        On Error GoTo handler
        For i = 1 To 16
            colour(i).BackColor = ColourFromByte(text(i))
        Next i
        back = Split(text(17), " ")
        fore = Split(text(18), " ")
        For i = 1 To 10
            Current(1, i) = CInt(back(i - 1))
            Current(2, i) = CInt(fore(i - 1))
        Next i
        Update
    End If
Exit Sub
handler:
    MsgBox "There was an error loading your file!", vbExclamation, "Error"
    Err.Clear
End Sub

Private Sub mSave_Click()
    Dim file As String
    Dim f As Long
    Dim i As Integer
    file = SaveAsCommonDialog(, "Style File (*.ini);*.ini", App.path) & ".ini"
    If file <> vbNullString Then
        f = FreeFile
        Open file For Output As #f
            On Error GoTo handler
            For i = 1 To 16
                Print #f, ColorToHex(colour(i).BackColor)
            Next i
            Print #f, Join2D(Current, 1)
            Print #f, Join2D(Current, 2)
            On Error GoTo 0
        Close #f
    End If
Exit Sub
handler:
    MsgBox "There was an error saving your file!", vbExclamation, "Error"
    Close #f
    Err.Clear
End Sub

Private Function Join2D(vArray As Variant, Optional bound As Integer = 1, Optional ByVal delim As String = " ") As String
    Dim i As Integer
    For i = LBound(vArray, 2) To UBound(vArray, 2)
        Join2D = Join2D & vArray(bound, i) & delim
    Next i
End Function

'Chenge colour of the sections
Private Sub norm_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single):       Current(Abs(Button - 3), 1) = CurrentColour: Update: End Sub
Private Sub general_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                      Current(Abs(Button - 3), 1) = CurrentColour: Update: End Sub
Private Sub selected_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                     Current(Abs(Button - 3), 2) = CurrentColour: Update: End Sub
Private Sub error_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                        Current(Abs(Button - 3), 3) = CurrentColour: Update: End Sub
Private Sub exepoint_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                     Current(Abs(Button - 3), 4) = CurrentColour: Update: End Sub
Private Sub breakpoint_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                   Current(Abs(Button - 3), 5) = CurrentColour: Update: End Sub
Private Sub comment_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                      Current(Abs(Button - 3), 6) = CurrentColour: Update: End Sub
Private Sub keyword_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single):    Current(Abs(Button - 3), 7) = CurrentColour: Update: End Sub
Private Sub identifier_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): Current(Abs(Button - 3), 8) = CurrentColour: Update: End Sub
Private Sub bookmark_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                     Current(Abs(Button - 3), 9) = CurrentColour: Update: End Sub
Private Sub callreturn_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single):                   Current(Abs(Button - 3), 10) = CurrentColour: Update: End Sub

Private Sub UpdateWheel()
    Const scaler As Double = 1.8
    Dim cx As Single, cy As Single
    Dim x As Single, y As Single
    Dim dist As Single
    Dim rad As Single
    
    If wheel.ScaleWidth > wheel.ScaleHeight Then
        rad = wheel.ScaleHeight / 2
    Else
        rad = wheel.ScaleWidth / 2
    End If
    
    For y = 0 To wheel.ScaleHeight
        cy = rad - y
        For x = 0 To wheel.ScaleWidth
            cx = rad - x
            dist = Sqr(cx * cx + cy * cy)
            If dist < rad Then
                dist = (Exp(scaler * (1 - ((rad - dist - 2) / rad))) - 1) / (Exp(scaler) - 1)
                If dist > 1 Then dist = 1
                SetPixel wheel.hdc, x, y, HSVtoRGB(255 * (ArcTan2(cx, cy) + PI) / (2 * PI), 255 * dist, 255)
            End If
        Next x
    Next y
    wheel.Refresh
End Sub

Private Sub MoveSelector(left As Single, top As Single)
    Dim vx As Single, vy As Single
    Dim V As Single
    
    Dim rad As Single

    rad = wheel.ScaleWidth / 2
    vx = left - rad
    vy = top - rad
    V = Sqr(vx * vx + vy * vy)
    If V > rad Then
        left = rad + vx / V * rad
        top = rad + vy / V * rad
    End If
    
    left = left - selector.Width / 2
    top = top - selector.Height / 2
    
    If left < rad Then
        left = left + 1
    Else
        left = left - 1
    End If
    
    If top < rad Then
        top = top + 1
    Else
        top = top - 1
    End If

    selector.Move left, top
    selector.BackColor = GetPixel(wheel.hdc, selector.left + selector.Width / 2, selector.top + selector.Height / 2)
    
    UpdateSlider
End Sub

Private Sub UpdateSlider()
    Dim x As Single, y As Single
    Dim H As Byte, S As Byte, V As Byte
    Dim lColour As Long
    Dim dist As Integer
    
    RGBtoHSV selector.BackColor, H, S, V
    
    dist = slider.ScaleHeight - slide.Height * 2
    
    For y = slide.Height To slider.ScaleHeight - slide.Height
        lColour = HSVtoRGB(H, S, 255 * (y - slide.Height) / (slider.ScaleHeight - slide.Height * 2))
        For x = 0 To slider.ScaleWidth
            SetPixel slider.hdc, x, y, lColour
        Next x
    Next y
    For y = slider.ScaleHeight - slide.Height To slider.ScaleHeight
        For x = 0 To slider.ScaleWidth
            SetPixel slider.hdc, x, y, lColour
        Next x
    Next y
    
    slider.Refresh
    slide.FillColor = GetPixel(slider.hdc, 15, slide.top + slide.Height / 2)
    colour(CurrentColour).BackColor = slide.FillColor
End Sub

Private Sub MoveSlider(top As Single)
    If top < 1 Then
        top = 1
    ElseIf top > slider.ScaleHeight - slide.Height Then
        top = slider.ScaleHeight - slide.Height
    End If
    slide.Move 8, top
    slide.FillColor = GetPixel(slider.hdc, 15, top + slide.Height / 2)
    colour(CurrentColour).BackColor = slide.FillColor
End Sub

Private Function ArcTan2(x As Single, y As Single) As Single
    Select Case x
    Case Is > 0
        ArcTan2 = Atn(y / x)
    Case Is < 0
        ArcTan2 = Atn(y / x) + PI * Sgn(y)
        If y = 0 Then ArcTan2 = ArcTan2 + PI
    Case Is = 0
        ArcTan2 = PI / 2 * Sgn(y)
    End Select
End Function

Private Sub MoveColour(colour As Long)
    Dim H As Byte, S As Byte, V As Byte
    Dim wH As Byte, wS As Byte, wV As Byte
    Dim x As Single, y As Single, z As Single
    RGBtoHSV colour, H, S, V
    
    For y = 0 To wheel.ScaleHeight - 1
        For x = 0 To wheel.ScaleWidth - 1
            RGBtoHSV GetPixel(wheel.hdc, x, y), wH, wS, wV
            If (wH = H And wS = S) Or S = 0 Then
                If S = 0 Then
                    x = wheel.ScaleWidth / 2
                    y = wheel.ScaleHeight / 2
                End If
                MoveSelector x, y
                For z = 0 To slider.ScaleHeight
                    RGBtoHSV GetPixel(slider.hdc, 8, z), wH, wS, wV
                    If wV = V Then
                        MoveSlider z
                        Exit Sub
                    End If
                Next z
            End If
        Next x
    Next y
End Sub

Private Sub slider_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveSlider y
End Sub

Private Sub slider_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveSlider y
End Sub

Private Sub wheel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveSelector x, y
End Sub

Private Sub wheel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then MoveSelector x, y
End Sub

Private Sub RGBtoHSV(C As Long, ByRef H As Byte, ByRef S As Byte, ByRef V As Byte)
    Dim MinVal As Byte
    Dim MaxVal As Byte
    Dim Chroma As Byte
    Dim TempH As Single
    Dim R As Byte, G As Byte, B As Byte
    
    R = C Mod 256
    G = (C \ 256) Mod 256
    B = (C \ 256 \ 256) Mod 256
    
    If R > G Then MaxVal = R Else MaxVal = G
    If B > MaxVal Then MaxVal = B
    If R < G Then MinVal = R Else MinVal = G
    If B < MinVal Then MinVal = B
    Chroma = MaxVal - MinVal
    
    V = MaxVal
    If MaxVal = 0 Then S = 0 Else S = Chroma / MaxVal * 255
    If Chroma = 0 Then
        H = 0
    Else
        Select Case MaxVal
        Case R
            TempH = (1& * G - B) / Chroma
            If TempH < 0 Then TempH = TempH + 6
            H = TempH / 6 * 255
        Case G
            H = (((1& * B - R) / Chroma) + 2) / 6 * 255
        Case B
            H = (((1& * R - G) / Chroma) + 4) / 6 * 255
        End Select
    End If
End Sub

Private Function HSVtoRGB(ByVal H As Byte, ByVal S As Byte, ByVal V As Byte) As Long
    Dim R As Byte, G As Byte, B As Byte
    Dim MinVal As Byte
    Dim MaxVal As Byte
    Dim Chroma As Byte
    Dim TempH As Single
    
    If V = 0 Then
        R = 0
        G = 0
        B = 0
    Else
        If S = 0 Then
            R = V
            G = V
            B = V
        Else
            MaxVal = V
            Chroma = S / 255 * MaxVal
            MinVal = MaxVal - Chroma
            Select Case H
            Case Is >= 170
                TempH = (H - 170) / 43
                If TempH < 1 Then
                    B = MaxVal
                    R = MaxVal * TempH
                Else
                    R = MaxVal
                    B = MaxVal * (2 - TempH)
                End If
                G = 0
            Case Is >= 85
                TempH = (H - 85) / 43
                If TempH < 1 Then
                    G = MaxVal
                    B = MaxVal * TempH
                Else
                    B = MaxVal
                    G = MaxVal * (2 - TempH)
                End If
                R = 0
            Case Else
                TempH = H / 43
                If TempH < 1 Then
                    R = MaxVal
                    G = MaxVal * TempH
                Else
                    G = MaxVal
                    R = MaxVal * (2 - TempH)
                End If
                B = 0
            End Select
            R = R / MaxVal * (MaxVal - MinVal) + MinVal
            G = G / MaxVal * (MaxVal - MinVal) + MinVal
            B = B / MaxVal * (MaxVal - MinVal) + MinVal
            HSVtoRGB = RGB(CInt(R), CInt(G), CInt(B))
        End If
    End If
End Function
