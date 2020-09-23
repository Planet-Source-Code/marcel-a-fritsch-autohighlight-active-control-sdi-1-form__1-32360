VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   " Highlighting aktiv Controls"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   1830
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   2205
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   1455
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   600
      TabIndex        =   9
      Top             =   2760
      Width           =   4575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Example: Automatic highlighting activ controls using the hook function."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   788
      TabIndex        =   8
      Top             =   240
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Test4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   600
      TabIndex        =   7
      Top             =   2245
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Test3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   6
      Top             =   1871
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Test2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   600
      TabIndex        =   5
      Top             =   1495
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Test1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Top             =   1120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ----------------------------------------------------------------------------------
' Autohigh sample
'
' Written by Marcel A. FRITSCH
' Copyright 2002 by Marcel A. FRITSCH
'
' This software is FREEWARE. You may use it for your own projects but you may not
' re-sell the original or the source code.
'
' No warranty express or implied, is given as to the use of this program.
' Use at your own risk.
'
' I am creating a global application hook. The callback procedure (WindowProc) is placed in
' the Module1 module. This call allows you to intercept all windows messages before they are
' processed, but you cannot change them. If there is a message WM_SETFOCUS or WM_KILLFOCUS
' the public procedure SETKILLFocus is called. In this procedure the control with the handle
' from the message is searched and the backcolor is changed if the type of the control is
' matching with the types that should be processed (It is not very usefull to change the
' backcolor of Buttons).
'
' ATTENTION: When working in the IDE do not stop this program with the STOP-Button
'            because the unhook-function will not be executed and the IDE crashes.
' ----------------------------------------------------------------------------------
'
Private Sub Form_Load()
' Set global application hook
lHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf WindowProc, App.hInstance, App.ThreadID)
End Sub
Public Sub SETKILLFocus(msg As Long, CHwnd As Long)
On Local Error Resume Next
Dim Ctrl As Control
For Each Ctrl In Controls
Err.Clear
If CHwnd = Ctrl.hwnd Then
If Err.Number = 0 Then
     If msg = WM_SETFOCUS Then
        If (TypeOf Ctrl Is TextBox) Or _
           (TypeOf Ctrl Is ComboBox) Then
            Ctrl.BackColor = &H80000018
        End If
     Else   ' WM_KILLFOCUS
        If (TypeOf Ctrl Is TextBox) Or _
           (TypeOf Ctrl Is ComboBox) Then
            Ctrl.BackColor = &H80000005
         End If
     End If
     Exit For
  End If
End If
Next ' Ctrl
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' Give up global application hook
UnhookWindowsHookEx lHook
End Sub
