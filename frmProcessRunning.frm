VERSION 5.00
Begin VB.Form frmProcessRunning 
   BackColor       =   &H00B37A06&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form 1"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProcessList 
      BackColor       =   &H00C0C0C0&
      Caption         =   "List Running Processes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2760
      MouseIcon       =   "frmProcessRunning.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "frmProcessRunning.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   2460
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B37A06&
      Caption         =   "Results / Messages"
      ForeColor       =   &H8000000E&
      Height          =   3210
      Left            =   120
      TabIndex        =   3
      Top             =   2085
      Width           =   5085
      Begin VB.TextBox txtResults 
         Height          =   2895
         Left            =   165
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   210
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdFindProcess 
      BackColor       =   &H00C0C0C0&
      Caption         =   "See If Process is Running"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      MouseIcon       =   "frmProcessRunning.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "frmProcessRunning.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   2460
   End
   Begin VB.TextBox txtProcess 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Process Name (.exe)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   75
      Width           =   5235
   End
End
Attribute VB_Name = "frmProcessRunning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '***************************************************************************
    '
    ' Checks if a process is running on your computer
    ' Based on the WMI (Windows Management Instrurmentation) code on MSDN
    '
    ' Mark Mokoski
    ' 15-APR-2005
    ' www.rjillc.com
    '
    ' This Project requires WMI (Windows Management Instrurmentation).
    ' WMI is part of Windows 2000, XP
    '
    ' For more information see the MSDN Web site
    ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_tasks__processes.asp
    '
    '****************************************************************************
    Option Explicit
    
Private Sub cmdFindProcess_Click()

    'Find if a process (task)is running

        If txtProcess.Text <> "" Then

                If IsProcessRunning(txtProcess.Text) = True Then

                    txtResults.Text = "Process " & UCase(txtProcess.Text) & " is running on this computer "
                Else
                    txtResults.Text = "Process " & UCase(txtProcess.Text) & " is not running on this computer "

                End If

        Else
            txtResults.Text = "Please Enter a Valid Process Name (xyz.exe)"
        End If

End Sub

Private Sub cmdProcessList_Click()

    'Get a list of all running processes (tasks)

    Dim x            As Integer

    ListRunningProcesses
    txtResults.Text = ""

        For x = 0 To ProcArraySize
            txtResults.Text = txtResults & ProcessArray(x) & vbCrLf
        Next x

    txtResults.Text = txtResults & "Total Processes Running = " & ProcArraySize & vbCrLf
    txtResults.SelStart = Len(txtResults.Text) + 1
    
End Sub

Private Sub Form_Load()

    Me.Caption = App.Title & " - Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    txtResults.Text = "Please Enter a Process Name (xyz.exe)" & _
    vbCrLf & "OR" & vbCrLf & _
    "Select List all Processes"

End Sub
