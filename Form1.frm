VERSION 5.00
Object = "*\A..\..\..\..\..\IEEVEN~1\PROGRA~1\MICROS~1\VB98\IEMONI~1\IEMonitor.vbp"
Begin VB.Form Form1 
   Caption         =   "Test App"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13755
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   13755
   StartUpPosition =   3  'Windows Default
   Begin IEMonitor.IEEvents IEEvents1 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   873
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8445
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   13335
   End
   Begin VB.Label Label1 
      Caption         =   "IEEvents Test Application - Click Start then Open a new IE Window and watch the events fly!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   13335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = True
    Command3.Enabled = True
    IEEvents1.Enabled = True
    'IEEvents1.Browsers
End Sub

Private Sub Command2_Click()
    IEEvents1.Enabled = False
    Command1.Enabled = True
    Command2.Enabled = False
    Command3.Enabled = False
End Sub

Private Sub Command3_Click()
    IEEvents1.Refresh
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Form2.Show
End Sub

Private Sub IEEvents1_BrowserCreated(Browser As SHDocVw.InternetExplorer)
    AddEvent Browser, "New Browser Created"
    List1.AddItem "New count is " & IEEvents1.Browsers.Count
End Sub

Private Sub IEEvents1_BrowserDestroyed()
    List1.AddItem "Browser Destroyed"
    List1.AddItem "New count is " & IEEvents1.Browsers.Count
End Sub

Private Sub IEEvents1_BrowserNavigating(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    AddEvent Browser, "Start Navigation - " & CStr(URL)
    
End Sub

Private Sub IEEvents1_DocumentComplete(Browser As SHDocVw.InternetExplorer, pDisp As Object, URL As Variant)
    AddEvent Browser, "Document Complete - " & CStr(URL)
End Sub

Private Sub IEEvents1_DownLoadComplete(Browser As SHDocVw.InternetExplorer)
    AddEvent Browser, "Download Complete"
End Sub


Private Sub IEEvents1_FileDownload(Browser As SHDocVw.InternetExplorer, Cancel As Boolean)
    AddEvent Browser, "File Download"
End Sub

Private Sub IEEvents1_NavigateComplete(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant)
    AddEvent Browser, "Complete Navigation - " & CStr(URL)
End Sub

Private Sub IEEvents1_NavigateError(Browser As SHDocVw.InternetExplorer, ByVal pDisp As Object, URL As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)
    AddEvent Browser, "Navigation Error - " & CStr(URL)
End Sub

Private Sub IEEvents1_NewWindow(Browser As SHDocVw.InternetExplorer, ppDisp As Object, Cancel As Boolean)
    AddEvent Browser, "New Window  "
End Sub

Private Sub IEEvents1_OnFullScreen(Browser As SHDocVw.InternetExplorer, ByVal FullScreen As Boolean)
    AddEvent Browser, IIf(FullScreen, "Full Screen", "Not FullScreen")
End Sub

Private Sub IEEvents1_ProgressChange(Browser As SHDocVw.InternetExplorer, ByVal Progress As Long, ByVal ProgressMax As Long)
    AddEvent Browser, "Progress Change - " & Progress
End Sub

Private Sub IEEvents1_TitleChange(Browser As SHDocVw.InternetExplorer, ByVal Text As String)
    AddEvent Browser, "Title Change - " & Text
End Sub

Private Sub IEEvents1_WindowClosing(Browser As SHDocVw.InternetExplorer, ByVal IsChildWindow As Boolean, Cancel As Boolean)
    AddEvent Browser, "Window Closing"
End Sub

Sub AddEvent(b As SHDocVw.InternetExplorer, strText As String)
    List1.AddItem "[" & Now & "] Browser " & b.hwnd & " - " & strText
End Sub

Private Sub Timer1_Timer()
End Sub
