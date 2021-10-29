VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Vesta Project - Parallel Data Acquisition software."
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   ScaleHeight     =   6825
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Delete all previews measurements."
      Height          =   255
      Left            =   5520
      TabIndex        =   15
      Top             =   4200
      Width           =   3015
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   360
      Max             =   120
      Min             =   1
      TabIndex        =   11
      Top             =   6240
      Value           =   60
      Width           =   4815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Info"
      Height          =   1935
      Left            =   5400
      TabIndex        =   8
      Top             =   4560
      Width           =   3495
      Begin VB.Label DurataTXT 
         Caption         =   "Total expected measurements: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label TotEf 
         Caption         =   "Total measurements made: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   3015
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   100
      Left            =   9120
      Top             =   1560
   End
   Begin VB.HScrollBar TakeTime 
      Height          =   255
      Left            =   360
      Max             =   60
      Min             =   3
      TabIndex        =   6
      Top             =   5640
      Value           =   5
      Width           =   4815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   855
      Left            =   3960
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   855
      Left            =   2520
      TabIndex        =   4
      Top             =   4440
      Width           =   1335
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9120
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9120
      Top             =   360
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   5400
      TabIndex        =   1
      Top             =   480
      Width           =   3465
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3660
      Left            =   360
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   240
      Width           =   4860
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   5400
      X2              =   9000
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label2 
      Caption         =   $"cam.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5400
      TabIndex        =   14
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "USB camcorders detected:"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label DurataTX 
      Caption         =   "Experiment duration: 0 sec"
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label TakeAT 
      Caption         =   "Take at: 0 sec"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label trec 
      Caption         =   "..."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3960
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   360
      Top             =   6720
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /  Data acquisition              \________________________/       v1.00        |
' |                                                                               |
' |            Name:  Data acquisition                                            |
' |          Author:  Paul A. Gagniuc                                             |
' |                                                                               |
' |    Date Created:  September 2016                                              |
' |       Tested On:  Windows XP, Windows Vista, Windows 7, Windows 8             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |             Use:  diabetes prediction                                         |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'


Dim cTimer As timTimer 'Our timer

Dim pozaNR As Variant

Private Sub cmdClose_Click()
    modWebCam.ClosePreviewWindow
    Unload Me
End Sub

Private Sub cmdStart_Click()

    If Check1.Value = 1 Then Call File_Delete(App.path & "\measurements")

    ' Srart WebCam Capture
    
    If Me.List1.ListCount = 0 Then Exit Sub
    If List1.ListIndex = -1 Then Exit Sub
    
    modWebCam.OpenPreviewWindow List1.ListIndex, Me.Picture1
    
    
    Set cTimer = New timTimer 'Assign a new timer
    cTimer.StartTimer 'Start it
    Timer2.Enabled = True
    Timer1.Enabled = True
    
    
End Sub


Private Sub cmdStop_Click()
    ' Stop Capture
    Timer1.Enabled = False
    Timer2.Enabled = False
    cTimer.EndTimer 'Stop timer
    
    pozaNR = 0
    modWebCam.ClosePreviewWindow
End Sub


Private Sub Form_Load()
    modWebCam.LoadDeviceList Me.List1
    pozaNR = 0
    
    'List1.Selected (0)
End Sub

Private Sub Timer1_Timer()

    trec.Caption = Round(cTimer.Elapsed / 1000) & " seconds."
    a = Round(cTimer.Elapsed / 1000)
    
    If a = 0 Then a = 1
    
    If a Mod TakeTime.Value = 0 Then
        Timer2.Enabled = True
    End If
    
    If a Mod (HScroll1.Value * 60) = 0 Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        cTimer.EndTimer 'Stop timer
        modWebCam.ClosePreviewWindow
        Picture1.Print "GATA !"
        MsgBox "The program has finished taking measurements !"
    End If
End Sub

Private Sub Timer2_Timer()

    Timer2.Enabled = False

    If Dir(App.path & "\measurements", vbDirectory) = "" Then MkDir (App.path & "\measurements")
    Set Me.Image1.Picture = hDCToPicture(GetDC(modWebCam.hHwnd), 0, 0, 320, 240)
    pozaNR = pozaNR + 1
    Picture1.Picture = Image1.Picture
    SaveJPG Picture1, App.path + "\measurements\" & pozaNR & "_" & Format(Date, "ddmmyyyy") & "_" & Format(Time, "hhmmss") & ".jpg"
End Sub

Private Sub Timer3_Timer()
    TakeAT.Caption = "Take at: " & TakeTime.Value & " sec"
    DurataTXT.Caption = "Total expected measurements: " & Round((60 * HScroll1.Value) / TakeTime.Value) '& " images"
    TotEf.Caption = "Total measurements made: " & pozaNR '& " matrices"
    DurataTX.Caption = "Experiment duration: " & HScroll1.Value & " min"
End Sub
