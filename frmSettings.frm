VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3195
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   3195
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtSnow 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Text            =   "Ho Ho Ho!"
         Top             =   3480
         Width           =   2415
      End
      Begin MSComctlLib.Slider sldWind 
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin MSComctlLib.Slider sldFlakes 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Max             =   500
         SelStart        =   300
         Value           =   300
      End
      Begin MSComctlLib.Slider sldSize 
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   2400
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   2
         Value           =   2
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Flake Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Snow Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label lblFlakes 
         Alignment       =   2  'Center
         Caption         =   "Flakes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblWind 
         Alignment       =   2  'Center
         Caption         =   "Wind"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mLoading As Boolean

Private Sub Form_Load()
    mLoading = True
    sldFlakes.Value = mFlakeNum
    lblFlakes.Caption = "Flakes (" & sldFlakes.Value & ")"
    mLoading = False
    Me.Left = frmMain.Left - Me.Width
    Me.Top = frmMain.Top
End Sub

Private Sub sldFlakes_Change()
    'If the number of flakes change
    'then reset frmMain
    If mLoading = False Then
        mFlakeNum = sldFlakes.Value
        frmMain.StopSnow = True
        frmMain.Cls
        Setup
        frmMain.StopSnow = False
    End If
End Sub

Private Sub sldFlakes_Scroll()
    lblFlakes.Caption = "Flakes (" & sldFlakes.Value & ")"
End Sub

Private Sub sldSize_Click()
    'If the flake size changes
    'then reset frmMain
    mtFlakeSize = sldSize.Value
    frmMain.StopSnow = True
    frmMain.Cls
    Setup
    frmMain.StopSnow = False
End Sub

Private Sub sldWind_Scroll()
    'This sub simply appears to add
    'wind to the snow.
    If sldWind.Value > 5 Then
        mRightWind = sldWind.Value - 3
        mLeftWind = 2
    ElseIf sldWind.Value < 5 Then
        mLeftWind = (7 - sldWind.Value)
        mRightWind = 2
    Else
        mLeftWind = 2
        mRightWind = 2
    End If
    
End Sub

Private Sub txtSnow_Change()
    'If the text changes, update frmMain
    'with the new text by resetting the form.
    mSnowText = txtSnow.Text
    frmMain.StopSnow = True
    frmMain.Cls
    Setup
    frmMain.StopSnow = False
End Sub
