VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kill the Terrorists"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   307
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   1080
   End
   Begin VB.Label lblLeft 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   10
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Terrorist Camps To Left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -120
      TabIndex        =   9
      Top             =   0
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   4500
      Width           =   7455
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   5640
      TabIndex        =   6
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   5
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4560
      TabIndex        =   3
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   5160
      TabIndex        =   2
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "j"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   48
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   960
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   6720
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00004080&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   0
      Top             =   3000
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fire As Boolean
Dim ddown As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If fire = False Then
    lblB.Top = lblU.Top + 20
    lblB.Left = lblU.Left + 10
    fire = True
End If
End Sub

Private Sub Form_Load()
start
End Sub

Private Sub Timer1_Timer()
    If ddown = 0 Then
        Timer1.Interval = 0
        Dim a As Integer
        a = MsgBox("Well Done!", vbYesNo, "Game Over")
        If a = vbYes Then start
    End If
        
    
   
    lblU.Left = lblU.Left - 1
    If lblU.Left < 0 Then
         
            lblU.Left = Me.ScaleWidth
    End If
    
    If fire = True Then
        lblB.Top = lblB.Top + 2
        check_hit
    End If
    
    If lblB.Top > Me.ScaleHeight Then fire = False
    If fire = True Then
       lblB.Visible = True
    End If
End Sub


Private Sub check_hit()

    Dim i As Integer
    
    For i = 0 To 6
        If lblE(i).Visible Then
            If lblB.Top > lblE(i).Top Then
                If lblB.Left > lblE(i).Left Then
                    If lblB.Left < lblE(i).Left + lblE(i).Width Then
                        lblE(i).Visible = False
                        ddown = ddown - 1
                        lblLeft.Caption = ddown
                        fire = False
                        lblB.Visible = False
                       
                    End If
                End If
            End If
        End If
        
    Next i

End Sub

Private Sub start()
    ddown = 7
    lblLeft.Caption = 7
    lblB.Visible = False
    Timer1.Interval = 6
    Dim i As Integer
       
    For i = 0 To 6
         lblE(i).Visible = True
    Next i
End Sub
