VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmCons 
   Caption         =   "ADO Access SEEK"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   1020
      Width           =   1155
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&Seek"
      Height          =   315
      Left            =   4740
      TabIndex        =   11
      Top             =   1020
      Width           =   1155
   End
   Begin ComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   3075
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "16/12/2002"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:24 a.m."
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   2
            Bevel           =   2
            Text            =   "Elapsed Time:"
            TextSave        =   "Elapsed Time:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Personal Information:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1635
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   5835
      Begin VB.Label lApe 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   9
         Top             =   1140
         Width           =   3495
      End
      Begin VB.Label lNom 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lCI 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   6
         Top             =   780
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Surname:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1200
         TabIndex        =   5
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ID Number:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Look for Field:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3315
      Begin VB.TextBox nCI 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "ID Number (1000 to 5468)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   300
         TabIndex        =   1
         Top             =   360
         Width           =   2220
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Use the botons below and see the diference. For Extreme type in the last number."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   3420
      TabIndex        =   13
      Top             =   60
      Width           =   2475
   End
End
Attribute VB_Name = "frmCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Blanquea()
    lCI = ""
    lNom = ""
    lApe = ""
End Sub

Sub Ubica()
    lCI = Tpx!ID
    lNom = Trim(NoNulo(Tpx!pnombre))
    lApe = Trim(NoNulo(Tpx!papellido))
End Sub

Private Sub Etr_GotFocus()
    nCI.SetFocus
End Sub

Private Sub cmdEnd_Click()
    Tpx.Seek nCI
    If Not Tpx.EOF Then
        Ubica
    Else
        MsgBox "Given ID not exist in DB!", vbOKOnly, "Seek in Access!"
    End If
End Sub

Private Sub cmdFind_Click()
    Dim Tr1 As Variant
    Dim Tr2 As Variant
    
    Blanquea
    Screen.MousePointer = vbHourglass
    Tr1 = Time
    Tpx.Find "ID=" & nCI, 0, adSearchForward, 1
    Tr2 = Time
    SB.Panels(4).Text = DateDiff("s", Tr1, Tr2)
    Screen.MousePointer = vbDefault
    If Not Tpx.EOF Then
        Ubica
    Else
        MsgBox "Given ID not exist in DB!", vbOKOnly, "Seek in Access!"
    End If
End Sub

Private Sub Form_Load()
    Set Conn = New ADODB.Connection
    Set Tpx = New ADODB.Recordset
    
    With Conn
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = App.Path & "\bd2.mdb"
        .Open
    End With
    CenterForm Me
    
    Tpx.CursorLocation = adUseServer
    
    ' Here is the 'Secret'. Look for the last option
    Tpx.Open "MAESTRO", Conn, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    
    Tpx.Index = "cedula"
    nCI = ""
    Blanquea
    CenterForm Me
End Sub

Private Sub nCI_GotFocus()
    nCI.SelStart = 0
    nCI.SelLength = Len(nCI)
End Sub
