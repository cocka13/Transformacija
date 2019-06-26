VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opcije"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSpremi 
      Caption         =   "Zapamti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   5775
   End
   Begin VB.TextBox txtRazmak 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.OptionButton radioZona1 
      Caption         =   "5. - 6. - 7. zona"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton radioZona2 
      Caption         =   "Sve u 16.30"
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.OptionButton radioZona3 
      Caption         =   "16.30 u 5. zonu"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton radioZona4 
      Caption         =   "16.30 u 6. zonu"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.OptionButton radioZona5 
      Caption         =   "16.30 u 7. zonu"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Izbor transformacije"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5775
   End
   Begin VB.Frame Frame2 
      Caption         =   "Formatirani ispis koordinata"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   5775
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Odvajanje koordinata"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSpremi_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Select Case radioZona
    Case 1
        radioZona1.Value = True
    Case 2
        radioZona2.Value = True
    Case 3
        radioZona3.Value = True
    Case 4
        radioZona4.Value = True
    Case 5
        radioZona5.Value = True
End Select
End Sub

Private Sub radioZona1_Click()
If radioZona1.Value = True Then radioZona = 1
End Sub

Private Sub radioZona2_Click()
If radioZona2.Value = True Then radioZona = 2
End Sub

Private Sub radioZona3_Click()
If radioZona3.Value = True Then radioZona = 3
End Sub

Private Sub radioZona4_Click()
If radioZona4.Value = True Then radioZona = 4
End Sub

Private Sub radioZona5_Click()
If radioZona5.Value = True Then radioZona = 5
End Sub

