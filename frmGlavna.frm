VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGlavna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transformacija"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "frmGlavna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Datoteka izvor podataka"
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   5295
      Begin VB.CommandButton comStartAuto 
         Caption         =   "Transformiraj"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   4815
      End
      Begin VB.CommandButton comIzlazna 
         Caption         =   "Odabir"
         Height          =   285
         Left            =   3840
         TabIndex        =   19
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton comUlazna 
         Caption         =   "Odabir"
         Height          =   285
         Left            =   3840
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtIzlaznaDatoteka 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtUlaznaDatoteka 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.OptionButton radioZona5 
      Caption         =   "16.30 u 7. zonu"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
   End
   Begin VB.OptionButton radioZona4 
      Caption         =   "16.30 u 6. zonu"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
   End
   Begin VB.OptionButton radioZona3 
      Caption         =   "16.30 u 5. zonu"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   3240
      Width           =   1455
   End
   Begin VB.OptionButton radioZona2 
      Caption         =   "Sve u 16.30"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton radioZona1 
      Caption         =   "5. - 6. - 7. zona"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   2400
      Value           =   -1  'True
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog comDialog 
      Left            =   5280
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtIspisKoordinata 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1560
      Width           =   4815
   End
   Begin VB.CommandButton cmdStartRucno 
      Caption         =   "Transformiraj"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox txtXkoordinata 
      Height          =   285
      Left            =   2880
      MaxLength       =   11
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtYkoordinata 
      Height          =   285
      Left            =   360
      MaxLength       =   11
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manualni unos podataka"
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5295
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3840
         TabIndex        =   0
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opcije"
      Height          =   1935
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   5295
      Begin VB.TextBox txtRazmak 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Odvajanje koordinata:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmGlavna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdStartRucno_Click()
Dim xKoordinata As Double
Dim yKoordinata As Double
Dim yKoordinataText As String
Dim xKoordinataText As String
Dim rezultat As Boolean

xKoordinata = Val(TockaZarez(txtXkoordinata))
yKoordinata = Val(TockaZarez(txtYkoordinata))

rezultat = modGenerator.Trans(xKoordinata, yKoordinata)

yKoordinataText = TockaZarez(Trim(Str(Round(yKoordinata, 3))))
xKoordinataText = TockaZarez(Trim(Str(Round(xKoordinata, 3))))

If rezultat Then
    txtIspisKoordinata.Text = yKoordinataText + txtRazmak.Text + xKoordinataText
    txtXkoordinata = ""
    txtYkoordinata = ""
End If
End Sub

Private Sub comIzlazna_Click()
    comDialog.FileName = txtIzlaznaDatoteka.Text
    comDialog.ShowSave
    txtIzlaznaDatoteka.Text = comDialog.FileName
End Sub

Private Sub comStartAuto_Click()
    Call RadiNesto
End Sub

Private Sub comUlazna_Click()
    comDialog.FileName = txtUlaznaDatoteka.Text
    comDialog.ShowOpen
    txtUlaznaDatoteka.Text = comDialog.FileName
End Sub

Private Sub Form_Terminate()
    Call Ending
End Sub
Private Sub Ending()
    Close All
    End
End Sub

Private Sub radioZona2_Click()
    If radioZona2 Then radioZona = 2
End Sub
Private Sub radioZona3_Click()
    If radioZona3 Then radioZona = 3
End Sub
Private Sub radioZona4_Click()
    If radioZona4 Then radioZona = 4
End Sub
Private Sub radioZona5_Click()
    If radioZona5 Then radioZona = 5
End Sub
