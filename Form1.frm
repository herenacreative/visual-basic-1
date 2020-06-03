VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton keluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ComboBox txtkode 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   28
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtnama 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   27
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   1200
   End
   Begin VB.CommandButton cmdinput 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdbatal 
      Caption         =   "Batal"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   24
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox txtkembali 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   23
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox txtbayar 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txttharga 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   21
      Top             =   3000
      Width           =   1935
   End
   Begin VB.OptionButton op2 
      Caption         =   "Delivery order"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Jenis Pemebelian"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   3615
      Begin VB.TextBox txtbiaya 
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton op1 
         Caption         =   "Take Away"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Biaya Antar"
         BeginProperty Font 
            Name            =   "Stencil"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtjumbel 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtharga 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   14
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txttgl 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5520
      TabIndex        =   13
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtnamap 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtno 
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1800
      TabIndex        =   11
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label times 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   26
      Top             =   240
      Width           =   45
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Left            =   120
      Top             =   600
      Width           =   7455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kembali"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   10
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bayar"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   9
      Top             =   3600
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Harga"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   8
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah beli"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   7
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Harga Kue"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3360
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Kue"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Kue"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal pesan"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   3
      Top             =   840
      Width           =   1485
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nama pemesan"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No. pesan"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bread talk"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
txtno.Text = ""
txtnamap.Text = ""
txtkode.Text = ""
txtnama.Text = ""
txtharga.Text = ""
txtjumbel.Text = ""
txttharga.Text = ""
txtbayar.Text = ""
txtkembali.Text = ""
txtbiaya.Text = ""
op1.Value = 0
op2.Value = 0
op1.Enabled = True
op2.Enabled = True
End Sub
Sub aktif()
txtno.Enabled = True
txtnamap.Enabled = True
txtkode.Enabled = True
txtjumbel.Enabled = True
txtbayar.Enabled = True
op1.Value = False
op2.Value = False
End Sub
Sub non_aktif()
txttgl.Enabled = False
txtnama.Enabled = False
txtharga.Enabled = False
txttharga.Enabled = False
txtkembali.Enabled = False
txtbiaya.Enabled = False
End Sub

Private Sub cmdbatal_Click()
If MsgBox("Batalkan Pesanan ini?", vbQuestion + vbYesNo, "Pembatalan") = vbYes Then
bersih
End If
End Sub



Private Sub cmdinput_Click()
aktif
End Sub

Private Sub cmdkeluar_Click()
If MsgBox("Tutup Aplikasi ini?", vbQuestion + vbYesNo, "Keluar") = vbYes Then
End
End If
End Sub

Private Sub Form_Load()
non_aktif
txtkode.AddItem "B01"
txtkode.AddItem "B02"
txtkode.AddItem "B03"
txtkode.Text = "Pilih Kode..."
End Sub

Private Sub keluar_Click()
If MsgBox("Tutup Aplikasi ini?", vbqoestion + vbYesNo, "Keluar") + vbYes Then
End
End If
End Sub

Private Sub op1_Click()
    txtbiaya = 0
op2.Enabled = False
End Sub

Private Sub op2_Click()
    txtbiaya = 5000
op1.Enabled = False
End Sub

Private Sub Timer1_Timer()
times.Caption = Time
txttgl = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub txtbayar_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
 txtkembali = Val(txtbayar) - Val(txttharga)
 End If
End Sub

Private Sub txtjumbel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txttharga = Val(txtharga) * Val(txtjumbel) + Val(txtbiaya)
    txtbayar.SetFocus
End If

End Sub

Private Sub txtkode_Click()
Select Case txtkode.Text
Case "B01"
    txtnama.Text = "Manggo Bread"
    txtharga.Text = 8000
Case "B02"
    txtnama.Text = "Orange Bread"
    txtharga.Text = 10000
Case "B03"
    txtnama.Text = "Apple Bread"
    txtharga.Text = 5000
End Select
txtjumbel.SetFocus
End Sub

