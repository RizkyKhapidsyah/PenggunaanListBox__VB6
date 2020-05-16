VERSION 5.00
Begin VB.Form FormUtama 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Penggunaan ListBox"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3120
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapusSemua 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdTambahSemua 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Command1"
      Height          =   735
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame FrameLstKunjung 
      Caption         =   "Frame2"
      Height          =   4335
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstKunjung 
         Height          =   3765
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2655
      End
   End
   Begin VB.Frame FrameLstKota 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstKota 
         Height          =   3660
         ItemData        =   "FormUtama.frx":0000
         Left            =   120
         List            =   "FormUtama.frx":0002
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FormUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHapus_Click()
On Error GoTo PesanError
    
    Dim CurItem As Integer
    CurItem = 0
    
    Do
        If lstKunjung.Selected(CurItem) Then
            lstKota.AddItem lstKunjung.List(CurItem)
            lstKunjung.RemoveItem (CurItem)
        Else
               CurItem = CurItem + 1
        End If
    Loop Until CurItem = lstKunjung.ListCount
    Exit Sub
    
PesanError:
    MsgBox Err.Description
End Sub

Private Sub cmdHapusSemua_Click()
Dim i As Integer

    For i = 0 To lstKunjung.ListCount - 1
        lstKota.AddItem lstKunjung.List(i)
    Next i
    
    lstKunjung.Clear
End Sub

Private Sub cmdKeluar_Click()
    End
End Sub

Private Sub cmdTambah_Click()
On Error GoTo PesanError
    
    Dim CurItem As Integer
    CurItem = 0
    
    Do
        If lstKota.Selected(CurItem) Then
            lstKunjung.AddItem lstKota.List(CurItem)
            lstKota.RemoveItem (CurItem)
        Else
               CurItem = CurItem + 1
        End If
    Loop Until CurItem = lstKota.ListCount
    Exit Sub
    
PesanError:
    MsgBox Err.Description

End Sub

Private Sub cmdTambahSemua_Click()
Dim i As Integer

    For i = 0 To lstKota.ListCount - 1
        lstKunjung.AddItem lstKota.List(i)
    Next i
    
    lstKota.Clear
End Sub

Private Sub Form_Load()
With Me
    .FrameLstKota.Caption = "Kota Di Indonesia"
    .FrameLstKunjung.Caption = "Kota Yang Dikunjungi"
    .lstKota.Clear
    
    .lstKota.AddItem "Medan"
    .lstKota.AddItem "Jakarta"
    .lstKota.AddItem "Surabaya"
    .lstKota.AddItem "Bandung"
    .lstKota.AddItem "Pontianak"
    .lstKota.AddItem "Banjar Masin"
    .lstKota.ListIndex = 1
    
    .cmdHapus.Caption = "<"
    .cmdHapusSemua.Caption = "<<"
    .cmdKeluar.Caption = "Keluar"
    .cmdTambah.Caption = ">"
    .cmdTambahSemua.Caption = ">>"
End With
    
    
    
    

End Sub
