VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New_List_Beta1-NO COMMENT"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "open txt"
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   5760
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   6480
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Project1.uj_lista uj_lista0 
      Height          =   2760
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4868
      KijeloltHatterSzin=   8388608
      KijeloltBetuSzin=   12648384
      ListaBetuSzin   =   12648384
      GorgetesiSebesseg=   1
   End
   Begin Project1.uj_lista uj_lista1 
      Height          =   2760
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4868
      KijeloltHatterSzin=   255
      KijeloltBetuSzin=   0
      ListaHatterSzin =   32768
      ListaBetuSzin   =   0
      GorgetesiSebesseg=   25
      Betûk           =   4
   End
   Begin Project1.uj_lista uj_lista2 
      Height          =   2760
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4868
      KijeloltHatterSzin=   49152
      KijeloltBetuSzin=   8421504
      ListaHatterSzin =   16777215
      ListaBetuSzin   =   8421504
      GorgetesiSebesseg=   65
      Sorbarendezés   =   1
      Betûk           =   2
   End
   Begin Project1.uj_lista uj_lista3 
      Height          =   2760
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4868
      KijeloltHatterSzin=   128
      KijeloltBetuSzin=   16777088
      ListaHatterSzin =   8421376
      ListaBetuSzin   =   16777088
      GorgetesiSebesseg=   100
      Sorbarendezés   =   1
      Betûk           =   2
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abc As Integer



Private Sub Command1_Click()

'txt file open
With cmd
        .DialogTitle = "-->Txt file megnyitása>---"
        .CancelError = False
        .Filter = "txt file-ok (*.txt)|*.txt"
        .ShowOpen
        If Len(.FileName) = 0 Then
           Exit Sub
        End If
        file = .FileName
End With


'open it and refresh
Open file For Input As #1
Do While EOF(1) = 0
Input #1, aa$
  uj_lista0.AddItem aa$
  uj_lista1.AddItem aa$
  uj_lista2.AddItem aa$
  uj_lista3.AddItem aa$
Loop
Close #1

uj_lista0.Refresh_It
uj_lista1.Refresh_It
uj_lista2.Refresh_It
uj_lista3.Refresh_It


End Sub

Private Sub new_list0_KEYDOWN(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
   MsgBox new_list0.Text
End If

End Sub

