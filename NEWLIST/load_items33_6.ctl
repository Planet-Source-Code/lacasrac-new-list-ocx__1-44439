VERSION 5.00
Begin VB.UserControl uj_lista 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   3180
   ScaleWidth      =   3495
   Begin VB.Timer Fel_Le_Timer2 
      Left            =   480
      Top             =   2760
   End
   Begin VB.PictureBox fel_1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   960
      Picture         =   "load_items33_6.ctx":0000
      ScaleHeight     =   135
      ScaleWidth      =   180
      TabIndex        =   12
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox kozep_1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1200
      Picture         =   "load_items33_6.ctx":037A
      ScaleHeight     =   750
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox le_1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   1440
      Picture         =   "load_items33_6.ctx":0748
      ScaleHeight     =   150
      ScaleWidth      =   165
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox le_2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   2280
      Picture         =   "load_items33_6.ctx":0AC4
      ScaleHeight     =   150
      ScaleWidth      =   165
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox kozep_2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   2040
      Picture         =   "load_items33_6.ctx":0E40
      ScaleHeight     =   750
      ScaleWidth      =   180
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox fel_2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1800
      Picture         =   "load_items33_6.ctx":120E
      ScaleHeight     =   135
      ScaleWidth      =   180
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox kozep 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3240
      Picture         =   "load_items33_6.ctx":1588
      ScaleHeight     =   750
      ScaleWidth      =   180
      TabIndex        =   6
      Top             =   240
      Width           =   180
   End
   Begin VB.PictureBox scr_hat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   3240
      ScaleHeight     =   135
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   5
      Top             =   165
      Width           =   150
   End
   Begin VB.PictureBox le_gomb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   3240
      Picture         =   "load_items33_6.ctx":1956
      ScaleHeight     =   150
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   2520
      Width           =   165
   End
   Begin VB.PictureBox fel_gomb 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3240
      Picture         =   "load_items33_6.ctx":1CD2
      ScaleHeight     =   135
      ScaleWidth      =   180
      TabIndex        =   3
      Top             =   0
      Width           =   180
   End
   Begin VB.Timer Fel_Le_Timer 
      Left            =   0
      Top             =   2760
   End
   Begin VB.PictureBox hatter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   0
      ScaleHeight     =   2745
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin VB.Shape selector 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         DrawMode        =   6  'Mask Pen Not
         Height          =   210
         Left            =   30
         Top             =   0
         Width           =   3195
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   0
         Width           =   3200
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 db elem van"
      Enabled         =   0   'False
      ForeColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "uj_lista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  NEW_LIST usercontrol
'  http://www.alvilag.hu/lacasrac
'  created by Kozari Laszlo in 2002 okt/nov
'  lacasrac@wap.hu
'  for VB_6.0 users

''''''''Mistakes''''''''''''''''''''''''''''''''''''''
' 1.)    b, "down btn"
' this is a free version code
' if u like this, u rewrite it!
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Lista elem values
Dim le, fel                             As Boolean
Dim SZ_BE                               As Boolean
Dim sor_sz                              As Integer
Dim felso, also                         As Integer
Dim last_item                           As Integer
Dim old_index, indexs                   As Integer
Dim szelekted, a_szelekted              As Integer
Dim entry_point                         As Byte
Dim visible_items                       As Byte
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Csuszka->button,button_fel,x,y
'cs_b=scroller_btn,cs_b_f=scroller_up-...
Dim cs_b
Dim cs_b_F
Dim cs_x, cs_y                          As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'm_= modus->>vars
Dim m_KijeloltHatterSzin                As OLE_COLOR
Dim m_KijeloltBetuSzin                  As OLE_COLOR
Dim m_ListaHatterSzin                   As OLE_COLOR
Dim m_ListaBetuSzin                     As OLE_COLOR
Dim m_Rendez                            As enumRendez
Dim m_Betuk                             As enumBetuk
Dim m_GorgetesiSebesseg                 As Byte
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'for the sorting / rendezéshez [max. 5000 elem]
'max 5000 items
Dim tomb(5000)                          As String
Dim nevek(5000)                         As String
Dim nevek2(5000)                        As String
Dim inde_x(5000)                        As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Alap konstansok
'Base consts
Const m_def_Betuk = 0
Const m_def_KijeloltHatterSzin = &H0&
Const m_def_KijeloltBetuSzin = &HFFFFFF
Const m_def_ListaHatterSzin = &H0&
Const m_def_ListaBetuSzin = &HFFFFFF
Const m_def_GorgetesiSebesseg = "50"
Const m_def_Rendez = 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Események a listához /hozzárendelések/
'events for LIST
Event CLICK()
Event DBLCLICK()
Event KEYDOWN(KeyCode As Integer, Shift As Integer)
Event KEYPRESS(KeyAscii As Integer)
Event KEYUP(KeyCode As Integer, Shift As Integer)
Event MOUSEDOWN(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MOUSEMOVE(Button As Integer, Shift As Integer, X As Single, Y As Single, ListItem As String)
Event MOUSEUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Rendezés típusok
'Sorting by...
Enum enumRendez
     ABC_Sorba = 0
     CBA_Sorba = 1
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Betûtípusok
'Font types
'0:normal
'1:small letters
'2:big -""-
'3:inverse
'4:funny
Enum enumBetuk
     Normál = 0
     Kisbetûk = 1
     Nagybetûk = 2
     Inverz = 3
     Tréfás = 4
End Enum
''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Get KijelöltBetûSzín() As OLE_COLOR
  KijelöltBetûSzín = m_KijeloltBetuSzin
End Property

Public Property Get KijelöltHáttérSzín() As OLE_COLOR
  KijelöltHáttérSzín = m_KijeloltHatterSzin
End Property

Public Property Get Rendezés() As enumRendez
  Rendezés = m_Rendez
End Property

Public Property Get Betûk() As enumBetuk
  Betûk = m_Betuk
End Property

Public Property Let Rendezés(ByVal New_rendez As enumRendez)
  m_Rendez = New_rendez
  PropertyChanged "Sorbarendezés"
End Property

Public Property Let Betûk(ByVal New_Betuk As enumBetuk)
  m_Betuk = New_Betuk
  PropertyChanged "Betûk"
End Property

Public Property Let KijelöltHáttérSzín(ByVal New_KijeloltHatterSzin As OLE_COLOR)
  m_KijeloltHatterSzin = New_KijeloltHatterSzin
  PropertyChanged "KijeloltHatterSzin"
  Refresh_All
End Property
Public Property Let KijelöltBetûSzín(ByVal New_KijeloltBetuSzin As OLE_COLOR)
  m_KijeloltBetuSzin = New_KijeloltBetuSzin
  PropertyChanged "KijeloltBetuSzin"
  Refresh_All
End Property

Public Property Get ListaHáttérSzín() As OLE_COLOR
  ListaHáttérSzín = m_ListaHatterSzin
End Property

Public Property Get GörgetésiSebesség() As Byte
  'értékátadás...
  GörgetésiSebesség = m_GorgetesiSebesseg
End Property

Public Property Let GörgetésiSebesség(ByVal New_GorgetesiSebesseg As Byte)
  'megváltoztattuk az értéket...
  m_GorgetesiSebesseg = New_GorgetesiSebesseg
  PropertyChanged "GorgetesiSebesseg"
  If m_GorgetesiSebesseg > 100 Then MsgBox "A görgetési sebesség értéke 1-100 közötti lehet csak...": m_GorgetesiSebesseg = 100
  
  Refresh_All
End Property


Public Property Get ListaBetûSzín() As OLE_COLOR
  ListaBetûSzín = m_ListaBetuSzin
End Property

Public Property Let ListaHáttérSzín(ByVal New_ListaHatterSzin As OLE_COLOR)
  m_ListaHatterSzin = New_ListaHatterSzin
  PropertyChanged "ListaHatterSzin"
  Refresh_All
End Property
Public Property Let ListaBetûSzín(ByVal New_ListaBetuSzin As OLE_COLOR)
  m_ListaBetuSzin = New_ListaBetuSzin
  PropertyChanged "ListaBetuSzin"
  Refresh_All
End Property



Private Sub fel_gomb_KeyDown(KeyCode As Integer, Shift As Integer)

Call hatter_KeyDown(KeyCode, 0)

End Sub

Private Sub fel_gomb_KeyUp(KeyCode As Integer, Shift As Integer)

Fel_Le_Timer.Interval = 0

End Sub

Private Sub fel_gomb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  a_szelekted = szelekted
  fel = True
  le = False
  Fel_Le_Timer2.Interval = m_GorgetesiSebesseg
End If

End Sub

Private Sub fel_gomb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

fel_gomb.Picture = fel_2.Picture
kozep.Picture = kozep_1.Picture
le_gomb.Picture = le_1.Picture

End Sub

Private Sub fel_gomb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Fel_Le_Timer2.Interval = 0

End Sub

Private Sub Fel_Le_Timer2_Timer()

Call Index_Szamitas
'''''''''''''''''''''''''''''''''''''''''''''''''''
'FEL GOMB
If fel = True Then
 
    indexs = 0
    szelekted = szelekted - 1
    
    If szelekted < 0 Then szelekted = 0: Exit Sub
    
    For a = 0 To visible_items - 1
            Label1(a).Caption = nevek(inde_x(szelekted + a - 1 + 1))
      Refresh
    Next a
    
ElseIf (le = True) And (fel = False) Then
'''''''''''''''''''''''''''''''''''''''''''''''''
'LE GOMB
    szelekted = szelekted + 1
    
    If szelekted >= last_item Then szelekted = last_item: Exit Sub
    
    For a = 0 To visible_items - 1
        Label1(a).Caption = nevek(inde_x(szelekted + a + 1))
     Refresh
    Next a
End If

Call Csuszka_Mozog

entry_point = 7: Call Elemek_Csuszkahoz

Exit Sub '------------------------------------

If (szelekted_2 >= a_szelekted) And _
   (szelekted_2 - (visible_items - 1) <= a_szelekted) Then
 
   For a = 0 To visible_items - 1
       If Label1(a).Caption = nevek(inde_x(a_szelekted)) Then
            Label1(a).BackColor = m_KijeloltHatterSzin
            Label1(a).ForeColor = m_KijeloltBetuSzin
            selector.Top = Label1(a).Top
            selector.Visible = True
      Else
         Label1(a).BackColor = m_ListaHatterSzin
         Label1(a).ForeColor = m_ListaBetuSzin
      End If
   Next a
 
Else
    selector.Visible = False
    Label1(indexs).BackColor = m_ListaHatterSzin
    Label1(indexs).ForeColor = m_ListaBetuSzin
End If

old_index = indexs

Call Csuszka_Mozog
    
End Sub

Private Sub hatter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

fel_gomb.Picture = fel_1.Picture
kozep.Picture = kozep_1.Picture
le_gomb.Picture = le_1.Picture

End Sub

Private Sub kozep_KeyDown(KeyCode As Integer, Shift As Integer)
  
  Call hatter_KeyDown(KeyCode, 0)

End Sub

Private Sub kozep_KeyUp(KeyCode As Integer, Shift As Integer)

Fel_Le_Timer.Interval = 0

End Sub


Private Sub kozep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

cs_b = -1
cs_x = X
cs_y = Y

End Sub


Private Sub kozep_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

fel_gomb.Picture = fel_1.Picture
kozep.Picture = kozep_2.Picture
le_gomb.Picture = le_1.Picture

entry_point = 0

If (Button = 1) And (cs_b) Then
    
 'tapadások...
 If kozep.Top < fel_gomb.Top + fel_gomb.Height Then
    kozep.Top = fel_gomb.Top + fel_gomb.Height
    cs_b = 0
    cs_b_F = -1
 ElseIf kozep.Top > le_gomb.Top - kozep.Height Then
    kozep.Top = le_gomb.Top - kozep.Height
    cs_b = 0
    cs_b_F = -1
 Else
    kozep.Move kozep.Left, kozep.Top + Y - cs_y, kozep.Width, kozep.Height
    Call Elemek_Csuszkahoz
 End If

End If

'tapadások ki...felengedések...
If ((kozep.Top > (fel_gomb.Top + fel_gomb.Height + cs_y - Y) And cs_b_F = -1)) And kozep.Top <= fel_gomb.Top + fel_gomb.Height + 50 Or _
    (kozep.Top < (le_gomb.Top - kozep.Height + cs_y - Y) And cs_b_F = -1) And kozep.Top > fel_gomb.Top + fel_gomb.Height + cs_y Then
    cs_b = -1
End If

End Sub

Private Sub kozep_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

cs_b = 0

End Sub


Private Sub Label1_Click(Index As Integer)
RaiseEvent CLICK

End Sub

Private Sub Label1_DblClick(Index As Integer)
RaiseEvent DBLCLICK

End Sub


Private Sub le_gomb_KeyDown(KeyCode As Integer, Shift As Integer)

Call hatter_KeyDown(KeyCode, 0)

End Sub

Private Sub le_gomb_KeyUp(KeyCode As Integer, Shift As Integer)

Fel_Le_Timer.Interval = 0

End Sub

Private Sub le_gomb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
  a_szelekted = szelekted
  
  fel = False
  le = True
  
  Fel_Le_Timer2.Interval = m_GorgetesiSebesseg
End If

End Sub

Private Sub le_gomb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

fel_gomb.Picture = fel_1.Picture
kozep.Picture = kozep_1.Picture
le_gomb.Picture = le_2.Picture

End Sub

Private Sub le_gomb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Fel_Le_Timer2.Interval = 0
End Sub

Private Sub hatter_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KEYDOWN(KeyCode, Shift)

SZ_BE = True

If last_item > 0 Then

Select Case KeyCode
 Case 33, 37:  Call Other_Keys("PAGE_UP")
 Case 34, 39:  Call Other_Keys("PAGE_DOWN")
 Case 35: Call Other_Keys("EN_D")
 Case 36: Call Other_Keys("HOM_E")
 Case 38: fel = True: le = False: Fel_Le_Timer.Interval = m_GorgetesiSebesseg
 Case 40: fel = False: le = True: Fel_Le_Timer.Interval = m_GorgetesiSebesseg
End Select

End If

End Sub

Private Sub hatter_KeyUp(KeyCode As Integer, Shift As Integer)
Fel_Le_Timer.Interval = 0
End Sub

Private Sub hatter_Resize()

le_gomb.Left = hatter.Left + hatter.Width
le_gomb.Top = hatter.Top + hatter.Height - le_gomb.Height

fel_gomb.Left = le_gomb.Left
fel_gomb.Top = hatter.Top

kozep.Left = le_gomb.Left
kozep.Top = hatter.Top + fel_gomb.Height

scr_hat.Left = fel_gomb.Left
scr_hat.Top = fel_gomb.Height
scr_hat.Height = le_gomb.Top - (fel_gomb.Top + fel_gomb.Height)
scr_hat.Width = kozep.Width

Label1(0).Width = hatter.Width - 110
selector.Width = Label1(0).Width

'240=120*2->> kihagyás fenn,lenn
visible_items = Int((Height) / Label1(0).Height)
hatter.Height = (visible_items * Label1(0).Height) + 30

End Sub


Private Sub Load_Item(name)

Select Case Len(CStr(sor_sz))
    Case 1: sor_sz2 = CStr(sor_sz) + "." + String(6, " ")
    Case 2: sor_sz2 = CStr(sor_sz) + "." + String(4, " ")
    Case 3: sor_sz2 = CStr(sor_sz) + "." + String(2, " ")
    Case 4: sor_sz2 = CStr(sor_sz) + "." + String(0, " ")
End Select


Select Case m_Betuk
    Case 0: 'normal
      tomb(sor_sz) = " " + UCase(Left(name, 1)) + Mid$(name, 2, Len(name))
    Case 1: 'kisbetûk
      tomb(sor_sz) = " " + LCase(name)
    Case 2: 'nagybetûk
      tomb(sor_sz) = " " + UCase(name)
    Case 3: 'inverz
      tomb(sor_sz) = " " + LCase(Left(name, 1)) + UCase(Mid$(name, 2, Len(name)))
    Case 4: 'tRéFáS:)
      For t = 1 To Len(name)
        If (t Mod 2) = 0 Then name2 = name2 + UCase(Mid$(name, t, 1))
        If Not (t Mod 2) = 0 Then name2 = name2 + LCase(Mid$(name, t, 1))
        tomb(sor_sz) = " " + name2
      Next t
End Select

last_item = sor_sz
sor_sz = sor_sz + 1

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MOUSEDOWN(Button, Shift, X, Y)

If Button = 1 Then
 If Index = old_index Then Exit Sub

 selector.Top = Label1(Index).Top
 selector.Visible = True
 
 Label1(Index).BackColor = m_KijeloltHatterSzin
 Label1(Index).ForeColor = m_KijeloltBetuSzin
 
 ''''ez
 If szelekted > -1 Then
   Call Index_Szamitas
    
  szelekted_y = Int((kozep.Top - fel_gomb.Height) / ((scr_hat.Height - kozep.Height) / last_item))
   
   If szelekted_y > visible_items Then
      a_szelekted = szelekted_y - (visible_items - indexs)
   Else
      a_szelekted = szelekted_y + visible_items - (visible_items - indexs)
   End If
   
   szelekted = szelekted_y + indexs
 End If
 
 Label1(old_index).BackColor = m_ListaHatterSzin
 Label1(old_index).ForeColor = m_ListaBetuSzin
 old_index = Index
End If

End Sub

Private Sub Fel_Le_Timer_Timer()

Dim s                                        As Integer

Call Index_Szamitas

'''''''''''''''''''''''''''''''''''''''''''''''''''
'FEL GOMB

If fel = True Then
 
 If (indexs >= 1) And (indexs <= visible_items - 1) Then
    
    If SZ_BE = True Then
       szelekted = szelekted - 1
       indexs = indexs - 1

       If szelekted < 0 Then szelekted = 0: Exit Sub
       If szelekted >= last_item Then Exit Sub
     
        Call Kivalaszt("0")
         selector.Visible = True
         selector.Move Label1(indexs).Left, Label1(indexs).Top, Label1(indexs).Width, Label1(indexs).Height
         First = 0
       
    End If
'''''''''
 Else
    indexs = 0
    szelekted = szelekted - 1

     If szelekted < 0 Then szelekted = 0
    
      For a = 0 To visible_items - 1
            Label1(a).Caption = nevek(inde_x(szelekted + a - 1 + 1))
             Refresh
            
            If SZ_BE = True Then
              selector.Visible = True
              selector.Move Label1(0).Left, Label1(0).Top, Label1(0).Width, Label1(0).Height
              First = 0
            End If
       
       Refresh
    Next a
 End If

ElseIf (le = True) And (fel = False) Then

'''''''''''''''''''''''''''''''''''''''''''''''''
'LE GOMB
 If indexs < visible_items - 1 Then
    
    If SZ_BE = True Then
     If szelekted >= last_item Then Exit Sub
     
       szelekted = szelekted + 1
       indexs = indexs + 1
    
       Call Kivalaszt("0")
       selector.Visible = True
       selector.Move Label1(indexs).Left, Label1(indexs).Top, Label1(indexs).Width, Label1(indexs).Height
      First = 0
    End If
 
 Else
    szelekted = szelekted + 1

    If szelekted > last_item Then szelekted = last_item: Exit Sub
    
     For a = 0 To visible_items - 1
            Label1(a).Caption = nevek(inde_x(szelekted + a - visible_items + 1))
             Refresh
     
             If SZ_BE = True Then
               selector.Visible = True
               selector.Move Label1(visible_items - 1).Left, Label1(visible_items - 1).Top, Label1(visible_items - 1).Width, Label1(visible_items - 1).Height
               First = 0
             End If
     
      Refresh
    Next a
 End If
End If

If First = 0 Then a_szelekted = szelekted: First = 1

Call Csuszka_Mozog

End Sub

Private Sub UserControl_InitProperties()

'ha változtatunk akkor...
m_KijeloltHatterSzin = m_def_KijeloltHatterSzin
m_KijeloltBetuSzin = m_def_KijeloltBetuSzin
m_ListaHatterSzin = m_def_ListaHatterSzin
m_ListaBetuSzin = m_def_ListaBetuSzin
m_GorgetesiSebesseg = m_def_GorgetesiSebesseg
m_Rendez = m_def_Rendez
m_Betuk = m_def_Betuk


End Sub

Public Property Get ListCount() As Integer

ListCount = last_item
End Property

Public Property Get ListIndex() As Integer

ListIndex = szelekted
End Property


Public Property Get Text(Optional mit) As String

txt = nevek(inde_x(a_szelekted))
Text = txt

End Property


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

fel_gomb.Picture = fel_1.Picture
kozep.Picture = kozep_1.Picture
le_gomb.Picture = le_1.Picture

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

'beolvassa a konstansokat...
m_KijeloltHatterSzin = PropBag.ReadProperty("KijeloltHatterSzin", m_def_KijeloltHatterSzin)
m_KijeloltBetuSzin = PropBag.ReadProperty("KijeloltBetuSzin", m_def_KijeloltBetuSzin)
m_ListaHatterSzin = PropBag.ReadProperty("ListaHatterSzin", m_def_ListaHatterSzin)
m_ListaBetuSzin = PropBag.ReadProperty("ListaBetuSzin", m_def_ListaBetuSzin)
m_GorgetesiSebesseg = PropBag.ReadProperty("GorgetesiSebesseg", m_def_GorgetesiSebesseg)
m_Rendez = PropBag.ReadProperty("Sorbarendezés", m_def_Rendez)
m_Betuk = PropBag.ReadProperty("Betûk", m_def_Betuk)

End Sub

Private Sub UserControl_Resize()

Call Resize

End Sub


Private Sub Index_Szamitas()

For b = 0 To visible_items - 1 Step 1
  If selector.Top = Label1(b).Top Then
     indexs = b
  End If
Next b

End Sub

Private Sub UserControl_Show()

Call Show
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'elmenti...
Call PropBag.WriteProperty("KijeloltHatterSzin", m_KijeloltHatterSzin, m_def_KijeloltHatterSzin)
Call PropBag.WriteProperty("KijeloltBetuSzin", m_KijeloltBetuSzin, m_def_KijeloltBetuSzin)
Call PropBag.WriteProperty("ListaHatterSzin", m_ListaHatterSzin, m_def_ListaHatterSzin)
Call PropBag.WriteProperty("ListaBetuSzin", m_ListaBetuSzin, m_def_ListaBetuSzin)
Call PropBag.WriteProperty("GorgetesiSebesseg", m_GorgetesiSebesseg, m_def_GorgetesiSebesseg)
Call PropBag.WriteProperty("Sorbarendezés", m_Rendez, m_def_Rendez)
Call PropBag.WriteProperty("Betûk", m_Betuk, m_def_Betuk)

End Sub

Private Sub Refresh_All()

For X = 0 To (visible_items - 1)
   With Label1(X)
      If X = 0 Then
         Label1(0).BackColor = m_KijeloltHatterSzin
      Else
         .BackColor = m_ListaHatterSzin
         .ForeColor = m_ListaBetuSzin
      End If
   End With
Next X

hatter.BackColor = m_ListaHatterSzin

End Sub



Private Sub Csuszka_Mozog()

kozep.Top = fel_gomb.Height + Int(szelekted) * ((scr_hat.Height - (kozep.Height)) / last_item)

End Sub

Private Sub Elemek_Csuszkahoz()

'szelekted számitás->visszafejtés
Dim s As Integer
     
If entry_point = 7 Then GoTo oda:
     
szelekted_2 = Int((kozep.Top - fel_gomb.Height) / ((scr_hat.Height - kozep.Height) / last_item))
If szelekted_2 > last_item Then szelekted_2 = last_item
  
  '(1)'
  If (szelekted_2 > visible_items) And (szelekted_2 < last_item - visible_items - 1) Then
         also = szelekted_2
         felso = also - visible_items
  '(2)'
  ElseIf (a_szelekted <= visible_items - 1) And _
         (szelekted_2 <= visible_items - (visible_items - indexs)) And _
         (szelekted_2 >= -visible_items - (indexs)) Then
     Exit Sub
  '(3)'
  ElseIf szelekted_2 >= last_item - visible_items Then ') And (szelekted_2 < last_item) Then
         also = last_item
         felso = also - (visible_items - 1)
  '(4)'
  ElseIf (a_szelekted >= last_item - visible_items - 1) And _
         (szelekted_2 >= last_item - visible_items - 1) And _
         (szelekted_2 <= last_item) Then
     Exit Sub
  '(5)'
  Else
     also = visible_items
     felso = 0
  End If

       For d = 0 To visible_items - 1
           Label1(d).Caption = nevek(inde_x(felso + d))
       Next d
    Refresh
  

oda:
Call Index_Szamitas

If (szelekted_2 > a_szelekted) And _
   (szelekted_2 - (visible_items) <= a_szelekted) Then

   For a = 0 To visible_items - 1
      If Label1(a).Caption = nevek(inde_x(a_szelekted)) Then
             Label1(a).BackColor = m_KijeloltHatterSzin
             Label1(a).ForeColor = m_KijeloltBetuSzin
             selector.Top = Label1(a).Top
             selector.Visible = True
      Else
         Label1(a).BackColor = m_ListaHatterSzin
         Label1(a).ForeColor = m_ListaBetuSzin
      End If
   Next a
   
Else
    selector.Visible = False
    Label1(indexs).BackColor = m_ListaHatterSzin
    Label1(indexs).ForeColor = m_ListaBetuSzin
End If

old_index = indexs

End Sub


Private Function Other_Keys(mit)

Dim s As Integer

'index számitás...
Call Index_Szamitas

Select Case mit

''''''''''''''''''''''''''''''''''''
Case "PAGE_UP": ''''''''''''''''''''
   szelekted = szelekted - visible_items + 1
   indexs = 0
   
   If szelekted < 0 Then szelekted = 0
    
   For a = 0 To visible_items - 1
       Label1(a).Caption = nevek(inde_x(szelekted + a))
   Next a
'(1)
Refresh

   Call Csuszka_Mozog
   
     Label1(indexs).BackColor = m_KijeloltHatterSzin
     Label1(indexs).ForeColor = m_KijeloltBetuSzin

If old_index = indexs Then GoTo at1:
     Label1(old_index).BackColor = m_ListaHatterSzin
     Label1(old_index).ForeColor = m_ListaBetuSzin
     old_index = indexs

    If SZ_BE = True Then
       selector.Move Label1(indexs).Left, Label1(indexs).Top, Label1(indexs).Width, Label1(indexs).Height
       selector.Visible = True
       First = 0
    End If
at1:
'(2)
Refresh
''''''''''''''''''''''''''''''''''''
Case "PAGE_DOWN":  '''''''''''''''''
  szelekted = szelekted + visible_items - 1
  indexs = visible_items - 1
  
  If szelekted >= last_item Then szelekted = last_item  ': MsgBox szelekted: Exit Function
   
   For a = 0 To visible_items - 1
       Label1(a).Caption = nevek(inde_x(szelekted - visible_items + a + 1))
   Next a
'(1)
Refresh
   
   Call Csuszka_Mozog
    
     Label1(indexs).BackColor = m_KijeloltHatterSzin
     Label1(indexs).ForeColor = m_KijeloltBetuSzin
    
If old_index = indexs Then GoTo at2:
     Label1(old_index).BackColor = m_ListaHatterSzin
     Label1(old_index).ForeColor = m_ListaBetuSzin
     old_index = indexs
    
    If SZ_BE = True Then
       selector.Move Label1(indexs).Left, Label1(indexs).Top, Label1(indexs).Width, Label1(indexs).Height
       selector.Visible = True
       First = 0
    End If
at2:
'(2)
Refresh

''''''''''''''''''''''''''''''''''''
Case "HOM_E": ''''''''''''''''''''''
   szelekted = 0
   indexs = 0
   
   For a = 0 To visible_items - 1
       Label1(a).Caption = nevek(inde_x(a))
   Next a
    
   Call Csuszka_Mozog
   
    If old_index = indexs Then Exit Function
     Label1(indexs).BackColor = m_KijeloltHatterSzin
     Label1(indexs).ForeColor = m_KijeloltBetuSzin
    
     Label1(old_index).BackColor = m_ListaHatterSzin
     Label1(old_index).ForeColor = m_ListaBetuSzin
     old_index = indexs

    If SZ_BE = True Then
       selector.Top = Label1(indexs).Top
       selector.Visible = True
    End If
''''''''''''''''''''''''''''''''''''
Case "EN_D": '''''''''''''''''''''''
   szelekted = last_item '- 1
   indexs = visible_items - 1
    
   For a = 0 To visible_items - 1
       Label1(a).Caption = nevek(inde_x(last_item - visible_items + a + 1))
   Next a
    
   Call Csuszka_Mozog

    If old_index = indexs Then Exit Function
     Label1(indexs).BackColor = m_KijeloltHatterSzin
     Label1(indexs).ForeColor = m_KijeloltBetuSzin
    
     Label1(old_index).BackColor = m_ListaHatterSzin
     Label1(old_index).ForeColor = m_ListaBetuSzin
     old_index = indexs

    If SZ_BE = True Then
       selector.Top = Label1(indexs).Top
       selector.Visible = True
    End If
''''''''''''''''''''''''''''''''''''
End Select

If First = 0 Then a_szelekted = szelekted: First = 1

End Function



Private Property Get Visibleindex() As Byte
Call Index_Szamitas

Visibleindex = indexs
End Property

Public Sub AddItem(name)
    Call Load_Item(name)
End Sub

Public Sub Rendezd()    'IndexVektoros rendezés..

Dim i, j, mi_n, s                      As Integer
 ' h,i,j,k,l,m,n
 ' Felveszi egy tombbe az értékeket elöször..
For h = 0 To last_item
    nevek(h) = tomb(h)
    inde_x(h) = h
Next h


If m_Rendez = 0 Then
' Majd rendezi azokat ABC-ben
  For i = 0 To last_item '- 1
   mi_n = i
       For j = i + 1 To last_item '- 1
        If nevek(inde_x(j)) < nevek(inde_x(mi_n)) Then
           mi_n = j
        End If
       Next j
        If i <> mi_n Then
           s = inde_x(i)
           inde_x(i) = inde_x(mi_n)
           inde_x(mi_n) = s
        End If
  Next i

Else
' Majd rendezi azokat_CBA-ban
  For i = 0 To last_item
   ma_x = i
      For j = i + 1 To last_item
       If nevek(inde_x(j)) > nevek(inde_x(ma_x)) Then
          ma_x = j
       End If
      Next j
       If i <> ma_x Then
          s = inde_x(i)
          inde_x(i) = inde_x(ma_x)
          inde_x(ma_x) = s
       End If
  Next i
End If

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' Majd ha kész és valamibõl több is van, akkor'''''''
' csak egyet hagy nekünk bent....''''''''''''''''''''
' Redundancia felszámolása...''''''''''''''''''''''''
Dim k As Long

For l = 1 To last_item
   If nevek(inde_x(l)) <> nevek(inde_x(l - 1)) Then
      k = k + 1
      nevek2(k) = nevek(inde_x(l))
      last_item = k
   End If
Next l

For m = 1 To last_item
    nevek(inde_x(m)) = nevek2(m)
Next m

'''''''''''''''''''''''''''''''''''
' És...kirakja a rendezett neveket'
For n = 0 To visible_items - 1
    Label1(n).Caption = nevek(inde_x(n))
Next n

Erase tomb
Erase nevek2

End Sub

Public Sub Elemek_Be()

On Error Resume Next

For bb = 0 To visible_items - 1
  If bb > 0 Then Load Label1(bb)
  
    With Label1(bb)
        .Top = Label1(bb - 1).Top + Label1(bb - 1).Height
        .Visible = True
    End With
Next bb

End Sub

Private Sub Kivalaszt(melyik As String)
    
Select Case melyik
  Case "0":
    Label1(indexs).BackColor = m_KijeloltHatterSzin
    Label1(indexs).ForeColor = m_KijeloltBetuSzin
    
    Label1(old_index).BackColor = m_ListaHatterSzin
    Label1(old_index).ForeColor = m_ListaBetuSzin
    old_index = indexs
  Case "1":
    Label1(indexs + 1).BackColor = m_KijeloltHatterSzin
    Label1(indexs + 1).ForeColor = m_KijeloltBetuSzin
    
    Label1(old_index).BackColor = m_ListaHatterSzin
    Label1(old_index).ForeColor = m_ListaBetuSzin
    old_index = indexs + 1
End Select
    
End Sub

Public Sub Clear_Items()

End Sub

Private Sub Show()

Call Refresh_All

'utolso x elem törlése...
If last_item < visible_items Then
   For a = visible_items - 1 To last_item + 1 Step -1
       Label1(a).Visible = False
   Next a
   kozep.Enabled = False
   fel_gomb.Enabled = False
   le_gomb.Enabled = False
Else
   kozep.Enabled = True
   fel_gomb.Enabled = True
   le_gomb.Enabled = True
End If

End Sub

Private Sub Resize()

If Width <= 2000 Then Width = 2000
If Height <= 1500 Then Height = 1500

hatter.Move 0, 0, ScaleWidth - 200, ScaleHeight

Call Elemek_Be
Height = hatter.Height

Call Refresh_All
szelekted = 0: a_szelekted = 0

End Sub

Public Sub Refresh_It()

Call Resize
Call Show
Call Refresh_All
Call Rendezd

Height = hatter.Height              ' visszaméretezés...

End Sub
