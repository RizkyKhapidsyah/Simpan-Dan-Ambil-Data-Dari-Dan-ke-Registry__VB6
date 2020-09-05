VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simpan dan Ambil Data dari dan ke Registry"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   5760
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdDeleteSetting 
      Caption         =   "Delete Setting"
      Height          =   375
      Left            =   4080
      TabIndex        =   13
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CommandButton cmdGetSetting 
      Caption         =   "Get Setting"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveSetting 
      Caption         =   "Save Setting"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   4440
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   3360
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Option4"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   7
      Top             =   3000
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'SaveSetting(AppName As String, Section As String, Setting As String)

'GetSetting(AppName As String, Section As String, Key As String, Default As String)

'GetAllSettings(AppName As String, Section As String)

'Untuk melihat hasil dari contoh ini, jalankan Registry 'dari menu Start->Run ketikkan: Regedit  lalu tekan 'Enter. Buka folder/direktori di explorer Regedit 'yaitu: HKEY_CURRENT_USER\Software\VB and VBA Program 'Settings\prjRegedit, kemudian periksa subfolder "Form" 'dan "TestRegedit". Khusus untuk Section "TestRegedit", 'seluruh nilai yang diambil dapat Anda lihat di List1.

Private Sub cmdDeleteSetting_Click()
On Error Resume Next
   DeleteSetting App.Title, "TestRegedit"
   MsgBox "Berhasil dihapus!", vbInformation, "Hapus OK"
End Sub

Private Sub cmdGetSetting_Click()
   form_load
End Sub

Private Sub cmdSaveQuit_Click()
   SimpanForm1
   SimpanLetakForm
   End
End Sub

Private Sub cmdSaveSetting_Click()
   SimpanForm1
   SimpanLetakForm
   MsgBox "Berhasil disimpan!", vbInformation, _
          "Simpan OK"
End Sub

Private Sub form_load()
Dim AtasForm, KiriForm As Integer
On Error Resume Next
   AtasForm = Screen.Height / 2 - Me.Height / 2
   KiriForm = Screen.Width / 2 - Me.Width / 2
   Me.Left = GetSetting(App.Title, "Form", "Kiri", _
             KiriForm)
   Me.Top = GetSetting(App.Title, "Form", "Atas", _
            AtasForm)
   Me.Width = GetSetting(App.Title, "Form", "Lebar", _
       0)
   Me.Height = GetSetting(App.Title, "Form", _
               "Tinggi", 5000)
    
   Dim avntSettings As Variant
   Dim intX As Integer
   avntSettings = GetAllSettings(App.Title, _
                 "TestRegedit")
   List1.Clear
   For intX = 0 To UBound(avntSettings, 1)
      List1.AddItem avntSettings(intX, 1)
   Next intX
    
   Text1 = List1.List(0)
   Text2 = List1.List(1)
   Text3 = List1.List(2)
   Check1 = List1.List(3)
   Check2 = List1.List(4)
   Option1(0) = List1.List(5)
   Option1(1) = List1.List(6)
   Option2 = List1.List(7)
   Option3 = List1.List(8)
   
   Combo1.List(0) = GetSetting(App.Title, _
                    "TestRegedit", "Combo1(0)", "")
   Combo1.List(1) = GetSetting(App.Title, _
                    "TestRegedit", "Combo1(1)", "")
   Combo1.List(2) = GetSetting(App.Title, _
                    "TestRegedit", "Combo1(2)", "")
   Combo1.Text = Text3.Text
End Sub

Sub SimpanForm1()
   SaveSetting App.Title, "TestRegedit", "Text1", Text1
   SaveSetting App.Title, "TestRegedit", "Text2", Text2
   SaveSetting App.Title, "TestRegedit", "Text3", Combo1.Text
   SaveSetting App.Title, "TestRegedit", "Check1", Check1.Value
   SaveSetting App.Title, "TestRegedit", "Check2", Check2.Value
   SaveSetting App.Title, "TestRegedit", _
               "Option1(0)", Option1(0).Value
   SaveSetting App.Title, "TestRegedit", _
            "Option1(1)", Option1(1).Value
   SaveSetting App.Title, "TestRegedit", "Option2", _
               Option2.Value
   SaveSetting App.Title, "TestRegedit", "Option3", _
               Option3.Value
   
   If Combo1.List(0) = "" Then
      SaveSetting App.Title, "TestRegedit", _
                  "Combo1(0)", Combo1.Text
   ElseIf Combo1.List(0) = Combo1.Text Or _
          Combo1.List(1) = Combo1.Text Or _
          Combo1.List(2) = Combo1.Text Then
      SaveSetting App.Title, "TestRegedit", _
         "Combo1(0)", Combo1.List(0)
      SaveSetting App.Title, "TestRegedit", _
         "Combo1(1)", Combo1.List(1)
      SaveSetting App.Title, "TestRegedit", _
         "Combo1(2)", Combo1.List(2)
   Else
      SaveSetting App.Title, "TestRegedit", _
                  "Combo1(2)", Combo1.List(1)
      SaveSetting App.Title, "TestRegedit", _
                  "Combo1(1)", Combo1.List(0)
      SaveSetting App.Title, "TestRegedit", _
                  "Combo1(0)", Combo1.Text
   End If
End Sub

Sub SimpanLetakForm()
  If Me.WindowState <> vbMinimized Then
     SaveSetting App.Title, "Form", "Kiri", Me.Left
     SaveSetting App.Title, "Form", "Atas", Me.Top
     SaveSetting App.Title, "Form", "Lebar", Me.Width
     SaveSetting App.Title, "Form", "Tinggi", Me.Height
  End If
End Sub


