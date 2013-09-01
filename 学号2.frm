VERSION 5.00
Begin VB.Form Ñ§ºÅ2 
   BorderStyle     =   0  'None
   Caption         =   "³éÑ§ºÅ by òÇò¾"
   ClientHeight    =   5448
   ClientLeft      =   120
   ClientTop       =   744
   ClientWidth     =   10788
   Icon            =   "Ñ§ºÅ2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5448
   ScaleWidth      =   10788
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer3 
      Left            =   480
      Top             =   4080
   End
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "µ¹¼ÆÊ±5Ãë"
      Height          =   1092
      Left            =   6000
      TabIndex        =   3
      Top             =   4080
      Width           =   3252
   End
   Begin VB.Timer Timer 
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "¿ØÖÆ"
      Height          =   1095
      Left            =   1920
      TabIndex        =   0
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   252
      Left            =   960
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Label dao 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   4920
      TabIndex        =   4
      Top             =   4320
      Width           =   612
   End
   Begin VB.Label numbersname 
      Height          =   3495
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   5655
   End
   Begin VB.Label number 
      Height          =   3255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Menu xingming 
      Caption         =   "Òş²ØĞÕÃû"
   End
   Begin VB.Menu oneperson 
      Caption         =   "1¸öÈË"
   End
   Begin VB.Menu fiveperson 
      Caption         =   "5¸öÈË"
   End
   Begin VB.Menu about 
      Caption         =   "¹ØÓÚ"
   End
End
Attribute VB_Name = "Ñ§ºÅ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dao.Visible = False
number = "6"
Label1 = "5"
Timer.Interval = 1
Timer2.Interval = 1000
Timer3.Interval = 1000
Timer2.Enabled = False
Timer3.Enabled = False
Timer.Enabled = True
number.Font = "ËÎÌå"
number.FontSize = 150
number.ForeColor = vbRed
numbersname.Font = "ËÎÌå"
numbersname.FontSize = 90
numbersname.ForeColor = vbRed
Command1.Caption = "Start"
End Sub
Private Sub Command2_Click()
dao.Visible = True
numbersname.Visible = False
Timer.Enabled = True
Timer3.Enabled = True
If Timer2.Enabled = True Then
   Timer2.Enabled = False
Else
    Timer2.Enabled = True
End If


If Timer2 = True Then
    Command1.Enabled = False
    Command2.Enabled = False
End If

End Sub

Private Sub Command1_Click()
Command1.Caption = "Start"
If Timer.Enabled = True Then
Timer.Enabled = False
numbersname.Visible = True
Else
Command1.Caption = "Stop"
Timer.Enabled = True
numbersname.Visible = False
End If


Select Case number
Case 1
numbersname.Caption = "²ÌÂúÓ¨"
Case 2
numbersname.Caption = "³²¹ã´È"
Case 3
numbersname.Caption = "³Â°Ø¾ù"
Case 4
numbersname.Caption = "³Âºè±ò"
Case 5
numbersname.Caption = "³ÂË¼ê¿"
Case 6
numbersname.Caption = "³ÂÖ¾Òµ"
Case 7
numbersname.Caption = "µËÏ£Ñî"
Case 8
numbersname.Caption = "µËÓ¾ÑÔ"
Case 9
numbersname.Caption = "¶ÅÛÚÁÖ"
Case 10
numbersname.Caption = "¹ùĞ¡á°"
Case 11
numbersname.Caption = "ºÎ³©"
Case 12
numbersname.Caption = "ºÎî£ì÷"
Case 13
numbersname.Caption = "ºÎêÀáÉ"
Case 14
numbersname.Caption = "»Æ°²Äİ"
Case 15
numbersname.Caption = "»Æºê´ï"
Case 16
numbersname.Caption = "»ÆÑÅÓ±"
Case 17
numbersname.Caption = "Àµº£ö©"
Case 18
numbersname.Caption = "ÀèÂ¶Òğ"
Case 19
numbersname.Caption = "ÀîÖÇÅô"
Case 20
numbersname.Caption = "Áº¾¸êÍ"
Case 21
numbersname.Caption = "ÁºÊ«ä¿"
Case 22
numbersname.Caption = "ÁºÕòÃú"
Case 23
numbersname.Caption = "ÁÎ³şæº"
Case 24
numbersname.Caption = "ÁÎÍñÓ¨"
Case 25
numbersname.Caption = "ÁÎĞÀåû"
Case 26
numbersname.Caption = "ÁÎÑåçü"
Case 27
numbersname.Caption = "ÁõÊ«ÔÏ"
Case 28
numbersname.Caption = "Áõè÷êÍ"
Case 29
numbersname.Caption = "ÂÀÌì"
Case 30
numbersname.Caption = "ÂŞÓÀä¿"
Case 31
numbersname.Caption = "ÂŞè÷·Æ"
Case 32
numbersname.Caption = "Âó»ÛæÃ"
Case 33
numbersname.Caption = "Äª¼Ã¼ü"
Case 34
numbersname.Caption = "ÄªóãêÆ"
Case 35
numbersname.Caption = "Äª×Ó³¯"
Case 36
numbersname.Caption = "Å·ºº½Ü"
Case 37
numbersname.Caption = "Åíç²Çç"
Case 38
numbersname.Caption = "ÇÇÎ°ÁÁ"
Case 39
numbersname.Caption = "ÇğÓ±ÁØ"
Case 40
numbersname.Caption = "ÇñĞÇÏè"
Case 41
numbersname.Caption = "ÉÛº®ÑÅ"
Case 42
numbersname.Caption = "ËÕ´ïÖÎ"
Case 43
numbersname.Caption = "Ì¸Ğ¡Ìğ"
Case 44
numbersname.Caption = "ñûĞ¾ÔÏ"
Case 45
numbersname.Caption = "Ì·Ó±Üõ"
Case 46
numbersname.Caption = "Íõ¼Îç÷"
Case 47
numbersname.Caption = "ÎÂ¼ÓÓê"
Case 48
numbersname.Caption = "ÚùÖ¾èº"
Case 49
numbersname.Caption = "ÏÄ¿£ÌÎ"
Case 50
numbersname.Caption = "ÏÄÒãÃ÷"
Case 51
numbersname.Caption = "Ğì¼ÎÔÃ"
Case 52
numbersname.Caption = "ÑîÔÁÄş"
Case 53
numbersname.Caption = "ÓàĞÀÂû"
Case 54
numbersname.Caption = "ÓàÒã»ª"
Case 55
numbersname.Caption = "ÔøµÂê»"
Case 56
numbersname.Caption = "Ôøºº¶«"
Case 57
numbersname.Caption = "ÔøÓî"
Case 58
numbersname.Caption = "Ôøè÷ú"
Case 59
numbersname.Caption = "ÕÅÀöÇå"
Case 60
numbersname.Caption = "ÕÅÓ±Éº"
Case 61
numbersname.Caption = "ÕÔ¼Îâù"
Case 62
numbersname.Caption = "Ö£ºÆÈ»"
Case 63
numbersname.Caption = "ÖÓ»Û¾ı"
Case 64
numbersname.Caption = "ÖÓÃ÷Îµ"
End Select

End Sub
Private Sub Timer_Timer()
number = number + 1

If number = 65 Then
number = "1"
End If

End Sub
Private Sub Timer2_Timer()
dao = dao - 1

If dao = 0 Then
numbersname.Visible = True
Command1.Enabled = True
Command2.Enabled = True
Timer2.Enabled = False
Timer.Enabled = False
dao = "5"
dao.Visible = False
End If

End Sub

Private Sub Timer3_Timer()
Label1 = Label1 - 1
If Label1 = 0 Then
Timer3.Enabled = False
Label1 = "5"

Select Case number
Case 1
numbersname.Caption = "²ÌÂúÓ¨"
Case 2
numbersname.Caption = "³²¹ã´È"
Case 3
numbersname.Caption = "³Â°Ø¾ù"
Case 4
numbersname.Caption = "³Âºè±ò"
Case 5
numbersname.Caption = "³ÂË¼ê¿"
Case 6
numbersname.Caption = "³ÂÖ¾Òµ"
Case 7
numbersname.Caption = "µËÏ£Ñî"
Case 8
numbersname.Caption = "µËÓ¾ÑÔ"
Case 9
numbersname.Caption = "¶ÅÛÚÁÖ"
Case 10
numbersname.Caption = "¹ùĞ¡á°"
Case 11
numbersname.Caption = "ºÎ³©"
Case 12
numbersname.Caption = "ºÎî£ì÷"
Case 13
numbersname.Caption = "ºÎêÀáÉ"
Case 14
numbersname.Caption = "»Æ°²Äİ"
Case 15
numbersname.Caption = "»Æºê´ï"
Case 16
numbersname.Caption = "»ÆÑÅÓ±"
Case 17
numbersname.Caption = "Àµº£ö©"
Case 18
numbersname.Caption = "ÀèÂ¶Òğ"
Case 19
numbersname.Caption = "ÀîÖÇÅô"
Case 20
numbersname.Caption = "Áº¾¸êÍ"
Case 21
numbersname.Caption = "ÁºÊ«ä¿"
Case 22
numbersname.Caption = "ÁºÕòÃú"
Case 23
numbersname.Caption = "ÁÎ³şæº"
Case 24
numbersname.Caption = "ÁÎÍñÓ¨"
Case 25
numbersname.Caption = "ÁÎĞÀåû"
Case 26
numbersname.Caption = "ÁÎÑåçü"
Case 27
numbersname.Caption = "ÁõÊ«ÔÏ"
Case 28
numbersname.Caption = "Áõè÷êÍ"
Case 29
numbersname.Caption = "ÂÀÌì"
Case 30
numbersname.Caption = "ÂŞÓÀä¿"
Case 31
numbersname.Caption = "ÂŞè÷·Æ"
Case 32
numbersname.Caption = "Âó»ÛæÃ"
Case 33
numbersname.Caption = "Äª¼Ã¼ü"
Case 34
numbersname.Caption = "ÄªóãêÆ"
Case 35
numbersname.Caption = "Äª×Ó³¯"
Case 36
numbersname.Caption = "Å·ºº½Ü"
Case 37
numbersname.Caption = "Åíç²Çç"
Case 38
numbersname.Caption = "ÇÇÎ°ÁÁ"
Case 39
numbersname.Caption = "ÇğÓ±ÁØ"
Case 40
numbersname.Caption = "ÇñĞÇÏè"
Case 41
numbersname.Caption = "ÉÛº®ÑÅ"
Case 42
numbersname.Caption = "ËÕ´ïÖ¾"
Case 43
numbersname.Caption = "Ì¸Ğ¡Ìğ"
Case 44
numbersname.Caption = "ñûĞ¾ÔÏ"
Case 45
numbersname.Caption = "Ì·Ó±Üõ"
Case 46
numbersname.Caption = "Íõ¼Îç÷"
Case 47
numbersname.Caption = "ÎÂ¼ÒÓê"
Case 48
numbersname.Caption = "ÚùÖ¾èº"
Case 49
numbersname.Caption = "ÏÄ¿£ÌÎ"
Case 50
numbersname.Caption = "ÏÄÒãÃ÷"
Case 51
numbersname.Caption = "Ğì¼ÎÔÃ"
Case 52
numbersname.Caption = "ÑîÔÁÄş"
Case 53
numbersname.Caption = "ÓàĞÀÂû"
Case 54
numbersname.Caption = "ÓàÒã»ª"
Case 55
numbersname.Caption = "ÔøµÂê»"
Case 56
numbersname.Caption = "Ôøºº¶«"
Case 57
numbersname.Caption = "ÔøÓî"
Case 58
numbersname.Caption = "Ôøè÷ú"
Case 59
numbersname.Caption = "ÕÅÀöÇå"
Case 60
numbersname.Caption = "ÕÅÓ±Éº"
Case 61
numbersname.Caption = "ÕÔ¼Îâù"
Case 62
numbersname.Caption = "Ö£ºÆÈ»"
Case 63
numbersname.Caption = "ÖÓ»Û¾ı"
Case 64
numbersname.Caption = "ÖÓÃ÷Îµ"
End Select

End If
End Sub
Private Sub xingming_Click()
Ñ§ºÅ1.Show
Ñ§ºÅ2.Hide
End Sub
Private Sub about_Click()
frmabout.Show
End Sub
Private Sub oneperson_Click()
Ñ§ºÅ1.Show
Ñ§ºÅ2.Hide
End Sub
Private Sub fiveperson_Click()
Ñ§ºÅ2.Hide
Ñ§ºÅ3.Show
End Sub
