VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C0C0&
   Caption         =   "Расчет среднего отклонения(v1.3.1)"
   ClientHeight    =   10200
   ClientLeft      =   2010
   ClientTop       =   3030
   ClientWidth     =   15795
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   15795
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   975
      Left            =   7560
      TabIndex        =   65
      Top             =   9000
      Width           =   3375
      Begin VB.CheckBox Check1 
         BackColor       =   &H000080FF&
         Caption         =   "Минус 0,3 в корректуру прицела"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   66
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.CommandButton btnIzmOtkl 
      BackColor       =   &H00FF00FF&
      Caption         =   "Ввести Изменения"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   11100
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   8280
      Width           =   4000
   End
   Begin VB.TextBox pDDovIzm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14160
      TabIndex        =   63
      Text            =   "0"
      Top             =   7450
      Width           =   700
   End
   Begin VB.TextBox pDPrIzm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11760
      TabIndex        =   61
      Text            =   "0"
      Top             =   7450
      Width           =   700
   End
   Begin VB.TextBox poleTimer 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   6000
      TabIndex        =   58
      Text            =   "0"
      Top             =   9240
      Width           =   1100
   End
   Begin VB.TextBox polePoletnoe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2040
      TabIndex        =   57
      Text            =   "0"
      Top             =   9200
      Width           =   1000
   End
   Begin VB.Timer Timer1 
      Left            =   15000
      Top             =   9480
   End
   Begin VB.CommandButton btnVustrel 
      BackColor       =   &H0000FF00&
      Caption         =   "ВЫСТРЕЛ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4000
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   9150
      Width           =   1700
   End
   Begin VB.CommandButton btnSprPoPoleXc 
      BackColor       =   &H0000FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   8040
      Width           =   500
   End
   Begin VB.CommandButton spravkaPoDXr 
      BackColor       =   &H0000FFFF&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   100
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2630
      Width           =   500
   End
   Begin VB.TextBox pYr 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2100
      TabIndex        =   52
      Text            =   "0"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.TextBox pXr 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   100
      TabIndex        =   51
      Text            =   "0"
      Top             =   1900
      Width           =   1500
   End
   Begin VB.TextBox poleKomanda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   11400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   48
      Text            =   "Form1.frx":0000
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox pKorrYgl 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5000
      TabIndex        =   45
      Text            =   "0"
      Top             =   3700
      Width           =   1200
   End
   Begin VB.TextBox pKorrPr 
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5000
      TabIndex        =   44
      Text            =   "0"
      Top             =   2700
      Width           =   1200
   End
   Begin VB.CommandButton btnPokazArhiv 
      BackColor       =   &H0000C000&
      Caption         =   "Показать Архив"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3300
      Width           =   2200
   End
   Begin VB.ComboBox ktoRabotaet 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   9000
      TabIndex        =   40
      Top             =   1200
      Width           =   2200
   End
   Begin VB.ComboBox poleNZeli 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6500
      TabIndex        =   39
      Top             =   1200
      Width           =   2000
   End
   Begin VB.TextBox pSootnsh 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4000
      TabIndex        =   36
      Text            =   "0"
      Top             =   8200
      Width           =   1500
   End
   Begin VB.TextBox pYc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   34
      Text            =   "0"
      Top             =   8200
      Width           =   1500
   End
   Begin VB.TextBox pXc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   33
      Text            =   "0"
      Top             =   8200
      Width           =   1500
   End
   Begin VB.TextBox pdPr 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4000
      TabIndex        =   30
      Text            =   "0"
      Top             =   6800
      Width           =   1095
   End
   Begin VB.TextBox pPlus 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2200
      TabIndex        =   28
      Text            =   "0"
      Top             =   8200
      Width           =   1000
   End
   Begin VB.TextBox pMinus 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   800
      TabIndex        =   27
      Text            =   "0"
      Top             =   8200
      Width           =   1000
   End
   Begin VB.TextBox pdDov 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1400
      TabIndex        =   23
      Text            =   "0"
      Top             =   6000
      Width           =   1000
   End
   Begin VB.TextBox pVd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   21
      Text            =   "0"
      Top             =   6800
      Width           =   1500
   End
   Begin VB.TextBox pdXtus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   20
      Text            =   "0"
      Top             =   6800
      Width           =   1500
   End
   Begin VB.TextBox pYop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   17
      Text            =   "0"
      Top             =   5300
      Width           =   1500
   End
   Begin VB.TextBox pXop 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   16
      Text            =   "0"
      Top             =   5300
      Width           =   1500
   End
   Begin VB.TextBox pvSrdY 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   2200
      TabIndex        =   13
      Text            =   "0"
      Top             =   4000
      Width           =   1500
   End
   Begin VB.TextBox pvSrDx 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   100
      TabIndex        =   10
      Text            =   "0"
      Top             =   4000
      Width           =   1500
   End
   Begin VB.TextBox pvNRazr 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   20.25
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   4000
      TabIndex        =   8
      Text            =   "0"
      Top             =   1200
      Width           =   1000
   End
   Begin VB.TextBox pdY 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2100
      TabIndex        =   7
      Text            =   "0"
      Top             =   800
      Width           =   1500
   End
   Begin VB.TextBox pdX 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   100
      TabIndex        =   6
      Text            =   "0"
      Top             =   800
      Width           =   1500
   End
   Begin VB.CommandButton clickOchistka 
      BackColor       =   &H008080FF&
      Caption         =   "ОЧИСТИТЬ"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2100
      Width           =   2200
   End
   Begin VB.CommandButton clickReshSredn 
      BackColor       =   &H00FF8080&
      Caption         =   "РЕШИТЬ"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6500
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3300
      Width           =   2000
   End
   Begin VB.CommandButton clicDobavRazriv 
      BackColor       =   &H00FF8080&
      Caption         =   "Добавить разрыв"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   6500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2100
      Width           =   2000
   End
   Begin VB.Label Label30 
      BackColor       =   &H0000C0C0&
      Caption         =   "Дов"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   13080
      TabIndex        =   62
      Top             =   7500
      Width           =   700
   End
   Begin VB.Label Label29 
      BackColor       =   &H0000C0C0&
      Caption         =   "Пр"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   60
      Top             =   7500
      Width           =   495
   End
   Begin VB.Label Label28 
      BackColor       =   &H0000C0C0&
      Caption         =   "Изменения отклонений с учетом корректуры"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11760
      TabIndex        =   59
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label Label27 
      BackColor       =   &H0000C0C0&
      Caption         =   "Полетное"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   56
      Top             =   9300
      Width           =   1455
   End
   Begin VB.Label Label26 
      BackColor       =   &H0000C0C0&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2520
      TabIndex        =   50
      Top             =   1500
      Width           =   400
   End
   Begin VB.Label Label25 
      BackColor       =   &H0000C0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   480
      TabIndex        =   49
      Top             =   1500
      Width           =   400
   End
   Begin VB.Label Label24 
      BackColor       =   &H0000C0C0&
      Caption         =   "Угл"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   47
      Top             =   3700
      Width           =   495
   End
   Begin VB.Label Label23 
      BackColor       =   &H0000C0C0&
      Caption         =   "Пр/м"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   46
      Top             =   2700
      Width           =   700
   End
   Begin VB.Label Label22 
      BackColor       =   &H0000C0C0&
      Caption         =   "Корректура"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   400
      Left            =   4200
      TabIndex        =   43
      Top             =   2100
      Width           =   2000
   End
   Begin VB.Label Label21 
      BackColor       =   &H0000C0C0&
      Caption         =   "Корректура по серии"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1800
      TabIndex        =   42
      Top             =   4920
      Width           =   3300
   End
   Begin VB.Label Label20 
      BackColor       =   &H0000C0C0&
      Caption         =   "Кто работает"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   38
      Top             =   300
      Width           =   2200
   End
   Begin VB.Label Label19 
      BackColor       =   &H0000C0C0&
      Caption         =   "№ Цели"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   37
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label Label18 
      BackColor       =   &H0000C0C0&
      Caption         =   "Соотношение"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   4000
      TabIndex        =   35
      Top             =   7500
      Width           =   2200
   End
   Begin VB.Label Label17 
      BackColor       =   &H0000C0C0&
      Caption         =   "Уц"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   32
      Top             =   7500
      Width           =   855
   End
   Begin VB.Label Label16 
      BackColor       =   &H0000C0C0&
      Caption         =   "Хц"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   31
      Top             =   7500
      Width           =   855
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C0C0&
      Caption         =   "Корректура в прицел"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   3700
      TabIndex        =   29
      Top             =   6100
      Width           =   1815
   End
   Begin VB.Label Label14 
      BackColor       =   &H0000C0C0&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2500
      TabIndex        =   26
      Top             =   7500
      Width           =   400
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C0C0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   1100
      TabIndex        =   25
      Top             =   7500
      Width           =   400
   End
   Begin VB.Label Label12 
      BackColor       =   &H0000C0C0&
      Caption         =   "Соотношение знаков"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   800
      TabIndex        =   24
      Top             =   6800
      Width           =   2400
   End
   Begin VB.Label Label11 
      BackColor       =   &H0000C0C0&
      Caption         =   "Корректура в направление"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   100
      TabIndex        =   22
      Top             =   5500
      Width           =   4200
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000C0C0&
      Caption         =   "Вд"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   19
      Top             =   6000
      Width           =   800
   End
   Begin VB.Label Label9 
      BackColor       =   &H0000C0C0&
      Caption         =   "dXтыс"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   18
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000C0C0&
      Caption         =   "Уоп"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9000
      TabIndex        =   15
      Top             =   4600
      Width           =   800
   End
   Begin VB.Label Label7 
      BackColor       =   &H0000C0C0&
      Caption         =   "Хоп"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6500
      TabIndex        =   14
      Top             =   4600
      Width           =   800
   End
   Begin VB.Label Label6 
      BackColor       =   &H0000C0C0&
      Caption         =   "dY"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   2700
      TabIndex        =   12
      Top             =   3200
      Width           =   700
   End
   Begin VB.Label Label5 
      BackColor       =   &H0000C0C0&
      Caption         =   "dX"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   500
      TabIndex        =   11
      Top             =   3200
      Width           =   700
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000C0C0&
      Caption         =   "Средний"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   2520
      Width           =   2000
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000C0C0&
      Caption         =   "№ Разрыва"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   3500
      TabIndex        =   2
      Top             =   300
      Width           =   2000
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000C0C0&
      Caption         =   "dY"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2400
      TabIndex        =   1
      Top             =   300
      Width           =   800
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C0C0&
      Caption         =   "dX"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   15.75
         Charset         =   204
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   500
      TabIndex        =   0
      Top             =   300
      Width           =   800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public schetchikSerii As Integer
Public zapisRazrFail As String
Public srednRazrFail As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim epoch As Currency

Dim totalSeconds As Integer ' Общее число прошедших секунд
Dim poletnoe As Integer

'изменение отклонений с учетом корректуры
Private Sub btnIzmOtkl_Click()
Dim dPr As Single, dDov As Single
Dim srDX As Single, srDY As Single
Dim xOP As Double, yOP As Double, xc As Double, yc As Double
Dim dTopo As Double, ygolTopo As Double
Dim newDTopo As Double, newYgolTopo As Double
Dim dXTus As Integer, scetRazr As Integer

On Error GoTo HandleError

'получаем корректуры
dPr = pDPrIzm: dDov = pDDovIzm

'получаем средний с файла
Open srednRazrFail For Input As #1
Input #1, srDX, srDY
Close #1

'получаем координаты ОП и дельта икс тыс
xOP = pXop: yOP = pYop: xc = pXc: yc = pYc
dXTus = pdXtus: scetRazr = pvNRazr

'определяем среднее по разрывам
srDX = srDX / scetRazr: srDY = srDY / scetRazr

'получаем дальность топо и угол топо по среднему до изменений
proOGZ xc + srDX, yc + srDY, xOP, yOP, dTopo, ygolTopo

'изменяем дальность м угол на корректуру
newDTopo = dTopo + (dXTus * dPr)
newYgolTopo = ygolTopo + dDov

'получаем координаты нового среднего
Dim newXsr As Double, newYsr As Double
PGZ xOP, yOP, newYgolTopo, newDTopo, newXsr, newYsr

'определяем новый средний
srDX = (newXsr - xc) * scetRazr: srDY = (newYsr - yc) * scetRazr

'записываем в файл новый средний
Open srednRazrFail For Output As #1
Write #1, srDX, srDY
Close #1

pDPrIzm.Text = "0": pDDovIzm.Text = "0"
Exit Sub
HandleError:
Dim q As Integer
q = MsgBox("Ошибка в процедуре <btnIzmOtkl>")

End Sub

Private Sub btnPokazArhiv_Click()
Shell "C:\Windows\system32\notepad.exe" + " " + zapisRazrFail, vbNormalNoFocus
End Sub

Private Sub btnSprPoPoleXc_Click()
Dim q As Double

q = MsgBox("Вставлять из буфера: 1. Скопированные координаты с SAS.Планет. 2. Скопированное сообщение из 'Телеграмм' пересланное из 'Ветерок' по цели")
End Sub

Private Sub btnVustrel_Click()

  If Timer1.Enabled Then
        ' Если таймер включен, отключаем его и очищаем данные
        btnVustrel.Caption = "ВЫСТРЕЛ"
        btnVustrel.BackColor = &HFF00&
        poleTimer.BackColor = &H80FFFF
        Timer1.Enabled = False
        totalSeconds = 0
        poleTimer.Text = ""
    Else
        ' Иначе начинаем измерение времени
        btnVustrel.Caption = "СТОП"
        btnVustrel.BackColor = &HFF&
        Timer1.Interval = 1000     ' Интервал обновления таймера в миллисекундах
        Timer1.Enabled = True       ' Запускаем таймер
    End If
    
End Sub


Private Sub pDPrIzm_Click()
pDPrIzm.Text = ""
End Sub

Private Sub pDPrIzm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pDDovIzm.Text = ""
    pDDovIzm.SetFocus
    Else
End If
End Sub


Private Sub polePoletnoe_Click()
polePoletnoe.Text = ""
End Sub

Private Sub Timer1_Timer()
    Dim elapsedTime As Date
    
    poletnoe = polePoletnoe.Text
    
    totalSeconds = totalSeconds + 1           ' Увеличиваем счётчик секунд
    If totalSeconds > poletnoe Then poleTimer.BackColor = &HFF&
    poleTimer.Text = totalSeconds
    
End Sub

Private Sub clicDobavRazriv_Click()
Dim dX As Single, dY As Single, nRazriv As Integer, nPlus As Integer, nMinus As Integer
Dim dXsFail As Single, dYsFail As Single
Dim schetchik As Integer, vd As Integer
Dim x As Single, y As Single, Xr As Single, Yr As Single
Dim xOP As Single, yOP As Single

schetchik = 1
dX = pdX: dY = pdY: nRazriv = pvNRazr + schetchik
Xr = pXr: Yr = pYr
pvNRazr = nRazriv

'запись в архив
If FileLen(zapisRazrFail) = 0 Then
        Open zapisRazrFail For Append As #1
        Write #1, "Запись стрельбы по цели : " & poleNZeli
        Write #1, schetchikSerii & "-я серия"
        If dX = 0 And dY = 0 Then
            Write #1, "Р" & nRazriv & "  Xr = " & Xr & "  Yr = " & Yr & "   " & Now
            Else
            Write #1, "Р" & nRazriv & "  dX = " & dX & "  dY = " & dY & "   " & Now
        End If
        Close #1
    Else
        Open zapisRazrFail For Append As #1
        Write #1, "Р" & nRazriv & "  dX = " & dX & "  dY = " & dY & "   " & Now
        Close #1
End If

'считаем знаки
vd = pVd
x = pXc: y = pYc: xOP = pXop: yOP = pYop
If dX = 0 And dY = 0 Then
    dX = Xr - x: dY = Yr - y
    Else
    Xr = x + dX: Yr = y + dY
End If
proOGZ x, y, xOP, yOP, Dt, Ygt
dXTus = pdXtus
podRASCHETXY Xr, Yr, xOP, yOP, dXTus, Dt, Ygt, dD, dDov, dPr
nPlus = pPlus: nMinus = pMinus
If dX = 0 And dY = 0 And Xr = 0 And Yr = 0 Then
    nPlus = nPlus + 1: nMinus = nMinus + 1
    Else
    If dD < 0 And Abs(dD) > (vd / 2) Then
        nPlus = nPlus + 1
        ElseIf dD > 0 And Abs(dD) > (vd / 2) Then
        nMinus = nMinus + 1
        Else
        nPlus = nPlus + 1: nMinus = nMinus + 1
    End If
End If
pPlus = nPlus: pMinus = nMinus

Open srednRazrFail For Input As #1
Input #1, dXsFail, dYsFail
Close #1

dXsFail = dXsFail + dX: dYsFail = dYsFail + dY

Open srednRazrFail For Output As #1
Write #1, dXsFail, dYsFail
Close #1

''''korrektyra
Dim korDX As Integer, korDy As Integer

On Error GoTo ErrorHundler

korDX = dX: korrdY = dY

x = pXc: y = pYc: xOP = pXop: yOP = pYop
If dX = 0 And dY = 0 Then
    Else
    Xr = x + korDX: Yr = y + korrdY
End If
proOGZ x, y, xOP, yOP, Dt, Ygt

dXTus = pdXtus
podRASCHETXY Xr, Yr, xOP, yOP, dXTus, Dt, Ygt, dD, dDov, dPr

'уменшаем корректуру в прицел на 0,3
If Check1 = 1 Then dPr = Round(dPr * 0.75)

If dDov > 0 Then
    If dDov < 10 Then
        pKorrYgl = "+0-0" & dDov
        Else
        pKorrYgl = "+0-" & dDov
    End If
    Else
    If dDov > -10 Then
        pKorrYgl = "-0-0" & Abs(dDov)
        Else
        pKorrYgl = "-0-" & Abs(dDov)
    End If
End If

If dD > 0 Then
    pKorrPr = "+" & dPr & "/" & "+" & dD
    Else
    pKorrPr = dPr & "/" & dD
End If
''''''

pdX = 0: pdY = 0: pXr = 0: pYr = 0
Exit Sub

ErrorHundler:
Dim msq As Integer
msq = MsgBox("поверить правильность ввода отклонений!!!")
End Sub

Private Sub clickOchistka_Click()
Open srednRazrFail For Output As #1
Write #1, 0, 0
Close #1
'добавляем счетчик серии
schetchikSerii = schetchikSerii + 1

Open zapisRazrFail For Append As #1
Write #1, schetchikSerii & "-я серия"
Close #1

pvNRazr = 0: pvSrDx = 0: pvSrdY = 0: pdX = 0: pdY = 0: pPlus = 0: pMinus = 0: pdDov = 0

End Sub

Private Sub clickReshSredn_Click()
Dim nRazriv As Integer, srDX As Single, srDY As Single
Dim dXsFail As Single, dYsFail As Single
Dim x As Single, y As Single, xOP As Single, yOP As Single
Dim bolshee As Single, menshee As Single, plus As Single, minus As Single, sootnoshenie As Single, vd As Single, korP As Single

nRazriv = pvNRazr

Open srednRazrFail For Input As #1
Input #1, dXsFail, dYsFail
Close #1

srDX = Round(dXsFail / (nRazriv + 0.0001))
srDY = Round(dYsFail / (nRazriv + 0.0001))

pvSrDx = srDX: pvSrdY = srDY

x = pXc: y = pYc: xOP = pXop: yOP = pYop
Xr = x + srDX: Yr = y + srDY
proOGZ x, y, xOP, yOP, Dt, Ygt
dXTus = pdXtus: vd = pVd
podRASCHETXY Xr, Yr, xOP, yOP, dXTus, Dt, Ygt, dD, dDov, dPr

If dDov > 0 Then
    If dDov < 10 Then
        pdDov = "+0-0" & dDov
        Else
        pdDov = "+0-" & dDov
    End If
    Else
    If dDov > -10 Then
        pdDov = "-0-0" & Abs(dDov)
        Else
        pdDov = "-0-" & Abs(dDov)
    End If
End If

plus = pPlus: minus = pMinus
If plus > minus Then
    bolshee = plus: menshee = minus
    Else
    bolshee = minus: menshee = plus
End If
If menshee = 0 Then menshee = 1
sootnoshenie = Round(bolshee / (menshee + 0.001))
pSootnsh.Text = "1 / " + str(sootnoshenie)

If dPr < 0 Then
    pdPr = dPr
    Else
    pdPr = "+" & dPr
End If

Open zapisRazrFail For Append As #1
Write #1, "К-во разрывов = " & nRazriv & ", Средн : dX = " & srDX & " dY = " & srDY
Write #1, "Корректура в прицел = " & pdPr & ", Корректура в угломер = " & pdDov.Text
Write #1, "+" & plus & "  " & "-" & minus & ", Соотношение знаков = " & pSootnsh.Text
'добавить соотношение

Close #1

End Sub

Private Sub Form_Load()
'формируем уникальные названия файлов
zapisRazrFail = App.Path & "\" & "Zap" & Round((DateDiff("s", "1/1/1970", Date) + Timer) * 1000) & ".txt"
Sleep (1000)
srednRazrFail = App.Path & "\" & "Sred" & Round((DateDiff("s", "1/1/1970", Date) + Timer) * 1000) & ".txt"


Open srednRazrFail For Output As #1
Write #1, 0, 0
Close #1

pvNRazr = 0: pvSrDx = 0: pvSrdY = 0: pdX = 0: pdY = 0

Dim i As Integer
Dim t(1 To 10) As String
poleNZeli.Clear
941 Open "D:\YO_NA\zeli" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 942
 Input #1, t(1), t(2), t(3), t(4), t(5), t(6)
poleNZeli.AddItem t(1)
Loop
942 Close #1

ktoRabotaet.Clear
9410 Open "D:\YO_NA\raschSrednBP.txt" For Input As #1
Do While Not EOF(1)
If EOF(1) Then GoTo 9420
 Input #1, t(1), t(2), t(3), t(4)
ktoRabotaet.AddItem t(1)
Loop
9420 Close #1

Open zapisRazrFail For Output As #1
Close #1
'присваеваем значение счетчику серии
schetchikSerii = 1

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim nFail As String
'добавить считывание записи стрельб и перезаписи в файл
nFail = "\archivZelei\" & poleNZeli.Text & "_" & Round((DateDiff("s", "1/1/1970", Date) + Timer) * 1000) & ".txt"

Open App.Path & nFail For Output As #1
Close #1

'скопировать запись стрельбы в архивный файл
FileCopy zapisRazrFail, App.Path & nFail
'добавляем запись с поля команда
Open App.Path & nFail For Append As #1
Write #1, poleKomanda.Text
Close

End Sub

Private Sub pdX_Click()
pdX.Text = ""
End Sub

Private Sub pdX_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pdY.Text = ""
    pdY.SetFocus
    Else
    End If
End Sub
Sub podRASCHETXY(ByVal Xr As Single, ByVal Yr As Single, ByVal xOP As Single, ByVal yOP As Single, ByVal dXTus As Single, ByVal Dt As Single, ByVal Ygolt As Single, dD, dDov, dPr)
        
proOGZ Xr, Yr, xOP, yOP, Dtr, Ygoltr
dD = Round(Dt - Dtr)
 dDov = Round(Ygolt - Ygoltr)
 dPr = Round(dD / (dXTus + 0.001))

End Sub

Sub proOGZ(ByVal xc As Single, ByVal yc As Single, ByVal xOP As Single, ByVal yOP As Single, Dt, Ygt)
Dim dxc As Single, dyc As Single

dxc = xc - xOP
dyc = yc - yOP
 Dt = Sqr(dxc ^ 2 + dyc ^ 2)
 Ar = Abs(Atn(dyc / (dxc + 0.1)) / 3.141592 * 30) * 100
 If dxc > 0 And dyc > 0 Then Ygt = Int(Ar)
 If dxc < 0 And dyc > 0 Then Ygt = Int(3000 - Ar)
 If dxc < 0 And dyc < 0 Then Ygt = Int(3000 + Ar)
 If dxc > 0 And dyc < 0 Then Ygt = Int(6000 - Ar)

End Sub
Sub PGZ(ByVal x As Double, ByVal y As Double, ByVal a As Double, ByVal d As Double, xc, yc)
xc = Cos(a / 100 * 6 * 3.141592 / 180) * d + x
yc = Sin(a / 100 * 6 * 3.141592 / 180) * d + y
End Sub

Private Sub pdXtus_Click()
pdXtus.Text = ""
End Sub

Private Sub pdY_Click()
pdY.Text = ""
End Sub
Private Sub poleNZeli_Click()
Dim z(1 To 10) As String
Dim nz As String
Dim xc As Single, yc As Single, hc As Single
nz = poleNZeli
1011 Open "D:\YO_NA\zeli" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
   If z(1) = nz Then xc = z(2): yc = z(3): hc = z(4): GoTo 1012
        GoTo 101111
1012 Close #1
pXc = xc: pYc = yc
End Sub
Private Sub poleNZeli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim xc As Single, yc As Single, hc As Single
nz = poleNZeli
If KeyCode = 13 Then
1011    Open "D:\YO_NA\zeli" For Input As #1
101111  If EOF(1) Then GoTo 1012
        Input #1, z(1), z(2), z(3), z(4), z(5), z(6)
        If z(1) = nz Then xc = z(2): yc = z(3): hc = z(4): GoTo 1012
        GoTo 101111
1012    Close #1
        pXc = xc: pYc = yc
    Else
End If
End Sub
Private Sub ktoRabotaet_Click()
Dim z(1 To 10) As String
Dim nz As String
Dim x As Double, y As Double
Dim osnNap As Integer

nz = ktoRabotaet
1011 Open "D:\YO_NA\raschSrednBP.txt" For Input As #1
101111 If EOF(1) Then GoTo 1012
   Input #1, z(1), z(2), z(3), z(4)
   If z(1) = nz Then x = z(2): y = z(3): osnNap = z(4): GoTo 1012
        GoTo 101111
1012 Close #1
pXop = x: pYop = y
End Sub
Private Sub ktoRabotaet_KeyDown(KeyCode As Integer, Shift As Integer)
Dim z(1 To 10) As String
Dim nz As String
Dim x As Double, y As Double
Dim osnNap As Integer

nz = ktoRabotaet
If KeyCode = 13 Then
1011    Open "D:\YO_NA\raschSrednBP.txt" For Input As #1
101111  If EOF(1) Then GoTo 1012
        Input #1, z(1), z(2), z(3), z(4)
        If z(1) = nz Then x = z(2): y = z(3): hc = z(4): GoTo 1012
        GoTo 101111
1012    Close #1
        pXop = x: pYop = y
    Else
End If
End Sub

Private Sub pXc_Click()
pXc.Text = ""
End Sub

Private Sub pXc_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String
Dim arr() As String
Dim arr2() As String
Dim x As Double, y As Double
Dim dlinaMass As Integer
Dim nX As Integer, nY As Integer
Dim xy() As Variant

On Error GoTo ErrorHandler

If KeyCode = 13 Then
    str = pXc
    arr = Split(str, " ")
    xy = getXY(arr)
    x = CDbl(xy(0)): y = CDbl(xy(1))
    x = x Mod 100000: y = y Mod 100000
    pXc = x: pYc = y
Else
End If

Exit Sub
ErrorHandler:

End Sub

Private Sub pXop_Click()
pXop.Text = ""
End Sub

Private Sub pXop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pYop.Text = ""
    pYop.SetFocus
    Else
End If
End Sub

Private Sub pXr_KeyDown(KeyCode As Integer, Shift As Integer)
Dim str As String
Dim arr() As String
Dim arr2() As String
Dim x As Double, y As Double
Dim dlinaMass As Integer
Dim nX As Integer, nY As Integer
Dim xy() As Variant

On Error GoTo ErrorHandler

If KeyCode = 13 Then
    str = pXr.Text
    arr = Split(str, " ")
    xy = getXY(arr)
   
    x = CDbl(xy(0)): y = CDbl(xy(1))
    x = x Mod 100000: y = y Mod 100000
    pXr = x: pYr = y
Else
End If

Exit Sub
ErrorHandler:
End Sub
Function getXY(ByVal masStr As Variant) As Variant()
Dim i As Integer, k As Integer, f As Integer
Dim x As Variant, y As Variant

On Error GoTo Shlapa

For k = 0 To sizeMass(masStr)
    For i = 1 To Len(masStr(k))
        If Val(Mid$(masStr(k), i)) <> 0 Then
            x = x & Strings.Mid(masStr(k), i, 1)
            If Len(x) = 7 Then GoTo FindY
            Else
        End If
    Next i
    x = ""
Next k

FindY:
For f = k + 1 To sizeMass(masStr)
    For i = 1 To Len(masStr(f))
        If Val(Mid$(masStr(k), i)) <> 0 Then
            y = y & Strings.Mid(masStr(f), i, 1)
            If Len(y) = 7 Then GoTo InitializeMass
            Else
        End If
    Next i
    y = ""
Next f

InitializeMass:
Dim mass(2) As Variant
mass(0) = x: mass(1) = y
getXY = mass
Exit Function

Shlapa:
Dim q As Double
g = MsgBox("Шляпа с данными в функции getXY!!!")
End Function
Function sizeMass(mass As Variant) As Integer
    sizeMass = UBound(mass) - LBound(mass) + 1
End Function
Private Sub pYop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pdXtus.Text = ""
    pdXtus.SetFocus
    Else
End If
End Sub

Private Sub pdXtus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pVd.Text = ""
    pVd.SetFocus
    Else
End If
End Sub
Private Sub pVd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    pXc.Text = ""
    pXc.SetFocus
    Else
End If
End Sub
Private Sub pXc_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorHadl

If KeyAscii = 13 Then
    If pYc = 0 Then
        pYc.Text = ""
        pYc.SetFocus
        ElseIf pYc = "" Then
        pYc.SetFocus
        Else
    End If
    Else
End If

Exit Sub

ErrorHadl:
End Sub
Private Sub pXr_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHadl

If KeyAscii = 13 Then
    If pYr = 0 Then
        pYr.Text = ""
        pYr.SetFocus
        ElseIf pYr = "" Then
        pYr.SetFocus
        Else
    End If
    Else
End If

Exit Sub

ErrorHadl:
    
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub
Private Sub pXr_Click()
pXr.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.Text = ""
    Text2.SetFocus
    Else
End If

End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub
'для русского текста
Function RussianStringToURLEncode_New(ByVal txt As String) As String
    For i = 1 To Len(txt)
        L = Mid(txt, i, 1)
        Select Case AscW(L)
            Case Is > 4095: t = "%" & Hex(AscW(L) \ 64 \ 64 + 224) & "%" & Hex(AscW(L) \ 64) & "%" & Hex(8 * 16 + AscW(L) Mod 64)
            Case Is > 127: t = "%" & Hex(AscW(L) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(L) Mod 64)
            Case 32: t = "%20"
            Case Else: t = L
        End Select
        RussianStringToURLEncode_New = RussianStringToURLEncode_New & t
    Next
End Function

Private Sub spravkaPoDXr_Click()
Dim q As Double

q = MsgBox("Вставлять из буфера: 1. Скопированные координаты с SAS.Планет. 2. Скопированное сообщение из 'Телеграмм' пересланное из 'Ветерок' по разрыву")
End Sub
