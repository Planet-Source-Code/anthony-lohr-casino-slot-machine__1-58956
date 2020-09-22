VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form JackPotSlots1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   Icon            =   "JackPotSlots1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "JackPotSlots1.frx":0ABA
   ScaleHeight     =   8790
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   50
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   2400
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1440
      Top             =   3360
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":3FE2E
            Key             =   "CloseUp"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":4299A
            Key             =   "CloseDown"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":4551A
            Key             =   "MinimizeUp"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":48002
            Key             =   "MinimizeDown"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":4AACE
            Key             =   "BetOneUp"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":4EEDE
            Key             =   "BetOneDown"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":53366
            Key             =   "MaxBetUp"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":5789A
            Key             =   "MaxBetDown"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":5BE0E
            Key             =   "SpinUp"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":614EA
            Key             =   "SpinDown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":66C06
            Key             =   "Bet1Normal"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":69C3E
            Key             =   "Bet2Normal"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":6CCEA
            Key             =   "Bet3Normal"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":6FDB6
            Key             =   "Bet4Normal"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":72E46
            Key             =   "Bet5Normal"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":75E92
            Key             =   "Bet6Normal"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":78F36
            Key             =   "Bet7Normal"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":7BB5E
            Key             =   "Bet8Normal"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":7E7FE
            Key             =   "Bet1Light"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":817DA
            Key             =   "Bet2Light"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":8482A
            Key             =   "Bet3Light"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":87896
            Key             =   "Bet4Light"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":8A8F2
            Key             =   "Bet5Light"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":8D8FE
            Key             =   "Bet6Light"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":9094E
            Key             =   "Bet7Light"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":9356E
            Key             =   "Bet8Light"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   69
      ImageHeight     =   89
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   34
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":961C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":987A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":9B0DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":9D612
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":9F982
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":A255E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":A4D8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":A753E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":AA3A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":ACBEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":AF662
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":B1812
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":B3DF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":B672A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":B8C5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":BAFCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":BDBAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":C03D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":C2B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":C59F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":C823A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":CACAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":CCE5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":CF43E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":D1D76
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":D42AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":D661A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":D91F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":DBA22
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":DE1D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":E103E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":E3886
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":E62FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "JackPotSlots1.frx":E84AA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   8
      Left            =   3450
      Tag             =   "0"
      Top             =   5690
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   5
      Left            =   2190
      Tag             =   "0"
      Top             =   5690
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   2
      Left            =   930
      Tag             =   "0"
      Top             =   5690
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   7
      Left            =   3450
      Tag             =   "0"
      Top             =   4130
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   4
      Left            =   2190
      Tag             =   "0"
      Top             =   4130
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   1
      Left            =   930
      Tag             =   "0"
      Top             =   4130
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   6
      Left            =   3450
      Tag             =   "0"
      Top             =   2560
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   3
      Left            =   2190
      Tag             =   "0"
      Top             =   2560
      Width           =   1035
   End
   Begin VB.Image Roll 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Index           =   0
      Left            =   930
      Tag             =   "0"
      Top             =   2560
      Width           =   1035
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   450
      TabIndex        =   8
      Tag             =   "0"
      Top             =   7210
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   450
      TabIndex        =   7
      Top             =   2020
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3870
      TabIndex        =   6
      Tag             =   "0"
      Top             =   1930
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2610
      TabIndex        =   5
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1330
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Tag             =   "0"
      Top             =   6180
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   4620
      Width           =   255
   End
   Begin VB.Label BetCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   350
      TabIndex        =   1
      Top             =   3040
      Width           =   255
   End
   Begin VB.Image Bet1 
      Height          =   570
      Index           =   8
      Left            =   250
      Picture         =   "JackPotSlots1.frx":EAA8A
      Tag             =   "0"
      Top             =   7090
      Width           =   585
   End
   Begin VB.Image Bet1 
      Height          =   585
      Index           =   7
      Left            =   250
      Picture         =   "JackPotSlots1.frx":ED718
      Tag             =   "0"
      Top             =   1900
      Width           =   570
   End
   Begin VB.Image Bet1 
      Height          =   690
      Index           =   6
      Left            =   3660
      Picture         =   "JackPotSlots1.frx":F0330
      Tag             =   "0"
      Top             =   1800
      Width           =   585
   End
   Begin VB.Image Bet1 
      Height          =   690
      Index           =   5
      Left            =   2400
      Picture         =   "JackPotSlots1.frx":F33C4
      Tag             =   "0"
      Top             =   1790
      Width           =   585
   End
   Begin VB.Image Bet1 
      Height          =   690
      Index           =   4
      Left            =   1140
      Picture         =   "JackPotSlots1.frx":F63FF
      Tag             =   "0"
      Top             =   1790
      Width           =   585
   End
   Begin VB.Image Bet1 
      Height          =   585
      Index           =   3
      Left            =   150
      Picture         =   "JackPotSlots1.frx":F947C
      Tag             =   "0"
      Top             =   6050
      Width           =   690
   End
   Begin VB.Image Bet1 
      Height          =   570
      Index           =   2
      Left            =   170
      Picture         =   "JackPotSlots1.frx":FC538
      Tag             =   "0"
      Top             =   4500
      Width           =   690
   End
   Begin VB.Image Bet1 
      Height          =   585
      Index           =   1
      Left            =   130
      Picture         =   "JackPotSlots1.frx":FF5D4
      Tag             =   "0"
      Top             =   2910
      Width           =   690
   End
   Begin VB.Label Credits1 
      BackStyle       =   0  'Transparent
      Caption         =   "360"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1140
      Width           =   1935
   End
   Begin VB.Image Spin1 
      Height          =   585
      Left            =   1070
      Picture         =   "JackPotSlots1.frx":1025FB
      Tag             =   "0"
      Top             =   8080
      Width           =   3285
   End
   Begin VB.Image MaxBet1 
      Height          =   585
      Left            =   2960
      Picture         =   "JackPotSlots1.frx":107CC5
      Tag             =   "0"
      Top             =   7280
      Width           =   1395
   End
   Begin VB.Image BetOne1 
      Height          =   585
      Left            =   1080
      Picture         =   "JackPotSlots1.frx":10C1E7
      Tag             =   "0"
      Top             =   7280
      Width           =   1395
   End
   Begin VB.Image CloseMe1 
      Height          =   450
      Left            =   4040
      Picture         =   "JackPotSlots1.frx":1105E6
      Top             =   510
      Width           =   450
   End
   Begin VB.Image MinimizeMe1 
      Height          =   465
      Left            =   3480
      Picture         =   "JackPotSlots1.frx":11313F
      Top             =   510
      Width           =   450
   End
End
Attribute VB_Name = "JackPotSlots1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'I did set the program to where you have to borrow credits if you run out of money instead
'It just shows a negative number in the credits, that way you can just keep playing plus
'You will know how far your in the hole....LOL :)
Public TracIt2 As Integer 'Declared as Public because this variable gets defined in the BetOne1-MouseDown Event
'Then used in the BetOne1-MouseUp Event
'The Names of Variables can be nearly whatever you want instead of TracIt2, I could have just called it Z
'But by given it a name with character helps to remember what the variable is for so you don't have to keep
'Looking for what it does and where it does it, specially when Declared as Public. You might use the value
'Of a Public variable several times through-out your program
Public RStop1 As Integer, RStop2 As Integer, RStop3 As Integer, RStop4 As Integer, RStop5 As Integer
Public RStop6 As Integer, RStop7 As Integer, RStop8 As Integer, RStop9 As Integer

Private Sub BetOne1_Click()
'Each click of the BetOne button adds 1 to the Tag property of the MaxBet Button
'I used the MaxBet1.Tag property to hold the total number of sinlge bets placed
'Which gets used in the MaxBet1-MouseDown Event

    MaxBet1.Tag = MaxBet1.Tag + 1
End Sub

Private Sub BetOne1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim TracIt As Integer
    
    BetOne1.Picture = ImageList2.ListImages(6).Picture 'Changes picture of BetOne1 to the Up stlye picture
    
    If BetCount(8).Caption = "4" Then 'Looks to see if the caption of BetCount(8) is a 4
        BetOne1.Enabled = False       'If it is then all bets have been placed so it disables the button
            BetOne1.Tag = 0           'Then sets the BetOne.Tag property back to 0 and exits the sub
            Exit Sub
            End If
        
    If BetOne1.Tag >= 8 Then BetOne1.Tag = 0 'When BetOne1.Tag property reaches 8 it has reached the end
    'Of the first round of placing single bets. So it resets to 0 for the for the second round of single
    'Bets, then resets again for third, and the fourth rounds till the last BetCount(8).Caption reaches 4

    'When placing controls in a array they default by starting with index number 0, I set them to start
    'With index number 1. Sense we want it to count thru the BetCount Labels and the Bet1 image controls
    'And sense TracIt starts out with a value of 0 you have to add 1 to its value each time the mouse button
    'Is pressed down, so I set BetOne1.Tag to 0 and use it to store the needed value of variable TracIt.
    
    TracIt = BetOne1.Tag + 1 'Defines TracIt value as the number from BetOne1.Tag
    
        TracIt2 = TracIt 'Defines TracIt2 as having the same value as TracIt, Remember TracIt2 is a Public
        'Variable because we need to use it's value from here in the BetOne1-MouseUp Event Below
        
            'Takes the current number in the BetCount.Caption and adds 1 to it
            BetCount(TracIt).Caption = BetCount(TracIt).Caption + 1
                
                'Tells BetOne1.Tag to be the current value of TracIt, so the next time the MouseDown Event
                'Is called TracIt starts with a value of 2 instead of 1, and so on till BetOne1.Tag reaches
                'A value of 8 where then it is reset to 0 above and the process starts over
                BetOne1.Tag = TracIt
                
                    'Same thing here uses the current value of TracIt to let it know which ImageControl to
                    'Change the picture in, and which picture to pull from the ImageList2. The Down style
                    'Pictures start at index number 19
                    Bet1(TracIt).Picture = ImageList2.ListImages(TracIt + 18).Picture
                    
                        Credits1.Caption = Credits1.Caption - 1 'Subtracts 1 from the current number in the
                        'Credits.Caption each time
                        
                            Spin1.Enabled = True 'Turns the Spin1 button on so it can be used
                    
End Sub

Private Sub BetOne1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Pulls picture number 5 from the ImageList2 control and shows it in the BetOne1 image control
    'So now it will look like the light was turned off.
    BetOne1.Picture = ImageList2.ListImages(5).Picture
        
        'Uses the value of TracIt2 from the BetOne1-MouseDown Event and uses it here to know
        'Which Bet1 image control to change the picture in, and which image to pull from the
        'ImageList2 control to show in Bet1 image control
        Bet1(TracIt2).Picture = ImageList2.ListImages(TracIt2 + 10).Picture

End Sub

Private Sub CloseMe1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Changes the picture of the CloseMe1 image control to picture number 2 from the ImageList2 control
    'Picture number 2 is the Down style of the CLOSE button picture
    CloseMe1.Picture = ImageList2.ListImages(2).Picture
End Sub

Private Sub CloseMe1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Changes the picture of the CloseMe1 image control to picture number 1 from the ImageList2 control
    'Picture number 1 is the Up style of the CLOSE button picture
    CloseMe1.Picture = ImageList2.ListImages(1).Picture
    
        Unload Me 'Program gets turned off here
End Sub

Private Sub Form_Load()
Dim PicLoad As Integer, X

    'For Loops are very handy, instead of manually coding each seperate control such as
    'Below EXAMPLE:
    '   Roll(0).Picture = ImageList1.ListImages(1).Picture
    '     Roll(1).Picture = ImageList1.ListImages(1).Picture
    '       Roll(2).Picture = ImageList1.ListImages(1).Picture
    '         Roll(3).Picture = ImageList1.ListImages(1).Picture
    '           Roll(4).Picture = ImageList1.ListImages(1).Picture
    '             Roll(5).Picture = ImageList1.ListImages(1).Picture
    '               Roll(6).Picture = ImageList1.ListImages(1).Picture
    '                 Roll(7).Picture = ImageList1.ListImages(1).Picture
    '                   Roll(8).Picture = ImageList1.ListImages(1).Picture
    '
    'Both the above code and the code below do exactly the same thing. You can uncomment the 9 lines of
    'Code above and comment out the 3 lines of the For Loop below and see for yourself
    
    For PicLoad = 0 To 8 'Starts with Roll(0) and when it reaches the Next PicLoad it goes back to the
    'Beginning of the For Loop to change the next Roll(1).Picture and so on till it reaches Roll(8)
    'Then it stops.
    
        'PicLoad is the current index number value
        Roll(PicLoad).Picture = ImageList1.ListImages(3).Picture
        
        Next PicLoad 'Moves to the next Roll( index number which is desinated by PicLoad )

    Spin1.Enabled = False 'Button is turned off till a Bet is placed
    
    Timer1.Enabled = False 'Don't want the timer starting until the Spin1 is clicked
     Timer2.Enabled = False
      Timer3.Enabled = False
    
End Sub

Private Sub MaxBet1_Click()
Dim X As Integer
    
    'Work exactly the same as above For Loop, only we're dealing with the captions of the BetCount Labels
    For X = 1 To 8
        BetCount(X).Caption = "4"
        Next X
        
End Sub

Private Sub MaxBet1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BetTally As Integer, PixDown As Integer
    
    'Tells program to show the Down style picture in all the Bet1 image controls
    For PixDown = 1 To 8
        Bet1(PixDown).Picture = ImageList2.ListImages(PixDown + 18).Picture
        Next PixDown
    
    If BetCount(1).Caption = "0" Then 'If the caption of BetCount(1) is a 0 then no single bets have been
    'Placed so says to subtract 32 from the credits and then exit the sub
        Credits1.Caption = Credits1.Caption - 32
        Exit Sub
        End If
    
    'If 1 or more single bets have been placed then, it takes the number in the MaxBet1.Tag and subtracts
    'From 32 the remaining amount is then called BetTally which is subtracted from the Credits1.Caption
    BetTally = 32 - MaxBet1.Tag
        Credits1.Caption = Credits1.Caption - BetTally
     
End Sub

Private Sub MaxBet1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim PixUp As Integer

    'Tells program to show the Up style picture in all the Bet1 image controls
    For PixUp = 1 To 8
        Bet1(PixUp).Picture = ImageList2.ListImages(PixUp + 10).Picture
        Next PixUp
        
    MaxBet1.Picture = ImageList2.ListImages(7).Picture 'Changes the picture back to the Up style
    
        BetOne1.Enabled = False ' Turns BetOne1 button off sense max number of bets have been placed
            MaxBet1.Enabled = False 'Turns MaxBet1 off also sense all bets have been placed
                MaxBet1.Tag = 0 'Resets the Tag property back to 0
                    Spin1.Enabled = True 'Turns the Spin1 button on
End Sub

Private Sub MinimizeMe1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Changes the picture of the Minimize button to the Down style
    MinimizeMe1.Picture = ImageList2.ListImages(4).Picture
End Sub

Private Sub MinimizeMe1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Changes the picture of the Minimize button back to the Up style
    MinimizeMe1.Picture = ImageList2.ListImages(3).Picture
    
        Me.WindowState = 1 'Minimizes program to the TaskBar, the ShowInTaskbar property of
        'JackPotSlots1 is set to True other wise it minimizes it to the bottom left side of the desktop
        'Above the Taskbar and not to the Taskbar. When you use the Form Property BorderStyle 0-none it
        'Defualts to ShowInTaskbar = False
End Sub

Private Sub Spin1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Changes the picture of the Spin1 button to the Down style
    Spin1.Picture = ImageList2.ListImages(10).Picture
End Sub

Private Sub Spin1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Changes the picture of the Spin1 button back to the Up style
    Spin1.Picture = ImageList2.ListImages(9).Picture
    
        Timer1.Enabled = True 'Turns Timers on here
         Timer2.Enabled = True
          Timer3.Enabled = True
        
    Call PicNumbers ' Says to run the PicNumber Sub to get the stop point values
    
End Sub

Private Sub PicNumbers()
Dim G, X As Integer, StopPoints

StopPoints = Array("1", "2", "5", "3", "4", "8", "3", "6", "7", "11", "9", "10", "4", "3")

    For X = 0 To 8 'Randomly pics a number from above array to put in the Tag property of each of
        Randomize  'The Roll() image controls, X desinates which image control to change the tag of
            G = Int(14 * Rnd)
                Roll(X).Tag = StopPoints(G)
                Next X

    RStop1 = Roll(0).Tag 'Defineing the RStop variables so the Timers know which picture to show
      RStop2 = Roll(1).Tag 'When they stop
        RStop3 = Roll(2).Tag
          RStop4 = Roll(3).Tag
            RStop5 = Roll(4).Tag
              RStop6 = Roll(5).Tag
                RStop7 = Roll(6).Tag
                  RStop8 = Roll(7).Tag
                    RStop9 = Roll(8).Tag

End Sub

Private Sub Timer1_Timer()
Dim CountThru1 As Integer, RollAll1 As Integer
    
    If Bet1(2).Tag >= 3 Then 'Looks to see if the number in the tag property of Bet1(2) has reached 3 or greater
        Timer1.Enabled = False 'If it has then it turns the timer off here
        
        'Pictures are stopped at the defined points from the PicNumbers sub above
         Roll(0).Picture = ImageList1.ListImages(RStop1).Picture
          Roll(1).Picture = ImageList1.ListImages(RStop2).Picture
           Roll(2).Picture = ImageList1.ListImages(RStop3).Picture
           Exit Sub
           End If
           
    If Bet1(1).Tag >= 11 Then 'This resets the tag back to 0 to begin rolling thru the images again
        Bet1(1).Tag = 0
            Bet1(2).Tag = Bet1(2).Tag + 1 'Each time it makes a complete pass thru the images it adds 1
            End If 'To keep track of how many times a complete roll has happened

CountThru1 = Bet1(1).Tag + 1 'Defines CountThru1 as the number from the Bet1(1) Tag property and adds 1 to it
    
    For RollAll1 = 0 To 2 'Rolls only the images of the first row of image controls on the left going down
        Roll(RollAll1).Picture = ImageList1.ListImages(CountThru1).Picture
        Next RollAll1
        
Bet1(1).Tag = CountThru1 'Updates the Bet1(1) Tag property as the current value of CountThru1

End Sub

Private Sub Timer2_Timer()
Dim CountThru2, RollAll2, RollCount2
'Explained in Timer1
    If Bet1(4).Tag >= 4 Then
        Timer2.Enabled = False
         Roll(3).Picture = ImageList1.ListImages(RStop4).Picture
          Roll(4).Picture = ImageList1.ListImages(RStop5).Picture
           Roll(5).Picture = ImageList1.ListImages(RStop6).Picture
           Exit Sub
           End If
           
    If Bet1(3).Tag >= 11 Then
        Bet1(3).Tag = 0
            Bet1(4).Tag = Bet1(4).Tag + 1
            End If

CountThru2 = Bet1(3).Tag + 1
RollCount2 = Bet1(4).Tag
    
    For RollAll2 = 3 To 5 '8
        Roll(RollAll2).Picture = ImageList1.ListImages(CountThru2).Picture
        Next RollAll2
        
Bet1(3).Tag = CountThru2

End Sub

Private Sub Timer3_Timer()
Dim CountThru3, RollAll3, RollCount3
'Explained in Timer1
    If Bet1(6).Tag >= 5 Then
        Timer3.Enabled = False
         Roll(6).Picture = ImageList1.ListImages(RStop7).Picture
          Roll(7).Picture = ImageList1.ListImages(RStop8).Picture
           Roll(8).Picture = ImageList1.ListImages(RStop9).Picture
            Call Row1Across
             Call Row2Across
              Call Row3Across
               Call Row1Down
                Call Row2Down
                 Call Row3Down
                  Call DiagonalRowDown
                   Call DiagonalRowUp
                    Call AllFruitCheck
           Exit Sub
           End If
           
    If Bet1(5).Tag >= 11 Then
        Bet1(5).Tag = 0
            Bet1(6).Tag = Bet1(6).Tag + 1
            End If

CountThru3 = Bet1(5).Tag + 1
RollCount3 = Bet1(6).Tag
    
    For RollAll3 = 6 To 8 '8
        Roll(RollAll3).Picture = ImageList1.ListImages(CountThru3).Picture
        Next RollAll3

Bet1(5).Tag = CountThru3

End Sub

Private Sub Row1Across()
'#############################################################################################################
'First number before slash is the index number and number after slash is how much each is worth
'Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp
'Have to check if all three are Cherries first, if they are not then it looks to see if the first two are
'Cherries, if not then it looks to see if only the first one is a Cherry
'If it checked for just one Cherry first then wether you had 2 or 3 wouldn't matter it would only count 1

    'Looks to see if all three are the Cherry images in the top row going across
    If Roll(0).Tag = 3 And Roll(3).Tag = 3 And Roll(6).Tag = 3 Then
            'If they are, then it takes the number of bets placed on that line and multiplies by six
            AddItUp = BetCount(1).Caption * 6
                'Then takes the number and adds the winning amount to the Credits1.Caption
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for two Cherry images in the first two image controls in the top row going across
    If Roll(0).Tag = 3 And Roll(3).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for only one Cherry image in the first image control in the top row going across
    If Roll(0).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three 7's
    If Roll(0).Tag = 2 And Roll(3).Tag = 2 And Roll(6).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Triple Bars
    If Roll(0).Tag = 5 And Roll(3).Tag = 5 And Roll(6).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Double Bars
    If Roll(0).Tag = 8 And Roll(3).Tag = 8 And Roll(6).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Single Bars
    If Roll(0).Tag = 9 And Roll(3).Tag = 9 And Roll(6).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three DollarSigns
    If Roll(0).Tag = 10 And Roll(3).Tag = 10 And Roll(6).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Bells
    If Roll(0).Tag = 11 And Roll(3).Tag = 11 And Roll(6).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three hearts
    If Roll(0).Tag = 6 And Roll(3).Tag = 6 And Roll(6).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Diamonds
    If Roll(0).Tag = 1 And Roll(3).Tag = 1 And Roll(6).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    'Checking for three Limes
    If Roll(0).Tag = 7 And Roll(3).Tag = 7 And Roll(6).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub
Private Sub Row2Across()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(1).Tag = 3 And Roll(4).Tag = 3 And Roll(7).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 3 And Roll(4).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(1).Tag = 2 And Roll(4).Tag = 2 And Roll(7).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(1).Tag = 5 And Roll(4).Tag = 5 And Roll(7).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(1).Tag = 8 And Roll(4).Tag = 8 And Roll(7).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 9 And Roll(4).Tag = 9 And Roll(7).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(1).Tag = 10 And Roll(4).Tag = 10 And Roll(7).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 11 And Roll(4).Tag = 11 And Roll(7).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 6 And Roll(4).Tag = 6 And Roll(7).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(1).Tag = 1 And Roll(4).Tag = 1 And Roll(7).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(1).Tag = 7 And Roll(4).Tag = 7 And Roll(7).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub Row3Across()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(2).Tag = 3 And Roll(5).Tag = 3 And Roll(8).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 3 And Roll(5).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 2 And Roll(5).Tag = 2 And Roll(8).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 5 And Roll(5).Tag = 5 And Roll(8).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 8 And Roll(5).Tag = 8 And Roll(8).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 9 And Roll(5).Tag = 9 And Roll(8).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(2).Tag = 10 And Roll(5).Tag = 10 And Roll(8).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 11 And Roll(5).Tag = 11 And Roll(8).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 6 And Roll(5).Tag = 6 And Roll(8).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 1 And Roll(5).Tag = 1 And Roll(8).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 7 And Roll(5).Tag = 7 And Roll(8).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub Row1Down()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(0).Tag = 3 And Roll(1).Tag = 3 And Roll(2).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 3 And Roll(1).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 2 And Roll(1).Tag = 2 And Roll(2).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 5 And Roll(1).Tag = 5 And Roll(2).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 8 And Roll(1).Tag = 8 And Roll(2).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 9 And Roll(1).Tag = 9 And Roll(2).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(0).Tag = 10 And Roll(1).Tag = 10 And Roll(2).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 11 And Roll(1).Tag = 11 And Roll(2).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 6 And Roll(1).Tag = 6 And Roll(2).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 1 And Roll(1).Tag = 1 And Roll(2).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 7 And Roll(1).Tag = 7 And Roll(2).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub Row2Down()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(3).Tag = 3 And Roll(4).Tag = 3 And Roll(5).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 3 And Roll(4).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(3).Tag = 2 And Roll(4).Tag = 2 And Roll(5).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(3).Tag = 5 And Roll(4).Tag = 5 And Roll(5).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(3).Tag = 8 And Roll(4).Tag = 8 And Roll(5).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 9 And Roll(4).Tag = 9 And Roll(5).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(3).Tag = 10 And Roll(4).Tag = 10 And Roll(5).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 11 And Roll(4).Tag = 11 And Roll(5).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 6 And Roll(4).Tag = 6 And Roll(5).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 1 And Roll(4).Tag = 1 And Roll(5).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(3).Tag = 7 And Roll(4).Tag = 7 And Roll(5).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub Row3Down()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(6).Tag = 3 And Roll(7).Tag = 3 And Roll(8).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 3 And Roll(7).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(6).Tag = 2 And Roll(7).Tag = 2 And Roll(8).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(6).Tag = 5 And Roll(7).Tag = 5 And Roll(8).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(6).Tag = 8 And Roll(7).Tag = 8 And Roll(8).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 9 And Roll(7).Tag = 9 And Roll(8).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(6).Tag = 10 And Roll(7).Tag = 10 And Roll(8).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 11 And Roll(7).Tag = 11 And Roll(8).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 6 And Roll(7).Tag = 6 And Roll(8).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(6).Tag = 1 And Roll(7).Tag = 1 And Roll(8).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
     If Roll(6).Tag = 7 And Roll(7).Tag = 7 And Roll(8).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub DiagonalRowDown()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(0).Tag = 3 And Roll(4).Tag = 3 And Roll(8).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 3 And Roll(4).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 2 And Roll(4).Tag = 2 And Roll(8).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 5 And Roll(4).Tag = 5 And Roll(8).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(0).Tag = 8 And Roll(4).Tag = 8 And Roll(8).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 9 And Roll(4).Tag = 9 And Roll(8).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(0).Tag = 10 And Roll(4).Tag = 10 And Roll(8).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 11 And Roll(4).Tag = 11 And Roll(8).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 6 And Roll(4).Tag = 6 And Roll(8).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 1 And Roll(4).Tag = 1 And Roll(8).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(0).Tag = 7 And Roll(4).Tag = 7 And Roll(8).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub DiagonalRowUp()
'#############################################################################################################
'Index No# for Pix: Diamonds = 1/5, Sevens = 2/50, Cherries = 3/2, Lemons = 4, 3Bars = 5/40,
'         Hearts = 6/10, Limes = 7/3 2Bars = 8/35, 1Bar = 9/30, DollarSign = 10/20, Bell = 11/15
'#############################################################################################################
Dim AddItUp

    If Roll(2).Tag = 3 And Roll(4).Tag = 3 And Roll(6).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 6
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 3 And Roll(4).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 4
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 3 Then
            AddItUp = BetCount(1).Caption * 2
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 2 And Roll(4).Tag = 2 And Roll(6).Tag = 2 Then
            AddItUp = BetCount(1).Caption * 150
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 5 And Roll(4).Tag = 5 And Roll(6).Tag = 5 Then
            AddItUp = BetCount(1).Caption * 120
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
                
    If Roll(2).Tag = 8 And Roll(4).Tag = 8 And Roll(6).Tag = 8 Then
            AddItUp = BetCount(1).Caption * 105
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 9 And Roll(4).Tag = 9 And Roll(6).Tag = 9 Then
            AddItUp = BetCount(1).Caption * 90
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
            
    If Roll(2).Tag = 10 And Roll(4).Tag = 10 And Roll(6).Tag = 10 Then
            AddItUp = BetCount(1).Caption * 60
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 11 And Roll(4).Tag = 11 And Roll(6).Tag = 11 Then
            AddItUp = BetCount(1).Caption * 45
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 6 And Roll(4).Tag = 6 And Roll(6).Tag = 6 Then
            AddItUp = BetCount(1).Caption * 30
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 1 And Roll(4).Tag = 1 And Roll(6).Tag = 1 Then
            AddItUp = BetCount(1).Caption * 15
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
    
    If Roll(2).Tag = 7 And Roll(4).Tag = 7 And Roll(6).Tag = 7 Then
            AddItUp = BetCount(1).Caption * 9
                Credits1.Caption = Credits1.Caption + AddItUp
                Exit Sub
                End If
End Sub

Private Sub AllFruitCheck()
Dim X As Integer
'Running thru all the Roll image controls looking to if the tag is either 3,4 or 7 if it is
'Then it adds 1 to Bet1(8) Tag property to keep track of the number of Roll image controls
'That have fruit pix in them.
    For X = 0 To 8
    
        If Roll(X).Tag = 3 Then
            Bet1(8).Tag = Bet1(8).Tag + 1
            
        ElseIf Roll(X).Tag = 4 Then
            Bet1(8).Tag = Bet1(8).Tag + 1
            
        ElseIf Roll(X).Tag = 7 Then
            Bet1(8).Tag = Bet1(8).Tag + 1
            End If
    Next X

    If Bet1(8).Tag >= 9 Then ' If all 9 have fruit pix then we add 100 to the credits
         Credits1.Caption = Credits1.Caption + 100
            Bet1(8).Tag = 0
                Call ClearAll1
                
    ElseIf Bet1(8).Tag <= 8 Then 'If all are not fruit then it exits the sub no points added
        Bet1(8).Tag = 0
            Call ClearAll1
            Exit Sub
            End If
    
End Sub

Private Sub ClearAll1()
Dim X As Integer, V As Integer, B As Integer

    For X = 0 To 8
        Roll(X).Tag = 0
        Next X
        
    For V = 1 To 8
        Bet1(V).Tag = 0
         Next V
    
    For B = 1 To 8
        BetCount(B).Tag = 0
         BetCount(B).Caption = "0"
         Next B
    
    BetOne1.Tag = 0
     BetOne1.Enabled = True
     
    MaxBet1.Tag = 0
     MaxBet1.Enabled = True
    
    Spin1.Enabled = False
    
End Sub
'
'
'
'
