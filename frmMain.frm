VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "老虎机v1.2 - 作者sysdzw"
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10200
   ScaleWidth      =   8025
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   9
      Left            =   7080
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   61
      Top             =   7920
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   8
      Left            =   7080
      Picture         =   "frmMain.frx":1C96
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   60
      Top             =   7080
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   7
      Left            =   7080
      Picture         =   "frmMain.frx":392C
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   59
      Top             =   6240
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   6
      Left            =   7080
      Picture         =   "frmMain.frx":55C2
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   58
      Top             =   5400
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   5
      Left            =   7080
      Picture         =   "frmMain.frx":7258
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   57
      Top             =   4560
      Width           =   765
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6600
      Top             =   0
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   4
      Left            =   7080
      Picture         =   "frmMain.frx":8EEE
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   55
      Top             =   3720
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   3
      Left            =   7080
      Picture         =   "frmMain.frx":AB84
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   54
      Top             =   2880
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   2
      Left            =   7080
      Picture         =   "frmMain.frx":C81A
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   53
      Top             =   2040
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   1
      Left            =   7080
      Picture         =   "frmMain.frx":E4B0
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   52
      Top             =   1200
      Width           =   765
   End
   Begin VB.PictureBox picCoin 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   0
      Left            =   7080
      Picture         =   "frmMain.frx":10146
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   51
      Top             =   240
      Width           =   765
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   960
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   28
      Top             =   7800
      Width           =   735
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   5
      Left            =   5760
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   48
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   5
      Left            =   5760
      TabIndex        =   47
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   5400
      TabIndex        =   39
      Text            =   "0"
      Top             =   405
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   1320
      TabIndex        =   37
      Text            =   "0"
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   4
      Left            =   4800
      TabIndex        =   36
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   3
      Left            =   3840
      TabIndex        =   35
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   2
      Left            =   2880
      TabIndex        =   34
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   1
      Left            =   1920
      TabIndex        =   33
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox txtXiazhu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00F6F6F6&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Index           =   0
      Left            =   960
      TabIndex        =   32
      Text            =   "0"
      Top             =   7320
      Width           =   735
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   3840
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   30
      Top             =   7800
      Width           =   735
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   2880
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   29
      Top             =   7800
      Width           =   735
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   1920
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   27
      Top             =   7800
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   25
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   25
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   24
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   24
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   23
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   23
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   22
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   22
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   21
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   21
      Top             =   5400
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   20
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   20
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   19
      Left            =   1200
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   19
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   18
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   18
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   17
      Left            =   2880
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   17
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   16
      Left            =   3720
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   16
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   15
      Left            =   4560
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   15
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   14
      Left            =   5400
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   14
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   13
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   13
      Top             =   6240
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   12
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   12
      Top             =   5400
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   11
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   11
      Top             =   4560
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   10
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   10
      Top             =   3720
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   9
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   9
      Top             =   2880
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   8
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   7
      Left            =   6240
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   6
      Left            =   5400
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   5
      Left            =   4560
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   4
      Left            =   3720
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   3
      Left            =   2880
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   2
      Left            =   2040
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   360
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picBottom 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   1
      Left            =   1200
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.PictureBox picXiazhu 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   4
      Left            =   4800
      ScaleHeight     =   735
      ScaleWidth      =   735
      TabIndex        =   31
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   56
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "随机分布"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   480
      TabIndex        =   50
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 10"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   49
      Tag             =   "100"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 5"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   46
      Tag             =   "100"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   45
      Tag             =   "10"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   2
      Left            =   2880
      TabIndex        =   44
      Tag             =   "5"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   43
      Tag             =   "4"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label lblBeilv 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X 2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   42
      Tag             =   "2"
      Top             =   8640
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3240
      Picture         =   "frmMain.frx":11DDC
      Top             =   480
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "余额不足！！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1140
      TabIndex        =   41
      Top             =   2400
      Visible         =   0   'False
      Width           =   5040
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "可用点"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   570
      Left            =   4000
      TabIndex        =   40
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "得分"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   570
      Left            =   360
      TabIndex        =   38
      Top             =   315
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFC0&
      BorderWidth     =   10
      Height          =   780
      Left            =   1320
      Top             =   5280
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1080
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   945
      Left            =   4800
      Picture         =   "frmMain.frx":1286E
      Top             =   9120
      Width           =   1275
   End
   Begin VB.Image Image1 
      Height          =   10560
      Left            =   -5280
      Picture         =   "frmMain.frx":167B0
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   13920
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================================
'名    称：vb老虎机程序
'描    述：模仿显示路边小店的老虎机程序做的
'使用方法：用法同实物老虎机，点击硬币投币即可，然后押注水果，然后开始
'编    程：sysdzw 原创开发，如果有需要对模块扩充或更新的话请邮箱发我一份
'发布日期：2016-02-16
'博    客：http://blog.163.com/sysdzw
'          http://blog.csdn.net/sysdzw
'Email   ：sysdzw@163.com
'QQ      ：171977759
'版    本：V1.0.0   初版                                                        2016-02-14
'版    本：V1.2.0   改进                                                        2016-02-16
'==============================================================================================
Option Explicit

Dim objTimer As New clsWaitableTimer
Dim intSecontAll As Integer

Private Sub Form_Load()
    intSecontAll = 300
    Me.Show
    Call refreshPic
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objTimer = Nothing
    End
End Sub

Private Sub Image2_Click()
    Dim i%, j%
    Dim intRnd%
    Dim intCircle%, intRndCircle%
    Dim lngCount%
    
    If txtXiazhu(0).Text + txtXiazhu(1).Text + txtXiazhu(2).Text + txtXiazhu(3).Text + txtXiazhu(4).Text + txtXiazhu(5).Text <= 0 Then
        If Val(Text2.Text) <= 0 Then
            If Val(Text1.Text) > 0 Then
                MsgBox "可用点为0，无法开始。请投币！提示：您可以点击右箭头将得分转为可用点。", vbExclamation
            Else
                MsgBox "可用点为0，无法开始。请投币！", vbExclamation
            End If
        Else
            MsgBox "您还没有下注 ，请下注后再点击“开始”按钮！", vbExclamation
        End If
        Exit Sub
    End If
    
    Me.Enabled = False
    Randomize
    intRndCircle = 3 + Int(Rnd * 10)
    Shape1.Visible = True
    For i = 0 To intRndCircle
        For j = 0 To 25
            objTimer.Wait 20
            Shape1.Move picBottom(j).Left - 15, picBottom(j).Top - 15
        Next
    Next
    
    Randomize
    intRnd = Int(Rnd * 25)
    
    For i = 0 To 1 + Int(Rnd * 2)
        For j = 0 To 25
            objTimer.Wait 70
            Shape1.Move picBottom(j).Left - 15, picBottom(j).Top - 15
        Next
    Next
    
    For j = 0 To intRnd
        objTimer.Wait 70
        Shape1.Move picBottom(j).Left - 15, picBottom(j).Top - 15
    Next
    flashShape1
    
    Dim lngGet&
    lngGet = Val(Split(lblBeilv(picBottom(intRnd).Tag - 1).Caption, " ")(1)) * Val(txtXiazhu(picBottom(intRnd).Tag - 1).Text)
    If lngGet > 0 Then
        Text1.Text = Val(Text1.Text) + lngGet
        flashTxt2
    End If
    
    For i = 0 To 5
        txtXiazhu(i).Text = "0"
    Next
    Me.Enabled = True

'    Me.Caption = Now & " ok"
End Sub

Private Sub refreshPic()
    Dim i%, j%, intRndPic%, strImg%, intRnd%, intXingCount%, intCarCount%
    intRnd = 3 + Int(Rnd * 8)
    Label1.Caption = intRnd
    Label1.Visible = True
    
    Me.Enabled = False
    For i = 1 To 6
        Set picXiazhu(i - 1).Picture = LoadPicture(App.Path & "\" & i & ".bmp")
        picXiazhu(i - 1).Tag = i
    Next
    
    For j = 1 To intRnd
        Label1.Caption = intRnd - j + 1: DoEvents
        For i = 0 To 25
            objTimer.Wait 5
err1:
            Randomize
            intRndPic = Int(Rnd * 6) + 1
            If j = intRnd Then
                If intRndPic = 5 Then intXingCount = intXingCount + 1
                If intRndPic = 6 Then intCarCount = intCarCount + 1
                If intXingCount >= 3 And intRndPic = 5 Then
                    GoTo err1
                End If
                If intCarCount >= 2 And intRndPic = 6 Then
                    GoTo err1
                End If
            End If
            Set picBottom(i).Picture = LoadPicture(App.Path & "\" & intRndPic & ".bmp")
            picBottom(i).Tag = intRndPic
        Next
    Next
    Me.Enabled = True
    Label1.Visible = False
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.BorderStyle = 1
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2.BorderStyle = 0
End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 1
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.BorderStyle = 0
End Sub
Private Sub Image4_Click()
    If Val(Text1.Text) > 0 Then
        Text1.Text = Val(Text1.Text) - 1
        Text2.Text = Val(Text2.Text) + 1
    Else
        showTips "得分为0，无法移到右边！！！"
    End If
End Sub

Private Sub Label5_Click()
    Call refreshPic
End Sub

Private Sub picCoin_Click(Index As Integer)
    Dim i%, lngLeft&, lngTop&
    picCoin(Index).BorderStyle = 0
    lngLeft = picCoin(Index).Left
    lngTop = picCoin(Index).Top
    Me.Enabled = False
    Do
        If picCoin(Index).Top > 0 Then
            picCoin(Index).Move picCoin(Index).Left - 10, picCoin(Index).Top - 10
            objTimer.Wait 1
        Else
            For i = 1 To 50
                picCoin(Index).Move picCoin(Index).Left - 10, picCoin(Index).Top + 10
                objTimer.Wait 2
            Next
        End If
        If picCoin(Index).Left < 0 Then
            Do
                picCoin(Index).Move picCoin(Index).Left + 10, 240: DoEvents
                objTimer.Wait 0.5
                If picCoin(Index).Left > Text2.Left Then Exit Do
            Loop
            Exit Do
        End If
    Loop
    Me.Enabled = True
    picCoin(Index).Visible = False
    picCoin(Index).Left = lngLeft
    picCoin(Index).Top = lngTop
    
    Text2.Text = Val(Text2.Text) + 10
    DoEvents
    flashTxt
End Sub

Private Sub picXiazhu_Click(Index As Integer)
    If Val(Text2.Text) > 0 Then
        txtXiazhu(Index).Text = Val(txtXiazhu(Index).Text) + 1
        Text2.Text = Val(Text2.Text) - 1
    Else
        showTips "可用点不足，请投币！！！"
    End If
End Sub

Private Sub showTips(strMsg$)
    Dim i%
    Me.Enabled = False
    Label4.Caption = strMsg: DoEvents
    For i = 1 To 3
        Label4.Visible = True: DoEvents
        objTimer.Wait 300
        Label4.Visible = False: DoEvents
        objTimer.Wait 50
    Next
    Me.Enabled = True
End Sub

Private Sub flashTxt()
    Dim i%
    Dim a
    a = Text2.BackColor
    For i = 1 To 3
        Text2.BackColor = vbRed
        objTimer.Wait 200
        Text2.BackColor = a
        objTimer.Wait 200
    Next
End Sub
Private Sub flashTxt2()
    Dim i%
    Dim a
    a = Text1.BackColor
    For i = 1 To 3
        Text1.BackColor = vbGreen
        objTimer.Wait 200
        Text1.BackColor = a
        objTimer.Wait 200
    Next
End Sub
Private Sub flashShape1()
    Dim i%
    Dim a
    a = Shape1.BorderColor
    For i = 1 To 5
        Shape1.BorderColor = vbGreen
        objTimer.Wait 200
        Shape1.BorderColor = a
        objTimer.Wait 200
    Next
End Sub
Private Sub picXiazhu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picXiazhu(Index).BorderStyle = 1
End Sub

Private Sub picXiazhu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picXiazhu(Index).BorderStyle = 0
End Sub

Private Sub Timer1_Timer()
    intSecontAll = intSecontAll - 1
    If intSecontAll = 0 Then
        Dim i%
        For i = 0 To picCoin.Count - 1 '恢复第一个隐藏了的硬币
            If picCoin(i).Visible = False Then
                picCoin(i).Visible = True
                picCoin(i).BorderStyle = 1
                Exit For
            End If
        Next
        intSecontAll = 300
    Else
        lblTime.Caption = intSecontAll \ 60 & ":" & intSecontAll Mod 60
    End If
    
End Sub
