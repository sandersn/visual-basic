VERSION 2.00
Begin Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   1095
   ClientTop       =   645
   ClientWidth     =   8190
   Height          =   6090
   Icon            =   CLICKER.FRX:0000
   Left            =   1035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   8190
   Top             =   300
   Width           =   8310
   Begin ComboBox cboUse 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin CommandButton cmdUse 
      Caption         =   "&Use..."
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin ComboBox cboGet 
      Height          =   300
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin CommandButton cmdGet 
      Caption         =   "&Get..."
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin CommandButton cmdTalk 
      Caption         =   "&Talk To..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin Image imgTop1 
      Height          =   2415
      Left            =   1680
      Top             =   840
      Width           =   4815
   End
   Begin Image imgBottm1 
      Height          =   2415
      Left            =   1680
      Top             =   3240
      Width           =   4815
   End
   Begin Image imgRight1 
      Height          =   4815
      Left            =   6480
      Top             =   840
      Width           =   1695
   End
   Begin Image imgLeft1 
      Height          =   4815
      Left            =   0
      Top             =   840
      Width           =   1695
   End
End
Option Explicit
Dim CounterTop As Integer
Dim CounterBottom As Integer
Dim CounterLeft As Integer
Dim CounterRight As Integer

Sub Form_Load ()
    imgTop1.Height = 2415
    imgTop1.Width = 4755
End Sub

Sub imgLeft1_Click ()
 CounterLeft = CounterLeft + 1
 MsgBox "CounterLeft's value is " & CounterLeft
End Sub

