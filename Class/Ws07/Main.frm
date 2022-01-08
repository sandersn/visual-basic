VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLibrary 
      Caption         =   "&Music Library Maintenance"
      Height          =   1695
      Left            =   2093
      Picture         =   "Main.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   368
      Width           =   1695
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "Music LIbrary &Browser"
      Height          =   1695
      Left            =   173
      Picture         =   "Main.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   368
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowser_Click()
    frmBrowser.Show
End Sub

Private Sub cmdLibrary_Click()
    frmMusicHdr.Show
End Sub
