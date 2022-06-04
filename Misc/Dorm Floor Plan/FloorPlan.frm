VERSION 5.00
Begin VB.Form frmFloorPlan 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   12
   ScaleMode       =   0  'User
   ScaleWidth      =   30
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape shpJeffStool 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   2400
      Width           =   735
   End
   Begin VB.Line linBottleNeck 
      X1              =   12.743
      X2              =   12.781
      Y1              =   6.239
      Y2              =   7.874
   End
   Begin VB.Line linJoshLenToTV 
      X1              =   27.307
      X2              =   10.013
      Y1              =   3.421
      Y2              =   7.849
   End
   Begin VB.Line linLenNathanToTV 
      X1              =   6.068
      X2              =   10.013
      Y1              =   5.635
      Y2              =   7.849
   End
   Begin VB.Shape shpJoshChair 
      FillStyle       =   6  'Cross
      Height          =   1335
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Shape shpNathanChair 
      FillStyle       =   6  'Cross
      Height          =   1335
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Shape shpRefrigerator 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1815
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1695
   End
   Begin VB.Shape shpRecliner 
      FillColor       =   &H00004080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   2055
      Left            =   9960
      Shape           =   2  'Oval
      Top             =   0
      Width           =   1695
   End
   Begin VB.Shape shpJoshCloset 
      FillColor       =   &H00808080&
      FillStyle       =   3  'Vertical Line
      Height          =   1455
      Left            =   8400
      Top             =   5640
      Width           =   3465
   End
   Begin VB.Shape shpJoshStack 
      BorderColor     =   &H80000006&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   1575
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape shpNathanCloset 
      FillColor       =   &H00808080&
      FillStyle       =   3  'Vertical Line
      Height          =   2775
      Left            =   0
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Shape shpNathanStack 
      BorderColor     =   &H80000006&
      FillColor       =   &H00FF0000&
      FillStyle       =   2  'Horizontal Line
      Height          =   1095
      Left            =   5040
      Top             =   0
      Width           =   1215
   End
   Begin VB.Line linWindow 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      X1              =   29.886
      X2              =   29.886
      Y1              =   0
      Y2              =   4.025
   End
   Begin VB.Line linDoor 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   7.585
      X2              =   3.641
      Y1              =   11.874
      Y2              =   11.874
   End
   Begin VB.Shape shpJoshBed 
      FillColor       =   &H000080FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   2385
      Left            =   3120
      Top             =   4680
      Width           =   4965
   End
   Begin VB.Shape shpNathanBed 
      FillColor       =   &H000080FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   2385
      Left            =   120
      Top             =   0
      Width           =   4965
   End
End
Attribute VB_Name = "frmFloorPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

