VERSION 5.00
Begin VB.Form frmNewNewFloorPlan 
   Caption         =   "Form1"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   ScaleHeight     =   12
   ScaleMode       =   0  'User
   ScaleWidth      =   30
   StartUpPosition =   3  'Windows Default
   Begin VB.Shape shpClutter 
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   2  'Horizontal Line
      Height          =   1815
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      Height          =   3375
      Left            =   3360
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Shape shpJeffStool 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   735
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   2880
      Width           =   735
   End
   Begin VB.Line linBottleNeck 
      X1              =   13.654
      X2              =   4.248
      Y1              =   10.063
      Y2              =   10.063
   End
   Begin VB.Line linJoshLenToTV 
      X1              =   28.217
      X2              =   23.97
      Y1              =   2.818
      Y2              =   8.252
   End
   Begin VB.Line linLenNathanToTV 
      X1              =   6.675
      X2              =   28.217
      Y1              =   2.214
      Y2              =   3.019
   End
   Begin VB.Shape shpJoshChair 
      FillStyle       =   6  'Cross
      Height          =   1335
      Left            =   6840
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Shape shpNathanChair 
      FillStyle       =   6  'Cross
      Height          =   1335
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Shape shpRefrigerator 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   1815
      Left            =   10080
      Shape           =   4  'Rounded Rectangle
      Top             =   5280
      Width           =   1695
   End
   Begin VB.Shape shpRecliner 
      FillColor       =   &H00004080&
      FillStyle       =   5  'Downward Diagonal
      Height          =   2055
      Left            =   8400
      Shape           =   2  'Oval
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Shape shpJoshCloset 
      FillColor       =   &H00808080&
      FillStyle       =   3  'Vertical Line
      Height          =   1815
      Left            =   120
      Top             =   5280
      Width           =   1545
   End
   Begin VB.Shape shpJoshStack 
      BorderColor     =   &H80000006&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   2  'Horizontal Line
      Height          =   1575
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Shape shpNathanCloset 
      FillColor       =   &H00808080&
      FillStyle       =   3  'Vertical Line
      Height          =   1815
      Left            =   4800
      Top             =   0
      Width           =   1425
   End
   Begin VB.Shape shpNathanStack 
      BorderColor     =   &H80000006&
      FillColor       =   &H00FF0000&
      FillStyle       =   2  'Horizontal Line
      Height          =   1095
      Left            =   3480
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
      X1              =   7.889
      X2              =   4.248
      Y1              =   11.874
      Y2              =   11.874
   End
   Begin VB.Shape shpJoshBed 
      FillColor       =   &H000080FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1785
      Left            =   6240
      Top             =   0
      Width           =   5445
   End
   Begin VB.Shape shpNathanBed 
      FillColor       =   &H000080FF&
      FillStyle       =   4  'Upward Diagonal
      Height          =   5145
      Left            =   120
      Top             =   120
      Width           =   1605
   End
End
Attribute VB_Name = "frmNewNewFloorPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'new design's bottleneck is twice as wide

Private Sub Form_Load()

End Sub
