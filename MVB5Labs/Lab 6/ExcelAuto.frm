VERSION 5.00
Object = "{74701400-9DD9-11CF-A662-00AA00C066D2}#1.0#0"; "IEMENU.OCX"
Begin VB.Form frmExcelAuto 
   Caption         =   "Excel Ole Automation"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   1965
   ScaleWidth      =   3615
   StartUpPosition =   2  'CenterScreen
   Begin IEPOPObjectsCtl.IEPOP IEPOP1 
      Height          =   1095
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1931
   End
   Begin VB.Frame fraEarning 
      Caption         =   "Estimated Earnings:"
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2055
      Begin VB.TextBox txtInflation 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtGrowth 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Est. Inflation:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Est. Growth:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdChartEarnings 
      Caption         =   "Chart Earnings"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdEstimateEarnings 
      Caption         =   "Estimate Earnings"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCreateXL 
      Caption         =   "Create XL Object"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmExcelAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub IeLabel1_Click()

End Sub

