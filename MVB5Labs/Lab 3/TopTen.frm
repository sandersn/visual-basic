VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmTopTen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Top Ten Most Expensive Products"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid dbgrdTopTen 
      Bindings        =   "TopTen.frx":0000
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "TopTen.frx":0010
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmTopTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbTopTen As Database

    Set dbTopTen = DBEngine.Workspaces(0).OpenDatabase("C:\Program Files\DevStudio\VB\Nwind.mdb")

    Set Data1.Recordset = dbTopTen.OpenRecordset("Ten Most Expensive Products")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Data1.Recordset.Close
    Unload Me
End Sub
