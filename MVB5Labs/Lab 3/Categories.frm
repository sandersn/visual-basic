VERSION 5.00
Begin VB.Form frmCategories 
   Caption         =   "Categories"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTopTen 
      Caption         =   "&Top  Ten"
      Height          =   1455
      Left            =   3960
      TabIndex        =   15
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&lose"
      Height          =   1455
      Left            =   3480
      TabIndex        =   14
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Height          =   1455
      Left            =   3000
      TabIndex        =   13
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   1455
      Left            =   2520
      TabIndex        =   12
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1455
      Left            =   2040
      TabIndex        =   11
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   1455
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   1455
      Left            =   1080
      TabIndex        =   9
      Top             =   2640
      Width           =   225
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   1455
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   225
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtCategoryName 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "Move Next >"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "< Move Previous"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Description:"
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Category &Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Category &ID"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   840
   End
   Begin VB.Label lblCategoryID 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCurrent As Database
Dim recCategories As Recordset

Private Sub cmdAdd_Click()
    recCategories.AddNew
    lblCategoryID.Caption = recCategories("CategoryID")
    txtCategoryName.Enabled = True
    txtDescription.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdFind.Enabled = False
    cmdMoveNext.Enabled = False
    cmdMovePrevious.Enabled = False
    cmdTopTen.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    recCategories.CancelUpdate
    FillFields
    txtCategoryName.Enabled = False
    txtDescription.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
    cmdFind.Enabled = True
    cmdMoveNext.Enabled = True
    cmdMovePrevious.Enabled = True
    cmdTopTen.Enabled = True
End Sub

Private Sub cmdClose_Click()
    dbCurrent.Close
    End
End Sub

Private Sub cmdDelete_Click()
With recCategories
    .Delete
    .MoveNext
    If .EOF = True Then .MoveLast   'oops, we
    'moved past the last record, and need to
    'move to the last.
    'now fill the fields
    FillFields
End With
End Sub

Private Sub cmdEdit_Click()
    recCategories.Edit
    txtCategoryName.Enabled = True
    txtDescription.Enabled = True
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdEdit.Enabled = False
    cmdFind.Enabled = False
    cmdMoveNext.Enabled = False
    cmdMovePrevious.Enabled = False
    cmdTopTen.Enabled = False
End Sub

Private Sub cmdFind_Click()
Dim Result As String
Dim recSearch As Recordset
    Result = InputBox("Input a search string based on the description of the category.", "Find")
    Result = "select * from categories where [Description] like '*" & Result & "*'"
    Set recSearch = dbCurrent.OpenRecordset(Result)
    If recSearch.RecordCount = 0 Then
        MsgBox "No records found!"
        Exit Sub
    End If
    Set recCategories = recSearch
    recCategories.MoveFirst
    

End Sub

Private Sub cmdMoveNext_Click()
    recCategories.MoveNext
    If recCategories.EOF Then recCategories.MoveLast
    FillFields
    txtCategoryName.Enabled = False
    txtDescription.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub cmdMovePrevious_Click()
    recCategories.MovePrevious
    If recCategories.BOF Then recCategories.MoveFirst
    FillFields
    txtCategoryName.Enabled = False
    txtDescription.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
End Sub

Private Sub cmdSave_Click()
    recCategories("CategoryName") = txtCategoryName.Text
    recCategories("Description") = txtDescription.Text
    recCategories.Update
    recCategories.Bookmark = recCategories.LastModified
    txtCategoryName.Enabled = False
    txtDescription.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = False
    cmdAdd.Enabled = True
    cmdDelete.Enabled = True
    cmdEdit.Enabled = True
    cmdFind.Enabled = True
    cmdMoveNext.Enabled = True
    cmdMovePrevious.Enabled = True
    cmdTopTen.Enabled = True
End Sub

Private Sub cmdTopTen_Click()
    frmTopTen.Show vbModal
End Sub

Private Sub Form_Load()
    Set dbCurrent = OpenDatabase("C:\Program Files\DevStudio\VB\Nwind.mdb")
    Set recCategories = dbCurrent.OpenRecordset("Categories")
    recCategories.MoveFirst
    FillFields
End Sub
Private Sub FillFields()
    lblCategoryID.Caption = recCategories.Fields("CategoryID")
    txtCategoryName.Text = recCategories.Fields("CategoryName")
    txtDescription.Text = recCategories.Fields("Description")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    dbCurrent.Close
End Sub
