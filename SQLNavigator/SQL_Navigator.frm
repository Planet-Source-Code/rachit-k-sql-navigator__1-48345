VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SQLNavigator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SQL Navigator"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7515
   Icon            =   "SQL_Navigator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox RText1 
      Appearance      =   0  'Flat
      Height          =   3855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   2040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6800
      _Version        =   393216
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label3 
      Caption         =   "Tables And Views"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Stored Procedure"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select Database"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "SQLNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Developed By: Rachit K
'Mail Me: krachit@indiatimes.com

'Remember: While putting the Connection String _
 just remove the ";INITIAL CATALOG='database name'" _
 sentence.
'In this Basically i have shown how one can _
 retrieve the information like all the tables,views & procedures _
 from the table just by using the _
 SQL System Functions.
'Just send me a mail if you find this useful which will be a credit for me.
 
Option Explicit

Dim rec As ADODB.Recordset
Dim con As ADODB.Connection
Dim conStr As String
Dim newConStr As String

Private Sub Form_Load()
'Put Any Valid Connection String for your _
 SQL Database which you use regularly. _
 But Just Remove the Initial Catalog _
 sentence from it as shown below
conStr = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Data Source=Rachit"

Set con = New ADODB.Connection
con.Open conStr
Set rec = New ADODB.Recordset
rec.Open "sp_databases", con, adOpenStatic, adLockReadOnly
While Not rec.EOF
    Combo1.AddItem rec(0)
    rec.MoveNext
Wend
Set rec = Nothing
Set con = Nothing
End Sub

Private Sub Combo1_Click()
Me.MousePointer = vbHourglass
newConStr = ""
newConStr = conStr & ";Initial Catalog=" & Trim(Combo1.Text)
Set con = New ADODB.Connection
con.Open newConStr
Set rec = New ADODB.Recordset

'This is SQL function for viewing the Tables
rec.Open "Sp_Tables", con, adOpenStatic, adLockReadOnly
'-------------------------------
Combo3.Clear
While Not rec.EOF
'To take only the USER made Tables & not the SYSTEM TABLES
'In order to display all table including the system tables _
just remove this If Condition
    If rec(3) = "TABLE" Or rec(3) = "VIEW" Then
        Combo3.AddItem rec(2)
    End If
    rec.MoveNext
Wend
Set rec = Nothing
'--------------------------------

'This is For viewing the Procedures
'--------------------------------
Set rec = New ADODB.Recordset
rec.Open "Select * from SysObjects Where Type='P'", con, adOpenStatic, adLockReadOnly
Combo2.Clear
While Not rec.EOF
    Combo2.AddItem rec(0)
    rec.MoveNext
Wend
'--------------------------------

Set rec = Nothing
Set con = Nothing
Me.MousePointer = vbNormal
End Sub

Private Sub Command1_Click()
Dim cleanStr As String
If Trim(Combo1.Text) = "" Or Trim(Combo2.Text) = "" Then Exit Sub
Me.MousePointer = vbHourglass
newConStr = ""
Set con = New ADODB.Connection
Set rec = New ADODB.Recordset
newConStr = conStr & ";Initial Catalog=" & Trim(Combo1.Text)
con.Open newConStr
'This is SQL function for Viewing the Text in the Procedure
rec.Open "Sp_HelpText [" & Trim(Combo2.Text) & "]", con, adOpenStatic, adLockReadOnly
RText1.Text = ""
While Not rec.EOF
    cleanStr = rec(0)
    RText1 = RText1 & cleanStr
    rec.MoveNext
Wend
If Grid1.Visible = True Then Grid1.Visible = False
RText1.Visible = True
Set rec = Nothing
Set con = Nothing
Me.MousePointer = vbNormal
End Sub

Private Sub Command2_Click()
'Grid1.Visible = True
If Trim(Combo1.Text) = "" Or Trim(Combo3.Text) = "" Then Exit Sub

Me.MousePointer = vbHourglass
Set rec = New ADODB.Recordset
Set con = New ADODB.Connection
newConStr = conStr & ";INITIAL CATALOG=" & Trim(Combo1.Text)
con.Open newConStr
'This is to View the Data of the Table in the Grid
rec.Open "Select * From [" & Trim(Combo3.Text) & "]", con, adOpenStatic, adLockReadOnly
'To chek wether records are not empty
If rec.EOF <> True And rec.BOF <> True Then
    If RText1.Visible = True Then RText1.Visible = False
    Grid1.Visible = True
    Grid1.ColWidth(0) = 100
    Set Grid1.DataSource = rec
    Me.MousePointer = vbNormal
Else
    Grid1.Visible = False
    Me.MousePointer = vbNormal
    MsgBox "No Records Present in the Table to View", vbCritical
End If
Set rec = Nothing
Set con = Nothing
End Sub

Private Sub Grid1_EnterCell()
    Grid1.CellBackColor = &HC0FFC0
End Sub

Private Sub Grid1_LeaveCell()
    Grid1.CellBackColor = vbWhite
End Sub
