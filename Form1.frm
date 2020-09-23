VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Remove Combo Dupes"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3270
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   3270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ReAdd"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove Dupes"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Remove _all_ duplicate Items From Combo"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  KillCmbDupes Combo1
End Sub

Private Sub Command2_Click()
  Dim i As Integer
  With Combo1
    For i = 1 To 10
      .AddItem "AAA"
      .AddItem "BBB"
      .AddItem "CCC"
      .AddItem "DDD"
      .AddItem "EEE"
      .AddItem "Example by RAK"
    Next
    .Text = "AAA"
  End With
End Sub

Private Sub Command3_Click()
  Unload Me
  End
End Sub

Private Sub Form_Load()
  Command2_Click
End Sub
