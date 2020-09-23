VERSION 5.00
Begin VB.Form frmComboBoxEx 
   Caption         =   "Form1"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Clear"
      Height          =   405
      Left            =   5655
      TabIndex        =   13
      Top             =   2700
      Width           =   1275
   End
   Begin VB.PictureBox Picture3 
      Height          =   570
      Left            =   7200
      Picture         =   "frmComboBoxEx.frx":0000
      ScaleHeight     =   510
      ScaleWidth      =   765
      TabIndex        =   12
      Top             =   795
      Width           =   825
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Style drop dwon list"
      Height          =   525
      Left            =   7200
      TabIndex        =   11
      Top             =   3240
      Width           =   1590
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Style dropdown"
      Height          =   540
      Left            =   5640
      TabIndex        =   10
      Top             =   3210
      Width           =   1425
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1605
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   2100
      Width           =   1320
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Change Selected Item"
      Height          =   555
      Left            =   7110
      TabIndex        =   8
      Top             =   2445
      Width           =   1425
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Select Item 2"
      Height          =   435
      Left            =   7125
      TabIndex        =   7
      Top             =   1920
      Width           =   1440
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Drop Down "
      Height          =   450
      Left            =   5535
      TabIndex        =   6
      Top             =   2205
      Width           =   1305
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Selected Item"
      Height          =   375
      Left            =   5550
      TabIndex        =   5
      Top             =   1725
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Last Item"
      Height          =   390
      Left            =   5565
      TabIndex        =   4
      Top             =   1230
      Width           =   1365
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add item"
      Height          =   390
      Left            =   5490
      TabIndex        =   3
      Top             =   720
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get Count"
      Height          =   405
      Left            =   5505
      TabIndex        =   2
      Top             =   210
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Height          =   390
      Left            =   7680
      Picture         =   "frmComboBoxEx.frx":0442
      ScaleHeight     =   330
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   180
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      Height          =   420
      Left            =   7125
      Picture         =   "frmComboBoxEx.frx":0544
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   90
      Width           =   420
   End
End
Attribute VB_Name = "frmComboBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' Name:     modComboBoxEx.bas
' Author:   Joshy Francis (joshylogicss@yahoo.co.in)
' Date:     3 March 2007
'
' Requires: None
'
' Copyright © 2000-2007 Joshy Francis
' --------------------------------------------------------------------------------------
'The implementation of ComboBoxEx in VB.All by API.
'you can freely use this code anywhere.But I wants you must include the copyright info
'All functions in this module written by me.
' --------------------------------------------------------------------------------------
'No updates.This is the first version.
'I Just included comments on every important lines.Sorry for my bad english.
'I developed this program by converting the C Documentation to VB and experiments with VB.
'You can improve this program by your experiments.I didn't done all parts of the
'ComboBoxEx.
Option Explicit

Private Sub Command1_Click()
MsgBox GetCount
End Sub

Private Sub Command10_Click()
Clear
End Sub

Private Sub Command2_Click()
InsertItem "Item " & GetCount, , (Rnd * 1) * 1, (Rnd * 1) * 1, Rnd * 6
End Sub

Private Sub Command3_Click()
DelItem GetCount - 1
End Sub

Private Sub Command4_Click()
Dim i As Long
    i = GetSelItem
MsgBox GetItemText(i), , i
End Sub

Private Sub Command5_Click()
DropDown
End Sub

Private Sub Command6_Click()
SelItem 1
End Sub

Private Sub Command7_Click()
Dim i As Long
    i = GetSelItem
SetItemText "Changed Item " & i, i
End Sub

Private Sub Command8_Click()
Dim stl As Long
stl = GetWindowLong(Wnd, GWL_STYLE)

If (stl And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST Then
    stl = stl And Not CBS_DROPDOWNLIST
End If
If (stl And CBS_DROPDOWN) = CBS_DROPDOWN Then
Else
    stl = stl Or CBS_DROPDOWN
End If

SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Command9_Click()
Dim stl As Long
stl = GetWindowLong(Wnd, GWL_STYLE)
If (stl And CBS_DROPDOWN) = CBS_DROPDOWN Then
    stl = stl And Not CBS_DROPDOWN
End If
If (stl And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST Then
Else
    stl = stl Or CBS_DROPDOWNLIST
End If
SetWindowLong Wnd, GWL_STYLE, stl
End Sub

Private Sub Form_Load()
CreateComboBoxEx hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyComboBoxEx
End Sub
