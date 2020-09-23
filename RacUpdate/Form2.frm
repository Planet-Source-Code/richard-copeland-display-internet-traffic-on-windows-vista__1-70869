VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Log"
   ClientHeight    =   4380
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3375
      Left            =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   50
      Width           =   9135
   End
   Begin VB.Menu mnuoptions 
      Caption         =   "Log Options"
      Begin VB.Menu mnuclear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Text1.Width = Me.Width - 250
Text1.Height = Me.Height - 950
End Sub
