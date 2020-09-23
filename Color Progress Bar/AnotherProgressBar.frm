VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "#####  My Color Progress Bar  #####"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Text            =   "35"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6600
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Top             =   720
      Width           =   8295
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   360
      Top             =   720
      Width           =   8300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Write a value for the progressbar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'\\\\\\\\\\Vote If You Like This Color progress bar ////////////
'\\\\\\\\\\Vote If You Like This Color progress bar ////////////
'\\\\\\\\\\Vote If You Like This Color progress bar ////////////
'\\\\\\\\\\Vote If You Like This Color progress bar ////////////
'\\\\\\\\\\Vote If You Like This Color progress bar ////////////



Function ShapeBar(ProgressBarValue As Integer, ProgressBarName As Shape)
ProgressBarName.Width = ProgressBarValue * 82.95 ' Write here the value for one percent
End Function

Private Sub Form_Load()
MyFunctionProgressBar = ShapeBar(Text1, Shape2)

End Sub

Private Sub Text1_Change()
MyFunctionProgressBar = ShapeBar(Text1, Shape2)
If Text1.Text > 100 Then
MsgBox "Maximum Value For Percentage = 100"
Text1.Text = 100
End If
End Sub
