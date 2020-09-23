VERSION 5.00
Begin VB.Form GetPath 
   AutoRedraw      =   -1  'True
   Caption         =   "GetPath Demo"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   BeginProperty Font 
      Name            =   "Godzilla"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   468
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Godzilla"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   60
      ScaleHeight     =   337
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   1
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Font"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   5340
      Width           =   2595
   End
End
Attribute VB_Name = "GetPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Get Path Demo
'by Scythe
'www.scythe-tools.de

'This is not the best code
'but it seems to be the only example
'existing for GetPath and Visual basic
'(I searched but i dont find any others)

'Thanks to Charles P.V.
'for the idea to create a vector font thru Path

Option Explicit

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetPath Lib "gdi32" (ByVal hdc As Long, lpPoint As PointApi, lpTypes As Byte, ByVal nSize As Long) As Long
Private Declare Function PolyBezierTo Lib "gdi32" (ByVal hdc As Long, lppt As PointApi, ByVal cCount As Long) As Long


'Private Declare Function FlattenPath Lib "gdi32" (ByVal hdc As Long) As Long
'If you dont want to use Bezier then use this api to
'convert all curves in the path into lines
'this shound be easyer for 3D text


Private Type PointApi
 X As Long
 Y As Long
End Type

Private Sub Command1_Click()
 Dim Pth(1000) As PointApi   'This will hold the Path
 Dim Tpe(1000) As Byte       'What is it (Bezier or Line)
 Dim Pbz(1000) As PointApi   'Holds the data for the Bezier
 Dim BzCtr As Long
 Dim Ctr As Long
 Dim i As Long
 Dim NewFig As Long

 'Lets create a path
 BeginPath Pic.hdc

 'Do something to fill our Path
 'You wont see this
 Pic.CurrentX = 20
 Pic.CurrentY = 100
 Pic.Print "Hello World"

 'End the Path
 EndPath Pic.hdc
 
 'Now get the Path
 Ctr = GetPath(Pic.hdc, Pth(0), Tpe(0), 1000)

 Ctr = Ctr - 1

 

 'Now draw Path
 For i = 0 To Ctr
  Select Case Tpe(i)

  Case 6 'Start a new figure
   'Set the Starting point
   'Pset and currentx wont work (dont know why)
   Pic.Line (Pth(i).X, Pth(i).Y)-(Pth(i).X, Pth(i).Y)
   'Hold the startpoit = endpoint
   NewFig = i

  Case 5 'end of a Bezier
   'Now we must increase the Bezier counter
   Pbz(BzCtr).X = Pth(i).X
   Pbz(BzCtr).Y = Pth(i).Y
   BzCtr = BzCtr + 1
   Pbz(BzCtr).X = Pth(i).X
   Pbz(BzCtr).Y = Pth(i).Y
   'Draw the Bezier
   PolyBezierTo Pic.hdc, Pbz(0), BzCtr
   'Reset Counter
   BzCtr = 0
   'Close the figure
   Pic.Line (Pth(i).X, Pth(i).Y)-(Pth(NewFig).X, Pth(NewFig).Y)
  
  Case 3 'end as line
   'if we have an bezier open then draw it
   If BzCtr > 0 Then
    'Set the last bezier Point
    Pbz(BzCtr).X = Pth(i).X
    Pbz(BzCtr).Y = Pth(i).Y
    PolyBezierTo Pic.hdc, Pbz(0), BzCtr
    'Set the current X,Y
    Pic.PSet (Pth(i - 1).X, Pth(i - 1).Y)
    BzCtr = 0
   End If
   'Draw last line
   Pic.Line -(Pth(i).X, Pth(i).Y)
   'Close the figure
   Pic.Line (Pth(i).X, Pth(i).Y)-(Pth(NewFig).X, Pth(NewFig).Y)

  Case 4 'Bezier
   Pbz(BzCtr).X = Pth(i).X
   Pbz(BzCtr).Y = Pth(i).Y
   BzCtr = BzCtr + 1

  Case Else 'Line
   'if we have an bezier open then draw it
   If BzCtr > 0 Then
    'Set the last bezier Point
    Pbz(BzCtr).X = Pth(i).X
    Pbz(BzCtr).Y = Pth(i).Y
    PolyBezierTo Pic.hdc, Pbz(0), BzCtr
    'Set the current X,Y
    Pic.PSet (Pth(i - 1).X, Pth(i - 1).Y)
    BzCtr = 0
   End If
   'Draw line
   Pic.Line -(Pth(i).X, Pth(i).Y)

  End Select

 Next i

End Sub

