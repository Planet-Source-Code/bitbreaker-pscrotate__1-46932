VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   1000
      Top             =   2760
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00000000&
      Height          =   6135
      Index           =   1
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   1
      Top             =   0
      Width           =   7215
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   6135
      Index           =   0
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   477
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declaration
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'types
Private Type DOT
    px   As Long
    py   As Long
    pz   As Single
    pc   As Long
End Type

'consts
Private Const MAXP As Integer = 50 * 50

'globals
Private OnTheRun    As Boolean
Private FPS         As Integer
Private shft        As Single
Private src(MAXP)   As DOT
Private dest(MAXP)  As DOT
Private tabl(50)    As String * 50

Private Sub Form_Load()
Dim x As Integer
Dim y As Integer
Dim i As Integer
    
     tabl(0) = ".................................................."
     tabl(1) = ".................................................."
     tabl(2) = ".........*****....*****........********..........."
     tabl(3) = "..........***......***..........******............"
     tabl(4) = "..........***......***...........****............."
     tabl(5) = "..........****....****...........****............."
     tabl(6) = "..........************...........****............."
     tabl(7) = "..........************...........****............."
     tabl(8) = "..........****....****...........****............."
     tabl(9) = "..........***......***...........****............."
    tabl(10) = "..........***......***..........******............"
    tabl(11) = ".........*****....*****........********..........."
    tabl(12) = ".................................................."
    tabl(13) = ".................................................."
    tabl(14) = ".................................................."
    tabl(15) = "**************************************************"
    tabl(16) = "**************************************************"
    tabl(17) = ".................................................."
    tabl(18) = ".................................................."
    tabl(19) = ".................................................."
    tabl(20) = "...........****.....****...****....***............"
    tabl(21) = "...........*****...*****...*****...***............"
    tabl(22) = "...........*************...******..***............"
    tabl(23) = "...........*************...*******.***............"
    tabl(24) = "...........***.*****.***...***.*******............"
    tabl(25) = "...........***..***..***...***..******............"
    tabl(26) = "...........***...*...***...***...*****............"
    tabl(27) = "...........***.......***...***....****............"
    tabl(28) = ".................................................."
    tabl(29) = ".................................................."
    tabl(30) = ".................................................."
    tabl(31) = "**************************************************"
    tabl(32) = "**************************************************"
    tabl(33) = ".................................................."
    tabl(34) = ".................................................."
    tabl(35) = ".................................................."
    tabl(36) = ".................................................."
    tabl(37) = ".....*********........*******.......*****........."
    tabl(38) = ".....***********.....*********....*********......."
    tabl(39) = ".....`***...`****..****....`**...****...****......"
    tabl(40) = "......***....****..*****........***..............."
    tabl(41) = ".....***********....`*******....***..............."
    tabl(42) = ".....*********..........`*****..***..............."
    tabl(43) = "......***..........**......***..`****...****......"
    tabl(44) = "......***..........**********....`*********......."
    tabl(45) = ".....*****..........*******.........*****........."
    tabl(46) = ".................................................."
    tabl(47) = ".................................................."
    tabl(48) = ".................................................."
    tabl(49) = ".................................................."
    
    
    With Me
        .ScaleMode = vbPixels
        .BackColor = &H0
        .Width = 500 * Screen.TwipsPerPixelX
        .Height = 500 * Screen.TwipsPerPixelY
        .AutoRedraw = True
    End With
    
    With pic(0)
        .ScaleMode = vbPixels
        .BackColor = &H0
        .AutoRedraw = True
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    With pic(1)
        .ScaleMode = vbPixels
        .BackColor = &H0
        .AutoRedraw = False
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
    End With
    
    For y = -24 To 25
        For x = -24 To 25
               src(i).px = (x * 4)
               src(i).py = (y * 4)
               src(i).pz = ((y ^ 2) + (x ^ 2)) / 3000
            
               If Mid(tabl(24 + y), 25 + x, 1) = "." Then
                   src(i).pc = RGB(164 - Abs(3 * (x - y)), 164 - Abs(3 * (x - y)), 64 + Abs(3 * (x - y))) Or 64
               Else
                   src(i).pc = Not (RGB(164 - Abs(3 * (x - y)), 164 - Abs(3 * (x - y)), 64 + Abs(3 * (x - y))) Or 64)
               End If
            
               dest(i).px = src(i).px
               dest(i).px = src(i).py
               dest(i).pz = src(i).pz
               dest(i).pc = src(i).pc
               i = i + 1
        Next x
    Next y
   
    shft = 0.01
    OnTheRun = True
    Me.Show
    FPS = 0
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Call DoProcess
End Sub


Private Sub DoProcess()
Dim i    As Integer
Dim nt   As Integer
Static p As Integer
    
    Do While OnTheRun
        p = p + 1
        If p = 100 Then p = 0
        
        pic(0).Cls
        
        Call RotatePoints(p * 3.6)
        
        For i = 0 To MAXP
             pic(0).PSet ((pic(0).ScaleWidth / 2) - dest(i).px, (pic(0).ScaleHeight / 2) - dest(i).py), dest(i).pc
             pic(0).PSet ((pic(0).ScaleWidth / 2) - dest(i).px + 1, (pic(0).ScaleHeight / 2) - dest(i).py + 1), dest(i).pc
        Next i
        
        BitBlt pic(1).hDC, 0, 0, pic(1).Width, pic(1).Height, pic(0).hDC, 0, 0, vbSrcCopy
        FPS = FPS + 1
        DoEvents
    Loop
    
End Sub

Private Sub RotatePoints(ptheta As Single)
Dim i       As Integer
Dim ctheta  As Single
Dim stheta  As Single
Dim theta   As Single

     theta = (3.14159265) * ptheta / 180
    ctheta = Cos(theta)
    stheta = Sin(theta)

    For i = 0 To MAXP
        dest(i).px = (src(i).px * ctheta + src(i).py * stheta) * dest(i).pz
        dest(i).py = (-src(i).px * stheta + src(i).py * ctheta) * dest(i).pz
        dest(i).pz = dest(i).pz + shft
         
        If dest(i).pz > 2 Or dest(i).pz < -2 Then shft = -shft
    Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    OnTheRun = False
End Sub

Private Sub Timer1_Timer()
    Me.Caption = "FPS: " & FPS
    FPS = 0
End Sub

