VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "frmMain"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9285
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   619
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTitle 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   0
      ScaleHeight     =   42
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   619
      TabIndex        =   0
      Top             =   6570
      Width           =   9285
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' February 14, 2002

'===================================================================================================
DefSng A-Z
'===================================================================================================

'===================================================================================================
Option Explicit
'===================================================================================================

'===================================================================================================
Private Type STRUC_FERN
    a(4) As Single
    b(4) As Single
    c(4) As Single
    d(4) As Single
    e(4) As Single
    f(4) As Single
    p(4) As Single
    
    lmtsx As Integer
    lmtsy As Integer
    lmtdx As Integer
    lmtdy As Integer
    
    xscale As Integer
    yscale As Integer
    
    xoffset As Integer
    yoffset As Integer
    
    clrR As Integer
    clrG As Integer
    clrB As Integer
End Type
'===================================================================================================

'===================================================================================================
Const Title As String = "Fern using Fractals"
'===================================================================================================

Private Sub Form_Activate()
    Do While True
        SetFernStyle
        DoEvents
    Loop
End Sub

'===================================================================================================
Private Sub Form_Load()
    DrawTitle
End Sub

'===================================================================================================
Private Sub Form_Click()
    End
End Sub

'===================================================================================================
Private Sub DrawTitle()
    Dim x As Integer, y As Integer
    Dim sw As Single, sh As Single
    Dim sx As Single, sy As Single
    Dim dx As Single, dy As Single
    
    sw = picTitle.ScaleWidth
    sh = picTitle.ScaleHeight
    
    sx = 0
    sy = sh - picTitle.TextHeight(Title)
    
    dx = picTitle.TextWidth(Title)
    dy = picTitle.TextHeight(Title)
        
    picTitle.Cls
    
    picTitle.ForeColor = RGB(0, 255, 0)
    picTitle.CurrentX = sx + 1
    picTitle.CurrentY = sy + 1
    picTitle.Print Title
    
    picTitle.ForeColor = vbWhite
    picTitle.CurrentX = sx
    picTitle.CurrentY = sy
    picTitle.Print Title
    
    For y = sy To dy + sy
        For x = sx To dx
            If picTitle.Point(x, y) = vbWhite Then
                picTitle.PSet (x, y), RGB(0, ((x + y) Mod 125) + 100, 0)
            End If
        Next x
    Next y
End Sub

'===================================================================================================
Private Sub SetFernStyle()
    Static i As Integer
    Dim SF As STRUC_FERN
    
    Me.Cls
    
    With SF
        .a(0) = 0
        .b(0) = 0
        .c(0) = 0
        .d(0) = 0.16
        .e(0) = 0
        .f(0) = 0
    
        .a(1) = 0.2
        .b(1) = -0.26
        .c(1) = 0.23
        .d(1) = 0.22
        .e(1) = 0
        .f(1) = 1.6
    
        .a(2) = -0.51
        .b(2) = 0.28
        .c(2) = 0.26
        .d(2) = 0.24
        .e(2) = 0
        .f(2) = 1.6
    
        .a(3) = 0.85
        .b(3) = 0.04
        .c(3) = -0.04
        .d(3) = 0.85
        .e(3) = 0
        .f(3) = 2
    
        .p(0) = 328
        .p(1) = 2621
        .p(2) = 4915
        .p(3) = 32767
    
        .xscale = 30
        .yscale = 30
        .xoffset = Me.ScaleWidth / 2
        .yoffset = -75
        
        .lmtsx = 0
        .lmtsy = 0
        .lmtdx = 640
        .lmtdy = 340
        
        Select Case i
        Case Is = 1
            .clrR = 255
            .clrG = 255
            .clrB = 255
        Case Is = 2
            .clrR = 0
            .clrG = 125
            .clrB = 255
        Case Is = 3
            .clrR = 0
            .clrG = 0
            .clrB = 125
        Case Is = 4
            .clrR = 125
            .clrG = 125
            .clrR = 0
        Case Is = 5
            .clrR = 0
            .clrG = 125
            .clrB = 125
        Case Is = 6
            .clrR = 125
            .clrG = 0
            .clrB = 125
        Case Else
            .clrR = 0
            .clrG = 125
            .clrB = 0
        End Select
        
        i = i + 1
        If i > 6 Then i = 0
        
        DrawFern SF
    End With
End Sub

'===================================================================================================
Private Sub DrawFern(SF As STRUC_FERN)
    Dim i As Integer, px As Integer, py As Integer
    Dim j As Integer, k As Single
    Dim newx As Single, x As Single, y As Single
    Dim xloc As Single, yloc As Single
    
    x = 0
    y = 0
    
    With SF
    For i = 1 To 20000
        j = 20000 * Rnd + 1
        k = IIf(j < .p(0), 0, IIf(j < .p(1), 1, IIf(j < .p(2), 2, 3)))
        newx = (.a(k) * x + .b(k) * y + .e(k))
        y = (.c(k) * x + .d(k) * y + .f(k))
        x = newx
        px = x * .xscale + .xoffset
        py = (y * .yscale + .yoffset)
        
        Me.PSet (px, 350 - py), _
            RGB(IIf(.clrR = 0, 0, (px + .clrR) Mod 255 + .clrR), _
                IIf(.clrG = 0, 0, (px + .clrG) Mod 255 + .clrG), _
                IIf(.clrB = 0, 0, (px + .clrB) Mod 255 + .clrB))
        
        If i Mod 100 = 0 Then DoEvents
    Next i
    End With
End Sub
'===================================================================================================

