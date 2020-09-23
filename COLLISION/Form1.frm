VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   336
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   2610
      Top             =   4140
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   60
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   292
      TabIndex        =   3
      Top             =   60
      Width           =   4380
   End
   Begin VB.PictureBox picGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   870
      Picture         =   "Form1.frx":29EF2
      ScaleHeight     =   196
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   292
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   4380
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1590
      Picture         =   "Form1.frx":53DE4
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1020
      Picture         =   "Form1.frx":54626
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CharacterX As Long
Private CharacterY As Long
Private CurrentDirection As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub Form_Load()
    CharacterX = 32
    CharacterY = 32
    Call UpdateImage
End Sub

Private Sub Timer1_Timer()
    Call MoveCharacter
    Call UpdateImage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyLeft Or KeyCode = vbKeyRight) Then
        CurrentDirection = KeyCode
    End If
End Sub

Private Sub MoveCharacter()
    If (CurrentDirection = vbKeyUp) Then
        If Not (CollisionDetected(CharacterX, CharacterY - 1)) Then
            CharacterY = CharacterY - 1
        End If
    ElseIf (CurrentDirection = vbKeyDown) Then
        If Not (CollisionDetected(CharacterX, CharacterY + 1)) Then
            CharacterY = CharacterY + 1
        End If
    ElseIf (CurrentDirection = vbKeyLeft) Then
        If Not (CollisionDetected(CharacterX - 1, CharacterY)) Then
            CharacterX = CharacterX - 1
        End If
    Else
        If (CurrentDirection = vbKeyRight) Then
            If Not (CollisionDetected(CharacterX + 1, CharacterY)) Then
                CharacterX = CharacterX + 1
            End If
        End If
    End If
End Sub

Private Function CollisionDetected(ByVal MoveX As Long, ByVal MoveY As Long) As Boolean
    Dim c As Long
    Dim r As Long

    For r = 0 To picMask.Height - 1
        For c = 0 To picMask.Width - 1
            If GetPixel(picMask.hdc, c, r) = vbBlack Then
                If (GetPixel(picGrid.hdc, MoveX + c, MoveY + r) = vbBlack) Then
                    CollisionDetected = True
                    Exit Function
                End If
            End If
        Next
    Next
    CollisionDetected = False
End Function

Private Sub UpdateImage()
    Call BitBlt(picGame.hdc, 0, 0, picGrid.ScaleWidth, picGrid.ScaleHeight, picGrid.hdc, 0, 0, vbSrcCopy)
    Call BitBlt(picGame.hdc, CharacterX, CharacterY, picMask.ScaleWidth, picMask.ScaleHeight, picMask.hdc, 0, 0, vbSrcAnd)
    Call BitBlt(picGame.hdc, CharacterX, CharacterY, picImage.ScaleWidth, picImage.ScaleHeight, picImage.hdc, 0, 0, vbSrcPaint)
    Call picGame.Refresh
End Sub
