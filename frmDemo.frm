VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image Processor v0.5 beta"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "frmDemo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   3960
      Left            =   7155
      TabIndex        =   37
      Top             =   15
      Width           =   1830
      Begin VB.CommandButton cmdCustomDoodad 
         Caption         =   "Custom..."
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   41
         Top             =   3540
         Width           =   1635
      End
      Begin VB.PictureBox picDoodad 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   15  'Merge Pen Not
         Height          =   1110
         Index           =   2
         Left            =   90
         Picture         =   "frmDemo.frx":2CFA
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   40
         Top             =   2415
         Width           =   1635
      End
      Begin VB.PictureBox picDoodad 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   15  'Merge Pen Not
         Height          =   1110
         Index           =   1
         Left            =   90
         Picture         =   "frmDemo.frx":396C
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   39
         Top             =   1305
         Width           =   1635
      End
      Begin VB.PictureBox picDoodad 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         DrawMode        =   15  'Merge Pen Not
         Height          =   1110
         Index           =   0
         Left            =   90
         Picture         =   "frmDemo.frx":48C0
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   105
         TabIndex        =   38
         Top             =   195
         Width           =   1635
      End
   End
   Begin VB.Frame fraConfiguration 
      Caption         =   "Configuration"
      Height          =   2145
      Left            =   15
      TabIndex        =   5
      Top             =   3990
      Width           =   8970
      Begin VB.OptionButton optCheapTab 
         Caption         =   "Effects"
         Height          =   360
         Index           =   4
         Left            =   5265
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1680
         Width           =   1275
      End
      Begin VB.CommandButton cmdProcess 
         Caption         =   "Process"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6975
         TabIndex        =   11
         Top             =   1680
         Width           =   1845
      End
      Begin VB.OptionButton optCheapTab 
         Caption         =   "Deformations"
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   3975
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   1275
      End
      Begin VB.OptionButton optCheapTab 
         Caption         =   "More Filters"
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   2685
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1275
      End
      Begin VB.PictureBox picCheapTabstrip 
         Height          =   1425
         Left            =   90
         ScaleHeight     =   1365
         ScaleWidth      =   8685
         TabIndex        =   8
         Top             =   225
         Width           =   8745
         Begin VB.PictureBox picTab 
            BorderStyle     =   0  'None
            Height          =   1365
            Index           =   1
            Left            =   0
            ScaleHeight     =   91
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   579
            TabIndex        =   57
            Top             =   0
            Visible         =   0   'False
            Width           =   8685
            Begin VB.Frame fraFilters 
               Caption         =   "Filters"
               Height          =   1185
               Left            =   45
               TabIndex        =   58
               Top             =   30
               Width           =   8625
               Begin VB.OptionButton optFilters 
                  Caption         =   "Simple Tri-Linear Blur"
                  Height          =   390
                  Index           =   2
                  Left            =   2130
                  Style           =   1  'Graphical
                  TabIndex        =   62
                  Top             =   255
                  Width           =   2010
               End
               Begin VB.OptionButton optFilters 
                  Caption         =   "Simple Bi-Linear Blur"
                  Height          =   390
                  Index           =   1
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   61
                  Top             =   645
                  Width           =   2010
               End
               Begin VB.OptionButton optFilters 
                  Caption         =   "None"
                  Height          =   390
                  Index           =   0
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   59
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   2010
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  Caption         =   "v1.0 coming in january, with dozens of new effects!"
                  Height          =   195
                  Left            =   4785
                  TabIndex        =   60
                  Top             =   915
                  Width           =   3750
               End
            End
         End
         Begin VB.PictureBox picTab 
            BorderStyle     =   0  'None
            Height          =   1365
            Index           =   4
            Left            =   0
            ScaleHeight     =   91
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   579
            TabIndex        =   43
            Top             =   0
            Visible         =   0   'False
            Width           =   8685
            Begin VB.Frame Frame4 
               Caption         =   "Settings"
               Height          =   1185
               Left            =   30
               TabIndex        =   45
               Top             =   30
               Width           =   1440
               Begin VB.TextBox txtOpacity 
                  Height          =   285
                  Left            =   840
                  TabIndex        =   48
                  Text            =   "255"
                  Top             =   195
                  Width           =   510
               End
               Begin VB.TextBox txtDestOpacity 
                  Height          =   285
                  Left            =   840
                  TabIndex        =   47
                  Text            =   "0"
                  Top             =   510
                  Width           =   510
               End
               Begin VB.TextBox txtSourceOpacity 
                  Height          =   285
                  Left            =   840
                  TabIndex        =   46
                  Text            =   "255"
                  Top             =   810
                  Width           =   510
               End
               Begin VB.Label lblOpacity 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Opacity:"
                  Height          =   195
                  Left            =   210
                  TabIndex        =   51
                  Top             =   240
                  Width           =   615
               End
               Begin VB.Label lblDestOpacity 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Dest:"
                  Height          =   195
                  Left            =   435
                  TabIndex        =   50
                  Top             =   555
                  Width           =   390
               End
               Begin VB.Label lblSourceOpacity 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Source:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   49
                  Top             =   855
                  Width           =   555
               End
            End
            Begin VB.Frame fraEffects 
               Caption         =   "Effects"
               Height          =   1185
               Left            =   1500
               TabIndex        =   44
               Top             =   30
               Width           =   7170
               Begin VB.OptionButton optEffects 
                  Caption         =   "Subtractive Blend"
                  Height          =   390
                  Index           =   2
                  Left            =   2130
                  Style           =   1  'Graphical
                  TabIndex        =   54
                  Top             =   255
                  Width           =   2010
               End
               Begin VB.OptionButton optEffects 
                  Caption         =   "Additive Blend"
                  Height          =   390
                  Index           =   1
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   53
                  Top             =   645
                  Width           =   2010
               End
               Begin VB.OptionButton optEffects 
                  Caption         =   "Alpha Blend"
                  Height          =   390
                  Index           =   0
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   52
                  Top             =   255
                  Value           =   -1  'True
                  Width           =   2010
               End
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "v1.0 coming in january, with dozens of new effects!"
                  Height          =   195
                  Left            =   3330
                  TabIndex        =   56
                  Top             =   915
                  Width           =   3750
               End
            End
         End
         Begin VB.PictureBox picTab 
            BorderStyle     =   0  'None
            Height          =   1365
            Index           =   0
            Left            =   0
            ScaleHeight     =   91
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   579
            TabIndex        =   12
            Top             =   0
            Width           =   8685
            Begin VB.Frame fraColorOptions 
               Caption         =   "Options"
               Height          =   1185
               Left            =   3855
               TabIndex        =   34
               Top             =   30
               Width           =   4800
               Begin VB.CheckBox chkInvert 
                  Caption         =   "Invert"
                  Height          =   225
                  Left            =   75
                  TabIndex        =   36
                  Top             =   480
                  Width           =   750
               End
               Begin VB.CheckBox chkGrayscale 
                  Caption         =   "Grayscale"
                  Height          =   225
                  Left            =   75
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1005
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "v1.0 coming in january, with dozens of new effects!"
                  Height          =   195
                  Left            =   975
                  TabIndex        =   55
                  Top             =   915
                  Width           =   3750
               End
            End
            Begin VB.Frame fraBlue 
               Caption         =   "Blue"
               Height          =   1185
               Left            =   2580
               TabIndex        =   27
               Top             =   30
               Width           =   1245
               Begin VB.TextBox txtBlueAdjust 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   30
                  Text            =   "0"
                  Top             =   195
                  Width           =   510
               End
               Begin VB.TextBox txtBlueMin 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   29
                  Text            =   "0"
                  Top             =   510
                  Width           =   510
               End
               Begin VB.TextBox txtBlueMax 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   28
                  Text            =   "255"
                  Top             =   810
                  Width           =   510
               End
               Begin VB.Label lblBlueAdjust 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Adjust:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   33
                  Top             =   240
                  Width           =   525
               End
               Begin VB.Label lblBlueMin 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Min:"
                  Height          =   195
                  Left            =   330
                  TabIndex        =   32
                  Top             =   555
                  Width           =   300
               End
               Begin VB.Label lblBlueMax 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Max:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   31
                  Top             =   855
                  Width           =   360
               End
            End
            Begin VB.Frame fraGreen 
               Caption         =   "Green"
               Height          =   1185
               Left            =   1305
               TabIndex        =   20
               Top             =   30
               Width           =   1245
               Begin VB.TextBox txtGreenAdjust 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   23
                  Text            =   "0"
                  Top             =   195
                  Width           =   510
               End
               Begin VB.TextBox txtGreenMin 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   22
                  Text            =   "0"
                  Top             =   510
                  Width           =   510
               End
               Begin VB.TextBox txtGreenMax 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   21
                  Text            =   "255"
                  Top             =   810
                  Width           =   510
               End
               Begin VB.Label lblGreenAdjust 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Adjust:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   26
                  Top             =   240
                  Width           =   525
               End
               Begin VB.Label lblGreenMin 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Min:"
                  Height          =   195
                  Left            =   330
                  TabIndex        =   25
                  Top             =   555
                  Width           =   300
               End
               Begin VB.Label lblGreenMax 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Max:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   24
                  Top             =   855
                  Width           =   360
               End
            End
            Begin VB.Frame fraRed 
               Caption         =   "Red"
               Height          =   1185
               Left            =   30
               TabIndex        =   13
               Top             =   30
               Width           =   1245
               Begin VB.TextBox txtRedMax 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   19
                  Text            =   "255"
                  Top             =   810
                  Width           =   510
               End
               Begin VB.TextBox txtRedMin 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   17
                  Text            =   "0"
                  Top             =   510
                  Width           =   510
               End
               Begin VB.TextBox txtRedAdjust 
                  Height          =   285
                  Left            =   645
                  TabIndex        =   15
                  Text            =   "0"
                  Top             =   195
                  Width           =   510
               End
               Begin VB.Label lblRedMax 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Max:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   18
                  Top             =   855
                  Width           =   360
               End
               Begin VB.Label lblRedMin 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Min:"
                  Height          =   195
                  Left            =   330
                  TabIndex        =   16
                  Top             =   555
                  Width           =   300
               End
               Begin VB.Label lblRedAdjust 
                  Alignment       =   1  'Right Justify
                  AutoSize        =   -1  'True
                  Caption         =   "Adjust:"
                  Height          =   195
                  Left            =   105
                  TabIndex        =   14
                  Top             =   240
                  Width           =   525
               End
            End
         End
      End
      Begin VB.OptionButton optCheapTab 
         Caption         =   "Color Control"
         Height          =   360
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optCheapTab 
         Caption         =   "Filters"
         Height          =   360
         Index           =   1
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1275
      End
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "Display"
      Height          =   3960
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   7110
      Begin VB.PictureBox picProgress 
         AutoRedraw      =   -1  'True
         FillColor       =   &H8000000D&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   3135
         ScaleHeight     =   1
         ScaleMode       =   0  'User
         ScaleWidth      =   99.216
         TabIndex        =   4
         ToolTipText     =   "Progress Bar"
         Top             =   3495
         Width           =   3855
      End
      Begin VB.OptionButton optOutput 
         Caption         =   "Output"
         Height          =   360
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3495
         Width           =   1500
      End
      Begin VB.OptionButton optOriginal 
         Caption         =   "Original"
         Height          =   360
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3495
         Value           =   -1  'True
         Width           =   1500
      End
      Begin VB.PictureBox picDisplay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3240
         Left            =   90
         Picture         =   "frmDemo.frx":500E
         ScaleHeight     =   212
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   457
         TabIndex        =   1
         Top             =   225
         Width           =   6915
      End
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   Image Processing demo for Planet-Source-Code
'   All code within is copyright (c) 2001 Kevin Gadd unless otherwise noted
'   Uses 'CapturePicture' by Benjamin Marty
'
'   This sample can be as much as 100 times faster when compiled, please try it compiled!
'   Also, all the effects in this sample work perfectly even with Array Bounds Checks off.
'   Feel free to turn on any of VB's compiler optimizations, and see how much faster this is :)
'   The processor is faster if you select 'Original' before processing.
'
'   Also, the processor combines effects using multiple passes across each row,
'   it could be MUCH faster if you specialized multiple versions of the loops
'   for each effect or combination of effects.
'
'   Enjoy!
'
'
'   Kevin Gadd a.k.a. 'Janus'
'   janusfury@citlink.net
'   UIN 130863581
'

Option Explicit

'   Used for accurate-to-the-millisecond timing
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'   Lookup tables
Private m_bytLookup(0 To 255, 0 To 255) As Byte

'   Image arrays
Private m_p32Original() As Pixel32
Private m_p32Output() As Pixel32
Private m_p32Doodad() As Pixel32

'   Picture size information
Private m_lngWidth As Long, m_lngHeight As Long
'   Progress bar data
Private m_sngProgress As Single, m_lngTimeElapsed As Long
'   Selected doodad
Private m_lngDoodad As Long
'   Timing information
Private m_lngStartTime As Long, m_lngEndTime As Long

'   Color control values
Private m_lngRedAdjust As Long, m_lngGreenAdjust As Long, m_lngBlueAdjust As Long
Private m_lngMinRed As Long, m_lngMinGreen As Long, m_lngMinBlue As Long
Private m_lngMaxRed As Long, m_lngMaxGreen As Long, m_lngMaxBlue As Long
'   Effect flags
Private m_booGrayscale As Boolean, m_booInvert As Boolean
'   For GetTimeMS
Private m_dblPerformanceFrequency As Double
'   Number of rows between each progress bar/screen update
Private Const c_lngUpdateDelay As Long = 10

Private Function GetTimeMS() As Long
Dim m_curTime As Currency
    Call QueryPerformanceCounter(m_curTime)
    GetTimeMS = CLng((CDbl(m_curTime) / m_dblPerformanceFrequency) * 1000)
End Function

'   Selects a doodad
Private Sub SelectDoodad(Doodad As Integer)
    picDoodad(m_lngDoodad).Cls
    picDoodad(m_lngDoodad).Refresh
    m_lngDoodad = Doodad
    picDoodad(m_lngDoodad).Cls
    picDoodad(m_lngDoodad).Line (0, 0)-(picDoodad(m_lngDoodad).ScaleWidth - 1, picDoodad(m_lngDoodad).ScaleHeight - 1), &HFFFFFF, B
    picDoodad(m_lngDoodad).Line (1, 1)-(picDoodad(m_lngDoodad).ScaleWidth - 2, picDoodad(m_lngDoodad).ScaleHeight - 2), &HA3A3A3, B
    picDoodad(m_lngDoodad).Line (2, 2)-(picDoodad(m_lngDoodad).ScaleWidth - 3, picDoodad(m_lngDoodad).ScaleHeight - 3), &H6D6D6D, B
    picDoodad(m_lngDoodad).Refresh
    DoEvents
    m_p32Doodad = GetPictureArrayInv(picDoodad(Doodad))
End Sub

'   Sets a pixel in the output buffer.
Private Sub SetBufferPixel(ByRef X As Long, ByRef Y As Long, ByVal Color As Long)
On Error Resume Next
    CopyMemory m_p32Output(X, Y), Color, 4
End Sub

'   Clips a long to two positive values, using boolean equasions.
Private Function ClipEx(ByVal Value As Long, ByRef Min As Long, ByRef Max As Long) As Byte
    Value = ((Value >= Min) And Value) Or ((Not (Value >= Min)) And Min)
    Value = ((Value <= Max) And Value) Or ((Not (Value <= Max)) And Max)
    ClipEx = Value
End Function

'   Clips a long to a byte, using boolean equasions.
Private Function ClipByte(ByVal Value As Long) As Byte
    Value = ((Value >= 0) And Value)
    Value = ((Value <= 255) And Value) Or ((Not (Value <= 255)) And 255)
    ClipByte = Value
End Function

'   Repaints the progress bar.
Private Sub RepaintProgressbar()
Dim m_strProgress As String
Dim m_lngTextX As Long, m_lngTextY As Long
    picProgress.Cls
    If m_sngProgress > 0 Then
        picProgress.Line (0, 0)-(m_sngProgress, 1), picProgress.FillColor, BF
    End If
    picProgress.ScaleMode = 3
    m_strProgress = Format(m_lngTimeElapsed, "####0000") + " milliseconds"
    m_lngTextX = (picProgress.ScaleWidth - picProgress.TextWidth(m_strProgress)) / 2
    m_lngTextY = (picProgress.ScaleHeight - picProgress.TextHeight(m_strProgress)) / 2
    picProgress.CurrentX = m_lngTextX + 1
    picProgress.CurrentY = m_lngTextY + 1
    picProgress.ForeColor = &H0
    picProgress.Print m_strProgress
    picProgress.CurrentX = m_lngTextX
    picProgress.CurrentY = m_lngTextY
    picProgress.ForeColor = &HFFFFFF
    picProgress.Print m_strProgress
    picProgress.ScaleMode = 0
    picProgress.ScaleWidth = 100
    picProgress.ScaleHeight = 1
    picProgress.Refresh
End Sub

'   Draws a doodad
Private Sub RepaintDoodad(ByVal X As Long, ByVal Y As Long, ByVal Doodad As Long)
On Error Resume Next
    ' Make sure they can see the 'effects' tab
    If optCheapTab(4).Value = False Then
        optCheapTab(4).Value = True
        optCheapTab_Click 4
        DoEvents
    End If
    m_lngStartTime = GetTimeMS
    If optEffects(0).Value = True Then ' Alpha Blend
        AlphaBlit m_p32Output, m_p32Doodad, X - (picDoodad(Doodad).ScaleWidth / 2), Y - (picDoodad(Doodad).ScaleHeight / 2), picDoodad(Doodad).ScaleWidth, picDoodad(Doodad).ScaleHeight, 0, 0, CLng(txtSourceOpacity.Text), CLng(txtDestOpacity.Text)
    ElseIf optEffects(1).Value = True Then ' Additive Blend
        AdditiveBlit m_p32Output, m_p32Doodad, X - (picDoodad(Doodad).ScaleWidth / 2), Y - (picDoodad(Doodad).ScaleHeight / 2), picDoodad(Doodad).ScaleWidth, picDoodad(Doodad).ScaleHeight, 0, 0, CLng(txtSourceOpacity.Text)
    ElseIf optEffects(2).Value = True Then ' Subtractive Blend
        SubtractiveBlit m_p32Output, m_p32Doodad, X - (picDoodad(Doodad).ScaleWidth / 2), Y - (picDoodad(Doodad).ScaleHeight / 2), picDoodad(Doodad).ScaleWidth, picDoodad(Doodad).ScaleHeight, 0, 0, CLng(txtSourceOpacity.Text)
    End If
    m_lngEndTime = GetTimeMS
    m_lngTimeElapsed = m_lngEndTime - m_lngStartTime
    RepaintProgressbar
    RepaintDisplay
End Sub

'   Copies a block of pixels, with alpha blending.
Private Sub AlphaBlit(ByRef OutputArray() As Pixel32, ByRef SourceArray() As Pixel32, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0, Optional SourceAlpha As Long = 255, Optional DestAlpha As Long = 0, Optional MaskColor As Long = -1)
On Error Resume Next
Dim m_lngX As Long, m_lngY As Long
Dim m_lngDestAlpha As Long, m_lngSourceAlpha As Long
Dim m_lngSourceWidth As Long, m_lngSourceHeight As Long
Dim m_lngDestWidth As Long, m_lngDestHeight As Long
Dim m_pxlSource As Pixel32, m_lngSource As Long
    ' Get some values
    m_lngDestAlpha = ClipByte(DestAlpha)
    m_lngSourceAlpha = ClipByte(SourceAlpha)
    m_lngDestWidth = UBound(OutputArray, 1)
    m_lngDestHeight = UBound(OutputArray, 2)
    m_lngSourceWidth = UBound(SourceArray, 1)
    m_lngSourceHeight = UBound(SourceArray, 2)
    ' Clip coordinates, in case you turn off array bounds checks
    If (X2 + Width) > (m_lngSourceWidth) Then
        Width = m_lngSourceWidth - X2
    End If
    If (Y2 + Height) > (m_lngSourceHeight) Then
        Height = m_lngSourceHeight - Y2
    End If
    If (X + Width) > (m_lngDestWidth) Then
        Width = m_lngDestWidth - X
    End If
    If (Y + Height) > (m_lngDestHeight) Then
        Height = m_lngDestHeight - Y
    End If
    If (X < 0) Then
        X2 = X2 + (-X)
        Width = Width + X
        X = 0
    End If
    If (Y < 0) Then
        Y2 = Y2 + (-Y)
        Height = Height + Y
        Y = 0
    End If
    If m_lngDestAlpha = 0 And m_lngSourceAlpha = 255 Then ' Special case: normal blit
        ' copy pixels
        For m_lngY = Y To (Y + (Height))
            For m_lngX = X To (X + (Width))
                ' Copy the source pixel to a Long
                CopyMemory m_lngSource, SourceArray((m_lngX - X) + X2, (m_lngY - Y) + Y2), 4
                ' Check to see if it's the mask color
                If (m_lngSource <> MaskColor) Then
                    CopyMemory m_p32Output(m_lngX, m_lngY), m_lngSource, 4
                End If
            Next m_lngX
        Next m_lngY
    ElseIf m_lngDestAlpha = 255 And m_lngSourceAlpha = 0 Then ' Special case: invisible blit
        Exit Sub
    Else
        ' Manipulate pixels
        For m_lngY = Y To (Y + (Height))
            For m_lngX = X To (X + (Width))
                ' Copy the source pixel to a Long and Pixel32 UDT
                CopyMemory m_pxlSource, SourceArray((m_lngX - X) + X2, (m_lngY - Y) + Y2), 4
                CopyMemory m_lngSource, m_pxlSource, 4
                ' Check to see if it's the mask color
                If (m_lngSource <> MaskColor) Then
                    ' Alpha blend
                    With m_p32Output(m_lngX, m_lngY)
                        ' Convert the destination color to long, so that no overflows will occur.
                        .Red = ClipByte(CLng(m_bytLookup(m_lngDestAlpha, .Red)) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Red))
                        .Green = ClipByte(CLng(m_bytLookup(m_lngDestAlpha, .Green)) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Green))
                        .Blue = ClipByte(CLng(m_bytLookup(m_lngDestAlpha, .Blue)) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Blue))
                    End With
                End If
            Next m_lngX
        Next m_lngY
    End If
End Sub

'   Copies a block of pixels, with additive blending.
Private Sub AdditiveBlit(ByRef OutputArray() As Pixel32, ByRef SourceArray() As Pixel32, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0, Optional SourceAlpha As Long = 255, Optional MaskColor As Long = -1)
On Error Resume Next
Dim m_lngX As Long, m_lngY As Long
Dim m_lngSourceAlpha As Long
Dim m_lngSourceWidth As Long, m_lngSourceHeight As Long
Dim m_lngDestWidth As Long, m_lngDestHeight As Long
Dim m_pxlSource As Pixel32, m_lngSource As Long
    ' Get some values
    m_lngSourceAlpha = ClipByte(SourceAlpha)
    m_lngDestWidth = UBound(OutputArray, 1)
    m_lngDestHeight = UBound(OutputArray, 2)
    m_lngSourceWidth = UBound(SourceArray, 1)
    m_lngSourceHeight = UBound(SourceArray, 2)
    ' Clip coordinates, in case you turn off array bounds checks
    If (X2 + Width) > (m_lngSourceWidth) Then
        Width = m_lngSourceWidth - X2
    End If
    If (Y2 + Height) > (m_lngSourceHeight) Then
        Height = m_lngSourceHeight - Y2
    End If
    If (X + Width) > (m_lngDestWidth) Then
        Width = m_lngDestWidth - X
    End If
    If (Y + Height) > (m_lngDestHeight) Then
        Height = m_lngDestHeight - Y
    End If
    If (X < 0) Then
        X2 = X2 + (-X)
        Width = Width + X
        X = 0
    End If
    If (Y < 0) Then
        Y2 = Y2 + (-Y)
        Height = Height + Y
        Y = 0
    End If
    ' Manipulate pixels
    For m_lngY = Y To (Y + (Height))
        For m_lngX = X To (X + (Width))
            ' Copy the source pixel to a Long and Pixel32 UDT
            CopyMemory m_pxlSource, SourceArray((m_lngX - X) + X2, (m_lngY - Y) + Y2), 4
            CopyMemory m_lngSource, m_pxlSource, 4
            ' Check to see if it's the mask color
            If (m_lngSource <> MaskColor) Then
                ' Alpha blend
                With m_p32Output(m_lngX, m_lngY)
                    ' Convert the destination color to long, so that no overflows will occur.
                    .Red = ClipByte(CLng(.Red) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Red))
                    .Green = ClipByte(CLng(.Green) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Green))
                    .Blue = ClipByte(CLng(.Blue) + m_bytLookup(m_lngSourceAlpha, m_pxlSource.Blue))
                End With
            End If
        Next m_lngX
    Next m_lngY
End Sub

'   Copies a block of pixels, with subtractive blending.
Private Sub SubtractiveBlit(ByRef OutputArray() As Pixel32, ByRef SourceArray() As Pixel32, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Optional ByVal X2 As Long = 0, Optional ByVal Y2 As Long = 0, Optional SourceAlpha As Long = 255, Optional MaskColor As Long = -1)
On Error Resume Next
Dim m_lngX As Long, m_lngY As Long
Dim m_lngSourceAlpha As Long
Dim m_lngSourceWidth As Long, m_lngSourceHeight As Long
Dim m_lngDestWidth As Long, m_lngDestHeight As Long
Dim m_pxlSource As Pixel32, m_lngSource As Long
    ' Get some values
    m_lngSourceAlpha = ClipByte(SourceAlpha)
    m_lngDestWidth = UBound(OutputArray, 1)
    m_lngDestHeight = UBound(OutputArray, 2)
    m_lngSourceWidth = UBound(SourceArray, 1)
    m_lngSourceHeight = UBound(SourceArray, 2)
    ' Clip coordinates, in case you turn off array bounds checks
    If (X2 + Width) > (m_lngSourceWidth) Then
        Width = m_lngSourceWidth - X2
    End If
    If (Y2 + Height) > (m_lngSourceHeight) Then
        Height = m_lngSourceHeight - Y2
    End If
    If (X + Width) > (m_lngDestWidth) Then
        Width = m_lngDestWidth - X
    End If
    If (Y + Height) > (m_lngDestHeight) Then
        Height = m_lngDestHeight - Y
    End If
    If (X < 0) Then
        X2 = X2 + (-X)
        Width = Width + X
        X = 0
    End If
    If (Y < 0) Then
        Y2 = Y2 + (-Y)
        Height = Height + Y
        Y = 0
    End If
    ' Manipulate pixels
    For m_lngY = Y To (Y + (Height))
        For m_lngX = X To (X + (Width))
            ' Copy the source pixel to a Long and Pixel32 UDT
            CopyMemory m_pxlSource, SourceArray((m_lngX - X) + X2, (m_lngY - Y) + Y2), 4
            CopyMemory m_lngSource, m_pxlSource, 4
            ' Check to see if it's the mask color
            If (m_lngSource <> MaskColor) Then
                ' Alpha blend
                With m_p32Output(m_lngX, m_lngY)
                    ' Convert the destination color to long, so that no overflows will occur.
                    .Red = ClipByte(CLng(.Red) - m_bytLookup(m_lngSourceAlpha, m_pxlSource.Red))
                    .Green = ClipByte(CLng(.Green) - m_bytLookup(m_lngSourceAlpha, m_pxlSource.Green))
                    .Blue = ClipByte(CLng(.Blue) - m_bytLookup(m_lngSourceAlpha, m_pxlSource.Blue))
                End With
            End If
        Next m_lngX
    Next m_lngY
End Sub

'   Processes the image.
Private Sub ProcessImage()
Dim m_lngX As Long, m_lngY As Long
Dim m_lngSource As Long, m_lngDest As Long
Dim m_lngBrightness As Long, m_sngProgressRatio As Single
Dim m_lngRed As Long, m_lngGreen As Long, m_lngBlue As Long, m_lngCount As Long
Dim m_lngUpdate As Long
    m_sngProgressRatio = CSng(100) / (CSng(m_lngHeight))
    ' Create a copy to work with
    m_p32Output() = m_p32Original()
    m_lngStartTime = GetTimeMS
    For m_lngY = 0 To m_lngHeight - 1
        ' Color control (simple)
        For m_lngX = 0 To m_lngWidth - 1
            With m_p32Output(m_lngX, m_lngY)
                .Red = ClipEx(.Red + m_lngRedAdjust, m_lngMinRed, m_lngMaxRed)
                .Green = ClipEx(.Green + m_lngGreenAdjust, m_lngMinGreen, m_lngMaxGreen)
                .Blue = ClipEx(.Blue + m_lngBlueAdjust, m_lngMinBlue, m_lngMaxBlue)
            End With
        Next m_lngX
        If m_booGrayscale = True Then
            ' Grayscale (simple)
            For m_lngX = 0 To m_lngWidth - 1
                With m_p32Output(m_lngX, m_lngY)
                    m_lngBrightness = (CLng(.Red) + CLng(.Green) + CLng(.Blue)) \ 3
                    .Red = m_lngBrightness
                    .Green = m_lngBrightness
                    .Blue = m_lngBrightness
                End With
            Next m_lngX
        End If
        If m_booInvert = True Then
            ' Invert (simple)
            For m_lngX = 0 To m_lngWidth - 1
                With m_p32Output(m_lngX, m_lngY)
                    .Red = 255 - .Red
                    .Green = 255 - .Green
                    .Blue = 255 - .Blue
                End With
            Next m_lngX
        End If
        If optFilters(1).Value = True Then
            ' Bi-linear blur (simple)
            For m_lngX = 0 To m_lngWidth - 1
                m_lngCount = 0: m_lngRed = 0: m_lngGreen = 0: m_lngBlue = 0
                If (m_lngX > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX - 1, m_lngY)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngY > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX, m_lngY - 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX < m_lngWidth - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX + 1, m_lngY)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngY < m_lngHeight - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX, m_lngY + 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                m_lngCount = m_lngCount + 1
                With m_p32Output(m_lngX, m_lngY)
                    .Red = ClipByte((.Red + m_lngRed) \ m_lngCount)
                    .Green = ClipByte((.Green + m_lngGreen) \ m_lngCount)
                    .Blue = ClipByte((.Blue + m_lngBlue) \ m_lngCount)
                End With
            Next m_lngX
        ElseIf optFilters(2).Value = True Then
            ' Tri-linear blur (simple)
            For m_lngX = 0 To m_lngWidth - 1
                m_lngCount = 0: m_lngRed = 0: m_lngGreen = 0: m_lngBlue = 0
                If (m_lngX > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX - 1, m_lngY)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX > 1) And (m_lngY > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX - 1, m_lngY - 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngY > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX, m_lngY - 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX < m_lngWidth - 1) And (m_lngY > 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX + 1, m_lngY - 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX > 1) And (m_lngY < m_lngHeight - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX - 1, m_lngY + 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX < m_lngWidth - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX + 1, m_lngY)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngX < m_lngWidth - 1) And (m_lngY < m_lngHeight - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX + 1, m_lngY + 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                If (m_lngY < m_lngHeight - 1) Then
                    m_lngCount = m_lngCount + 1
                    With m_p32Output(m_lngX, m_lngY + 1)
                        m_lngRed = m_lngRed + .Red
                        m_lngGreen = m_lngGreen + .Green
                        m_lngBlue = m_lngBlue + .Blue
                    End With
                End If
                m_lngCount = m_lngCount + 1
                With m_p32Output(m_lngX, m_lngY)
                    .Red = ClipByte((.Red + m_lngRed) \ m_lngCount)
                    .Green = ClipByte((.Green + m_lngGreen) \ m_lngCount)
                    .Blue = ClipByte((.Blue + m_lngBlue) \ m_lngCount)
                End With
            Next m_lngX
        End If
        m_lngUpdate = m_lngUpdate + 1
        If Not (m_lngUpdate < c_lngUpdateDelay) Then
            m_lngUpdate = 0
            ' Update progress-bar, and yield to windows. Prevents 'not responding'.
            m_sngProgress = CSng(m_lngY) * (m_sngProgressRatio)
            m_lngTimeElapsed = GetTimeMS - m_lngStartTime
            RepaintProgressbar
            ' If output tab is selected, update display so they can see it in action.
            If optOutput.Value = True Then RepaintDisplay
            DoEvents
        End If
    Next m_lngY
    m_lngEndTime = GetTimeMS
    m_lngTimeElapsed = m_lngEndTime - m_lngStartTime
    m_sngProgress = 100
    RepaintProgressbar
    optOutput.Value = True
    RepaintDisplay
End Sub

'   Repaints the display with the selected image.
Private Sub RepaintDisplay()
    If optOriginal.Value = True Then
        CopyPixelsToDC picDisplay.hDC, m_p32Original
    ElseIf optOutput.Value = True Then
        CopyPixelsToDC picDisplay.hDC, m_p32Output
    End If
End Sub

'   Begins processing.
Private Sub cmdProcess_Click()
On Error Resume Next
    ' Load settings
    m_lngRedAdjust = CLng(txtRedAdjust.Text): m_lngGreenAdjust = CLng(txtGreenAdjust.Text): m_lngBlueAdjust = CLng(txtBlueAdjust.Text)
    m_lngMinRed = ClipByte(CLng(txtRedMin.Text)): m_lngMinGreen = ClipByte(CLng(txtGreenMin.Text)): m_lngMinBlue = ClipByte(CLng(txtBlueMin.Text))
    m_lngMaxRed = ClipByte(CLng(txtRedMax.Text)): m_lngMaxGreen = ClipByte(CLng(txtGreenMax.Text)): m_lngMaxBlue = ClipByte(CLng(txtBlueMax.Text))
    m_booGrayscale = CBool(chkGrayscale.Value): m_booInvert = CBool(chkInvert.Value)
    ProcessImage
End Sub

'   Initializes program.
Private Sub Form_Load()
Dim m_sngAlpha As Single, m_lngAlpha As Long, m_lngValue As Long
Dim m_curPerformanceFrequency As Currency
    ' Create the lookup table.
    For m_lngAlpha = 0 To 255
        m_sngAlpha = CSng(m_lngAlpha) / 255
        For m_lngValue = 0 To 255
            m_bytLookup(m_lngAlpha, m_lngValue) = (CSng(m_lngValue) * m_sngAlpha)
        Next m_lngValue
    Next m_lngAlpha
    ' Retrieve the speed of the system performance counter
    Call QueryPerformanceFrequency(m_curPerformanceFrequency)
    m_dblPerformanceFrequency = CDbl(m_curPerformanceFrequency)
    ' Select the first doodad
    SelectDoodad 0
    ' Load the image's pixels into an array beforehand
    m_p32Original() = GetPictureArrayInv(picDisplay.Picture)
    ' Destroy the GDI bitmap for the picturebox
    Set picDisplay.Picture = Nothing
    ' Get width and height
    m_lngWidth = UBound(m_p32Original, 1) + 1
    m_lngHeight = UBound(m_p32Original, 2) + 1
    ' Create empty output
    ReDim m_p32Output(0 To m_lngWidth - 1, 0 To m_lngHeight - 1)
    ' Initialize settings
    m_lngRedAdjust = 0: m_lngGreenAdjust = 0: m_lngBlueAdjust = 0
    m_lngMinRed = 0: m_lngMinGreen = 0: m_lngMinBlue = 0
    m_lngMaxRed = 255: m_lngMaxGreen = 255: m_lngMaxBlue = 255
    m_booGrayscale = False: m_booInvert = False
End Sub

'   Show/hide the cheapo tabs
Private Sub optCheapTab_Click(Index As Integer)
On Error Resume Next
Dim m_lngTabs As Long
    For m_lngTabs = optCheapTab.LBound To optCheapTab.UBound
        picTab(m_lngTabs).Visible = optCheapTab(m_lngTabs).Value
    Next m_lngTabs
End Sub

Private Sub optOriginal_Click()
    ' Refresh the contents of the picturebox
    RepaintDisplay
End Sub

Private Sub optOutput_Click()
    ' Refresh the contents of the picturebox
    RepaintDisplay
End Sub

Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 1 Then
        If optOutput.Value = True Then
            ' Refresh the doodad
            RepaintDoodad X, Y, m_lngDoodad
        End If
    End If
End Sub

Private Sub picDisplay_Paint()
    ' Refresh the contents of the picturebox
    RepaintDisplay
End Sub

Private Sub picDoodad_Click(Index As Integer)
    SelectDoodad Index
End Sub

Private Sub txtOpacity_Change()
On Error Resume Next
    txtSourceOpacity.Text = CStr(ClipByte(CLng(txtOpacity.Text)))
    txtDestOpacity.Text = CStr(ClipByte(255 - CLng(txtOpacity.Text)))
End Sub
