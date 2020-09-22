VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visual Effects #3 - AlphaBlending (by MrHippo"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   ForeColor       =   &H8000000F&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   337
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.Slider barAmount 
      Height          =   2235
      Left            =   4440
      TabIndex        =   2
      Top             =   2520
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3942
      _Version        =   327682
      Orientation     =   1
      LargeChange     =   51
      Max             =   255
      SelStart        =   255
      TickStyle       =   2
      TickFrequency   =   51
      Value           =   255
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   240
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   270
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2520
      Width           =   4050
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   240
      Picture         =   "frmMain.frx":1F4D
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   270
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   4050
   End
   Begin VB.Label Label4 
      Caption         =   "Enjoy, Sveinn R. Sigurdsson, computer engineer at Tal Telecommunications, Iceland (MrHippo)"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label3 
      Caption         =   "By voting on planet-source-code.com you incourage me, and other programmers in providing you with various coding solutions."
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "For any assistance on this or other image related topic, please e-mail to depill2000@hotmail.com"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "Please vote and look out for more samples in the Visual Effects sample series on www.planet-source-code.com."
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4575
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "AlphaBlending was introduced in Windows 98. The function allows you to change the transparency of an image. "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ********************************************
' ** Purpose: Alpha Blending Pictures
' ** Website : www.svenni.com
' ** Programmer : Sveinn R. Sigurdsson
' ** e-mail : depill2000@hotmail.com
' ********************************************
Option Explicit

' API DECLARATION [ ALPHA BLEND FUNCTION ]
Private Declare Function AlphaBlend Lib "msimg32" ( _
ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

' API DECLARATIONS [ COPY MEMORY FUNCTION ]
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

' TYPE STRUCTURES
Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Private Sub barAmount_Scroll()
    ' Procedure Scope Declarations
    Dim tProperties As typeBlendProperties
    Dim lngBlend As Long
        ' Clear the destination picture
    picDestination.Cls
    tProperties.tBlendAmount = 255 - barAmount
    ' Call the 'CopyMemory' with the specified parameters
    CopyMemory lngBlend, tProperties, 4
    ' Blend the pictures together and show them at the specified
    ' location in the specified picture box.
    AlphaBlend picDestination.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, _
    picSource.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, lngBlend
    ' Refresh the picture box with the new image
    picDestination.Refresh
End Sub

