VERSION 5.00
Begin VB.Form frmProtectionDemo 
   Caption         =   "Image Capture Protection demo"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Protection image"
      Height          =   2775
      Left            =   3720
      TabIndex        =   5
      Top             =   2760
      Width           =   3495
      Begin VB.PictureBox Picture3 
         Height          =   2175
         Left            =   240
         Picture         =   "frmImageProtect.frx":0000
         ScaleHeight     =   2115
         ScaleWidth      =   2835
         TabIndex        =   8
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "This image is shown when captured."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display image"
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   3495
      Begin VB.PictureBox Picture2 
         Height          =   2175
         Left            =   240
         Picture         =   "frmImageProtect.frx":14442
         ScaleHeight     =   2115
         ScaleWidth      =   2835
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "This image will be displayed."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Capture protected image"
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox Picture1 
         Height          =   2175
         Left            =   120
         ScaleHeight     =   2115
         ScaleWidth      =   2835
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "This picture box is protected against screen capturing. When you capture this image, you should see the protection image below."
         Height          =   1695
         Left            =   3120
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
   End
End
Attribute VB_Name = "frmProtectionDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API Declarations
Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

' Constants
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

' Variables
Private blCancel As Boolean

Private Sub StartProtection()
    
    Dim Width As Long, Height As Long
    
    blCancel = False
    
    ' Get the width and height of the destination picture
    Width = Picture1.ScaleWidth / Screen.TwipsPerPixelX
    Height = Picture1.ScaleHeight / Screen.TwipsPerPixelY
        
    Do
                
        ' Show the display image
        Call BitBlt(Picture1.hDC, 0, 0, Width, Height, Picture2.hDC, 0, 0, SRCCOPY)
        
        ' Sleep 200 milliseconds without handling windows messages.
        ' The duration may be decreased or increased. (40 ms = 1 frame at 25 frames/s)
        SleepEx 200, True
        
        ' Draw the protection image
        Call BitBlt(Picture1.hDC, 0, 0, Width, Height, Picture3.hDC, 0, 0, SRCCOPY)
                
        ' Handle windows messages. At this point the screen capture will take place.
        ' Because we have drawn the protection image, this image will appear in the
        ' screen capture.
        
        DoEvents
        
        ' Now repeat this process until the user has clicked the 'Exit' button.
        ' Because the protection image is shown a fraction of a second, it won't
        ' be visible to the human eye. However, sometimes the protected picture box
        ' may 'flicker' a little bit.
        
    Loop Until blCancel = True  ' If the 'Exit' button was clicked, exit loop
                
    ' Refresh the picture box
    Picture1.Refresh

    End
    
End Sub

Private Sub btnExit_Click()
    
    ' Exit program
    blCancel = True

End Sub

Private Sub Form_Load()
    
    ' Show form
    Me.Show
    DoEvents
    
    ' Start the protection
    Call StartProtection
    
End Sub
