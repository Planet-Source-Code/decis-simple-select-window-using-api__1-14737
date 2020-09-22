VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim lFoundWindow As Long
Dim lOK As Long
Dim lOK1 As Long
Dim X As Variant
    lFoundWindow = FindWindow(vbNullString, "MS-DOS Prompt")
    If lFoundWindow = 0 Then
        ' Did Not Find Window
        MsgBox ("Specified Application Is Not Running")
        ' You Could Use The Shell Command Here If You Wanted
        ' To Start It in This Instance.
    Else
        lOK = SetForegroundWindow(lFoundWindow)
        
        ' You may only need one of these lines, This needed both
        ' Due to the nature of the App I was Selecting so I will
        ' Leave ' both lines
            lOK1 = ShowWindow(lFoundWindow, 9)
            lOK1 = ShowWindow(lFoundWindow, 10)
        ' End
        
        lFoundWindow = 0
        lOK = 0
        lOK1 = 0
    End If
End Sub

