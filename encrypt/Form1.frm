VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   2175
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   6615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt Text"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt Text"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   written by paolo parungao
'   july 6, 2004
'
'
'   this code is free. you can reuse it, destroy it, have fun with it,
'   learn from it, whatever. just email me at parungaopb@hondamakati.com.ph
'   if you do something about it.

Option Explicit

'the logic of this encryption/decryption program, it divides the string every
'second character then concatenates it to a single string before converting the
'characters to ascii characters. it encrypts/ decrypts so fast...
    
    Function Encrypt(ByVal source As String) As String
        Dim StrLenght As Integer
        Dim ctr As Integer
        Dim StrF As String
        Dim StrG As String
        Dim OutText As String

        For ctr = 1 To Len(source) Step 2
            StrF = Mid(source, ctr, 1)      'this will get the first string
            StrG = StrG & Mid(source, ctr + 1, 1)   'this will get all the second string
            OutText = OutText & StrF        'then concatenate it to a single word
        Next ctr

        OutText = OutText & StrG
        
        For ctr = 1 To Len(OutText)
            StrF = Mid(OutText, ctr, 1)     'then get the first concatenated character
            StrLenght = (255 - Asc(StrF))   '255 is the largest possible ascii character in the set
            Encrypt = Encrypt & Chr(StrLenght)  'use it to produce a cipher text
        Next ctr

    End Function

'before proceeding the decryption method, we have to decrypt the cipher text first
    
    Function Decrypt(ByVal Encrypted As String) As String
        Dim StrLenght As Integer
        Dim ctr As Integer
        Dim StrF As String
        Dim StrG As String
        Dim OutText As String
        Dim length As Double
        Dim string1 As String
        Dim string2 As String


        For ctr = 1 To Len(Encrypted)       'just reverse the process. first
            StrF = Mid(Encrypted, ctr, 1)   'get the first character and subtract it to the largest
            StrLenght = (255 - Asc(StrF))   'possible ascii character(255) to get the equivalent decipher character
            OutText = OutText & Chr(StrLenght)  'then concatenate it
        Next ctr


        length = Len(OutText) / 2           'this method will then join the broken
        If (CDbl(length) - CInt(length)) <> 0 Then length = CInt(length)
        string1 = Mid$(OutText, 1, length)
        string2 = Mid$(OutText, length + 1, Len(OutText))
        For ctr = 1 To length
            StrF = Mid(string1, ctr, 1)
            If Len(string2) >= ctr Then
                StrG = Mid(string2, ctr, 1)
            End If
            Decrypt = Decrypt & StrF & StrG 'then concatenate the characters to produce the original message
            StrF = ""               'reset our variable so that the code can reuse it again
            StrG = ""
        Next ctr

    End Function

Private Sub Command1_Click()
Text2.Text = Encrypt(Text1.Text)
End Sub

Private Sub Command2_Click()
Text1.Text = Decrypt(Text2.Text)
End Sub

'enjoy and learn from this simple but powerful code
