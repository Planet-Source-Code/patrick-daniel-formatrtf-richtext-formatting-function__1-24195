VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FormatRTF"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3525
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   6218
      _Version        =   393217
      BackColor       =   -2147483633
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":000C
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert"
      Height          =   345
      Left            =   6780
      TabIndex        =   2
      Top             =   3270
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   2865
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   8385
   End
   Begin VB.Label Label2 
      Caption         =   "Converted Richtext:"
      Height          =   225
      Left            =   30
      TabIndex        =   4
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Plain Text:"
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   
   RichTextBox1.Text = Text1.Text
   
   FormatRTF RichTextBox1
   
End Sub

Private Sub Form_Load()
   
   Text1.Text = "<ALIGN=Center>This function takes <B>multiple tags <COLOR=" & vbRed & ">and</B> uses the built" & vbCrLf & _
                  "in sel functions</COLOR> to <U>convert them.  You can</U> use as many tags as you would like" & vbCrLf & _
                  "in <I><B><U><STRIKE>any order you want</STRIKE></U></B></I>.</ALIGN>" & vbCrLf & vbCrLf & _
                  "It Supports:" & vbCrLf & _
                  "<SIZE=14>Sizes</SIZE>, <FONT=Tahoma>Font Names</FONT>, " & _
                     "<B>Bold</B>, <U>Underline</U>, <U>Italics</U>, <STRIKE>Strikethru</STRIKE>, " & _
                     "<COLOR=" & vbBlue & ">Colors</COLOR>" & vbCrLf & _
                     "<BULLET>Bullets</BULLET><ALIGN=Center> And Alignment</ALIGN>" & vbCrLf & vbCrLf & _
                     "<ALIGN=Center><B>I use this mostly in places where" & vbCrLf & _
                     "I need bottomless, autosizing, formatted text such as in Email Headers like in Outlook and Outlook Express!!</B></ALIGN>" & vbCrLf & vbCrLf & _
                     "<B>From:</B> " & "patrick1@mediaone.net" & Space(8) & "<B>To:</B> " & "user@email.com" & vbCrLf & _
                        "<B>Subject:</B> " & "This would be a good example of using FormatRTF for email headers"
   
   Command1_Click
   
End Sub

Public Sub FormatRTF(ByRef txtRTF As RichTextBox)
On Error Resume Next

   Dim i As Integer
   
   Dim strTags() As Variant
   
   Dim iLength As Integer
   Dim strValue As String
   Dim iStart As Integer
   Dim iEnd As Integer
   Dim strStartTag As String
   Dim strEndTag As String
   Dim iStartTag As Integer
   Dim iEndTag As Integer
   Dim iLenST As Integer
   Dim iCount As Integer
   
   With txtRTF
            
      '-- Look For Font Tags
      strStartTag = "<FONT="
      strEndTag = "</FONT>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0

         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .SelStart = iEnd - 1
         .SelLength = iEndTag
         .SelText = ""
         
         .SelStart = iStart - 1
         .SelLength = iStartTag
         .SelText = ""

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelFontName = strValue
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop

      '-- Look For Font Size Tags
      strStartTag = "<SIZE="
      strEndTag = "</SIZE>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0

         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelFontSize = CInt(strValue)
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- Font Colors
      strStartTag = "<COLOR="
      strEndTag = "</COLOR>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0
      
         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         .SelColor = CLng(strValue)
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- Alignment
      strStartTag = "<ALIGN="
      strEndTag = "</ALIGN>"
      iStart = InStr(1, .Text, strStartTag)
      iCount = 0
      
      Do While iStart > 0
      
         iLenST = Len(strStartTag)
         iEndTag = Len(strEndTag)
         
         iStart = InStr(1, .Text, strStartTag)
         
         If iStart = 0 Then Exit Do

         iEnd = InStr(iStart, .Text, strEndTag)

         strValue = Mid(.Text, (iStart + iLenST), InStr((iStart + iLenST), .Text, ">") - (iStart + iLenST))

         iStartTag = Len(strStartTag & strValue & ">")

         iLength = iEnd - iStartTag - iStart
         
         .TextRTF = Replace(.TextRTF, strStartTag & strValue & ">", "", , 1)
         .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)

         .SelStart = iStart - 1
         .SelLength = iLength
         
         Select Case UCase(strValue)
            Case "LEFT"
               .SelAlignment = rtfLeft
            Case "RIGHT"
               .SelAlignment = rtfRight
            Case "CENTER"
               .SelAlignment = rtfCenter
         End Select
         
         iCount = iCount + 1
         If iCount > 100 Then Exit Do
         
      Loop
      
      '-- All Others
      
      ReDim strTags(4)
      strTags(0) = "B"
      strTags(1) = "U"
      strTags(2) = "I"
      strTags(3) = "STRIKE"
      strTags(4) = "BULLET"
      
      For i = LBound(strTags) To UBound(strTags)
      
         strStartTag = "<" & strTags(i) & ">"
         strEndTag = "</" & strTags(i) & ">"
         iStart = InStr(1, .Text, strStartTag)
         iCount = 0
         
         Do While iStart > 0
   
            iLenST = Len(strStartTag)
            iEndTag = Len(strEndTag)
            
            iStart = InStr(1, .Text, strStartTag)
            
            If iStart = 0 Then Exit Do
   
            iEnd = InStr(iStart, .Text, strEndTag)
   
            iStartTag = Len(strStartTag)
   
            iLength = iEnd - iStartTag - iStart
            
            .TextRTF = Replace(.TextRTF, strStartTag, "", , 1)
            .TextRTF = Replace(.TextRTF, strEndTag, "", , 1)
   
            .SelStart = iStart - 1
            .SelLength = iLength
            
            If i = 0 Then
               .SelBold = True
            ElseIf i = 1 Then
               .SelUnderline = True
            ElseIf i = 2 Then
               .SelItalic = True
            ElseIf i = 3 Then
               .SelStrikeThru = True
            ElseIf i = 4 Then
               .SelBullet = True
            End If
            
            iCount = iCount + 1
            If iCount > 100 Then Exit Do
            
         Loop
            
      Next
        
      .SelStart = 0
        
   End With
   
End Sub

