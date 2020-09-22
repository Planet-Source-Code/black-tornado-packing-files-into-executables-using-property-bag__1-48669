VERSION 5.00
Begin VB.Form frmReader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Property Bag - Reader"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

' Begin of code Reader.exe
' The code below, will attempt to read the EXTRA DATA attached with the copy of Reader.exe
' which is the patched version using Writer.exe. If you run this program without patching an
' error message will be displayed saying 'Bad Record Number' because there is not attached
' EXTRA DATA with the file, so the program can't read the EXTRA DATA

Dim Extracted_Bag As New PropertyBag ' Name of the extracted property bag
Dim Reading_Position As Long ' Start point of file reading
Dim Temp As Variant ' Variable to store the property bag contents from the Extracted_Bag
Dim RealContents() As Byte

On Error GoTo ReadError ' If there is an error in reading data, tell the user that!

Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1 ' Open the patched file (the file itself)
Get #1, LOF(1) - 3, Reading_Position ' Gets that start position of data and put it in Reading_Position

Seek #1, Reading_Position ' Set the file read position which is length of file - 3 characters
Get #1, , Temp            ' Temp = PropertyBag.Contents
RealContents = Temp       ' RealData = Temp converted into Bytes

Extracted_Bag.Contents = RealContents ' Put the contents to the bag
Close #1                  ' After we finished reading, we must close the file.

Dim Password As String    ' Password variable
Password = InputBox("Enter the password for the file") ' Ask user for password
If Password <> Extracted_Bag.ReadProperty("Password") Then
MsgBox "The password you entered is invalid!", vbCritical, "Invalid Password"
End
End If

' If Not Password <> Extracted_Bag.ReadProperty("Password") that means that the passwords match
' Display the good message

MsgBox Extracted_Bag.ReadProperty("Message")
End ' End of Application

ReadError: ' There is an error in reading data from file
           MsgBox "Error during data read, data may be empty or it is invalid", vbCritical
End
End Sub
