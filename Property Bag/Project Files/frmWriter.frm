VERSION 5.00
Begin VB.Form frmWriter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Property Bag - Writer"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmWriter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "frmWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
' Begin of code Writer.exe
' The code below, will attempt to mix Reader.exe (PropertyBag reader) with A property bag contents.
' In this example you will see the words EXTRA DATA, these 2 word means Property Bag Contents
' which is really an EXTRA DATA.
On Error Resume Next ' Resume code even if error happened, I have typed Resume because we will delete
' the file (Test.exe) if exists. But if it is not exists an error will happen
' So we will say 'On Error GoTo HaveError' when we reach the level of code after
' Kill statement, because now every error will happen after the Kill statement
' Its reason is error in opening file (Reader.exe) or in writing test file (Test.exe)
Dim EXE_File As String ' The compiled file path
Dim Writing_Position As Long ' This variable stores the position of writing EXTRA CODE!
Dim MyBag As New PropertyBag ' Make a new property bag to put in the EXE
Dim Temp As Variant ' This value will store the contents of MyBag, It is a variant because there is
' no specific type of Property Bag contents, it may by picture, string, integer...
' Now we will write some property items to the property bag, because we will retrieve the contents in the other
' program (Reader). And if the property doesn't contains data we will be not able to know if the program (Reader)
' has really read the data.
MyBag.WriteProperty "Message", InputBox("Enter a good message") ' A Message displayed if the password is OK!
MyBag.WriteProperty "Password", InputBox("Enter the protection password")   ' A test password
' If you want to add your own property just type:
' MyBag.WriteProperty "Your Property", "Property Value"
' Now, we will make a copy of Reader.exe to a new file named Test.exe
' this copy will read the EXTRA DATA. But there is no EXTRA DATA in it so
' we will open the file as binary and then add our bag contents
EXE_File = InputBox("Enter the compiled Executable path:", "Compiled File Patch", "C:\Test.exe")
Kill EXE_File ' Kill the compiled file if it is already exists
FileCopy App.Path & "\Reader.exe", EXE_File ' Make a copy of Reader.exe (Template) to EXE_File

' Now, we are going to write the REAL CODE!!!
On Error GoTo HaveError ' If there is an error, then its a writing error
Open EXE_File For Binary As #1 ' Open template file for applying patch in binary mode

Writing_Position = LOF(1) ' Writing_Position=Length Of File (1) which is the file we opened with Open
' statement (EXE_File), So Writing_Position = Size Of Target File (EXE_File) = End of original file
Temp = MyBag.Contents ' Copy the contents of property bag into Temp variable
Seek #1, LOF(1) ' Set position of file writing to the end of original file (File without EXTRA DATA)
Put #1, , Temp ' Put the contents of property bag to Test.exe file
Put #1, , Writing_Position ' At last, we must put the original file size (length)

Close #1 ' Close the file
' If operation is done without any errors, that means that the file has been written successfully
MsgBox "Testing file has been wrote successfully, press 'OK' to launch the file", vbInformation, "Congratulations!"
Shell EXE_File, vbNormalFocus ' Run the patched file
End ' End the program
HaveError: ' Sorry, we have an error

MsgBox "Error during file compilation", vbCritical, "Error compiling file"
' End of code Writer.exe
End ' End the program
End Sub
