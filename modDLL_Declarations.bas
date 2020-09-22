Attribute VB_Name = "modDLL_Declarations"
Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Sub vbOutw Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Declare Function vbInp Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer
Declare Function vbInpw Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer



