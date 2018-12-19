'*********************************************************************
   'Declarations section of the module.
 '*********************************************************************
   Option Explicit
   Dim x As Integer
   Dim PadLength As Integer

 '=====================================================================
   'The following function will left pad a string with a specified
   'character. It accepts a base string which is to be left padded with
   'characters, a character to be used as the pad character, and a
   'length which specifies the total length of the padded result.
 '=====================================================================
   Function Lpad (MyValue$, MyPadCharacter$, MyPaddedLength%)

      Padlength = MyPaddedLength - Len(MyValue)
      Dim PadString As String
      For x = 1 To Padlength
         PadString = PadString & MyPadCharacter
      Next
      Lpad = PadString + MyValue

   End Function

 '=====================================================================
   'The following function will right pad a string with a specified
   'character. It accepts a base string which is to be right padded with
   'characters, a character to be used as the pad character, and a
   'length which specifies the total length of the padded result.
 '=====================================================================
   Function Rpad (MyValue$, MyPadCharacter$, MyPaddedLength%)

      Padlength = MyPaddedLength - Len(MyValue)
      Dim PadString As String
      For x = 1 To Padlength
         PadString = MyPadCharacter & PadString
      Next
      Rpad = MyValue + PadString

   End Function