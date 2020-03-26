REM declare variables
 Option Explicit
  Dim numInput
  Dim numWorkingValue
  Dim arrExponentsFound
  REM initialize as array
  arrExponentsFound = Array()
  Dim intExponent
  Dim strMessage
  Dim i
  Dim boolValidated

REM function to convert Log base e to base 2
 Function LogBase2(x)
  REM check if value of X > 0 to prevent runtime error
  If(x > 0) Then
    LogBase2 = Log(x) / Log(2)
  Else
    REM just return 0
    LogBase2 = 0
  End If
 End Function

REM function to add item to array
REM while preserving existing values
 Function AddItem(arr, val)
  ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
  AddItem = arr
 End Function

REM get input, check if not a number or not integer
REM and make them re-enter until we get an integer
 Do
  numInput = Inputbox("enter a whole number (integer)","Input")
  REM validate input, default is false
  boolValidated = False
   REM check if numeric
  If (IsNumeric(numInput)) Then
     REM also check if > 0
     REM also check if integer (not decimal) using 
     REM technique absolute value of difference = 0
    If (numInput > 0 AND Abs(numInput - Int(numInput)) = 0) Then
      boolValidated = True
    End If
  End If
 Loop While (Not boolValidated)

REM assign initial value of numInput to numWorkingValue
 numWorkingValue = numInput


REM loop until we've exhausted all possibilities
REM update arrExponentsFound each time using result of Int(LogBase2(numWorkingValue))
 Do
  intExponent = Int(LogBase2(numWorkingValue))
  arrExponentsFound = AddItem(arrExponentsFound, intExponent )
   REM need to subtract 2^? from numWorkingValue
  numWorkingValue = numWorkingValue - 2^intExponent
  'wscript.echo "remainder: " & numWorkingValue
 Loop While numWorkingValue > 0


REM create result message by looping through arrExponentsFound
 For i = 0 to Ubound(arrExponentsFound)
  REM trick to only add comma before 2nd exponent value
  If i > 0 Then
   strMessage = strMessage & ", "
  End If

  strMessage = strMessage & arrExponentsFound(i)
 Next

REM print final message
 wscript.echo "Exponents List: " & strMessage



