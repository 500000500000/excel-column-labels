Module Program
	'Input a column header string and return a number.
	Function ConvertColumnLabelToNumber(str As String)

		'Declare the return value.
		Dim numToReturn As UInteger

		'Make sure the input string is in uppercase.
		str = UCase(str)

		'Get the length of the string.
		Dim length As UInteger
		length = Len(str)

		If length > 1 Then

			'Split the string into all but the last character, and the last character.
			Dim allButLast As String
			Dim last As String
			allButLast = Left(str, length - 1)
			last = Right(str, 1)

			'Make two recursive calls.
			Dim numAllButLast As UInteger
			Dim numLast As UInteger
			numAllButLast = ConvertColumnLabelToNumber(allButLast)
			numLast = ConvertColumnLabelToNumber(last)

			'Set the output.
			numToReturn = 26 * numAllButLast + numLast

		Else

			'Set the output.
			numToReturn = Asc(str) - 64

		End If

		ConvertColumnLabelToNumber = numToReturn

	End Function

	Function ConvertColumnNumberToLabel(num As UInteger) As String

		Dim strToReturn As String

		If num > 26 Then

			'Divide.
			Dim q As UInteger
			Dim r As UInteger
			q = num \ 26
			r = num Mod 26

			'Get left and right part of string to return.
			Dim left As String
			Dim right As String
			If r > 0 Then
				left = ConvertColumnNumberToLabel(q)
				right = ConvertColumnNumberToLabel(r)
			Else
				left = ConvertColumnNumberToLabel(q - 1)
				right = "Z"
			End If

			'Combine left and right.
			strToReturn = left + right

		Else

			If num >= 1 And num <= 26 Then

				Dim alphabet As String
				alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
				strToReturn = Mid(alphabet, num, 1)

			Else
				strToReturn = "*undefined*"
			End If

		End If

		ConvertColumnNumberToLabel = strToReturn

	End Function


	Sub Main(args As String())
		Console.WriteLine("Convert each natural number to an excel column string and back")

		For n As UInteger = 1 To 1000 Step 1
			'Convert n to an excel string s.
			Dim s As String
			s = ConvertColumnNumberToLabel(n)

			'Convert s back to a number. 
			Dim check As UInteger
			check = ConvertColumnLabelToNumber(s)

			'Write the number, the string, and the check.
			Console.WriteLine(n & " " & s & " " & check)
		Next

		'Large Example
		Dim nLarge As UInteger
		nLarge = UInteger.MaxValue
		Dim sLarge As String
		sLarge = ConvertColumnNumberToLabel(nLarge)
		Dim checkLarge As UInteger
		checkLarge = ConvertColumnLabelToNumber(sLarge)
		Console.WriteLine(nLarge & " " & sLarge & " " & checkLarge)

		'Error example
		Dim nError = 0
		Dim sError As String = ConvertColumnNumberToLabel(nError)
		Console.WriteLine(nError & " " & sError)

	End Sub
End Module