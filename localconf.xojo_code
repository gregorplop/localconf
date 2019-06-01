#tag Class
Protected Class localconf
	#tag Method, Flags = &h0
		Function Close() As string
		  // returns empty string if OK, error message if error
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, optional password as string = "")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function create(file as FolderItem, password as string) As string
		  // returns empty string if OK, error message if error
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Open() As string
		  // returns empty string if OK, error message if error
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function preparePassword(plaintext as String) As string
		  // empty salt means use internal fixed salt via getSalt method
		  dim hash as MemoryBlock = Crypto.PBKDF2(salt , plaintext , 7 , 8 , Crypto.Algorithm.SHA512)
		  dim output as String
		  dim char as string
		  
		  for i as Integer = 0 to hash.Size - 1
		    char = str(hash.UInt8Value(i).ToHex(2))
		    if i mod 2 = 0 then char = char.Lowercase
		    output = output + char
		  next i
		  
		  return output  // should always be a 16 character string
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Salt() As string
		  // replace it with your own salt if needed
		  
		  Return "dEfAu1Ts@LT!"
		  
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private db As SQLiteDatabase
	#tag EndProperty

	#tag Property, Flags = &h21
		Private file As FolderItem
	#tag EndProperty


	#tag ViewBehavior
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Class
#tag EndClass
