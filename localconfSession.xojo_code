#tag Class
Protected Class localconfSession
	#tag Method, Flags = &h0
		Sub Constructor(file as FolderItem, optional password as string = "")
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function create(file as FolderItem, password as string) As string
		  // returns empty string if OK, error message if error
		  
		  if IsNull(file) then return "Configuration file is null!"
		  if file.Exists then Return "Configuration file already exists!"
		  if password.Trim = "" then return "No password set for new configuration file!"
		  
		  dim db as new SQLiteDatabase
		  
		  db.DatabaseFile = file
		  db.Password = preparePassword(password)
		  
		  dim outcome as Boolean 
		  dim errorMsg as string
		  
		  outcome = db.CreateDatabaseFile
		  if outcome = false then Return "Error creating new configuration file: " + db.ErrorMessage
		  
		  outcome = db.Connect
		  if outcome = false then
		    errorMsg = db.ErrorMessage
		    db.DatabaseFile.Delete
		    Return "Error opening newly created configuration file: " + errorMsg
		  end if
		  
		  dim CREATETABLE as string = "CREATE TABLE localconf (objidx INTEGER PRIMARY KEY AUTOINCREMENT , application VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , user VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , section VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , key VARCHAR NOT NULL , value VARCHAR , comments VARCHAR)"
		  
		  db.SQLExecute(CREATETABLE)
		  if db.Error then
		    errorMsg = db.ErrorMessage
		    db.Close
		    db.DatabaseFile.Delete
		    Return "Error initializing newly created configuration file: " + errorMsg
		  end if
		  
		  db.Close
		  Return ""  // success
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function LastError() As string
		  Return mLastError
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function preparePassword(plaintext as String) As string
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

	#tag Method, Flags = &h0
		Function Read(recordObject as localconfRecord) As localconfRecord
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ReadAll() As localconfRecord()
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function Salt() As string
		  // replace it with your own salt if needed
		  
		  Return "dEfAu1Ts@LT!"
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Write(recordObject as localconfRecord) As localconfRecord
		  
		End Function
	#tag EndMethod


	#tag Property, Flags = &h21
		Private file As FolderItem
	#tag EndProperty

	#tag Property, Flags = &h21
		Private mLastError As string
	#tag EndProperty

	#tag Property, Flags = &h21
		Private passwd As String
	#tag EndProperty


	#tag Constant, Name = GlobalName, Type = String, Dynamic = False, Default = \"GLOBAL", Scope = Public
	#tag EndConstant


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
