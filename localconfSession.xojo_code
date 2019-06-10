#tag Class
Protected Class localconfSession
	#tag Method, Flags = &h0
		Function Append2Array(recordObject as localconfRecord) As localconfRecord
		  // only inserts records with the (application/user/section/key) combination
		  // when used repeatedy for the same combination, it effectively creates an array which can be read using ReadArray
		  
		  
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Sub Constructor(initfile as FolderItem, password as string)
		  // all localconf settings files are password-protected
		  // the actual password is calculated using the user-supplied password: only localconfSession is meant to access the config file
		  
		  // returns empty string if OK, error message if error
		  // although it's called a session, you'll notice that it only connects to the db file when needed. it does not keep the actual session to the database constantly open
		  
		  mLastError = ""
		  
		  if IsNull(initfile) then 
		    mLastError = "Configuration file is null!"
		  ElseIf initfile.Exists = false then
		    mLastError = "Configuration file does not exist!"
		  ElseIf password.Trim = "" Then
		    mLastError = "No password for configuration file!"
		  end if
		  if mLastError <> "" then Return  // error in init parameters
		  
		  dim db as new SQLiteDatabase
		  
		  db.DatabaseFile = initfile
		  db.EncryptionKey = preparePassword(password)
		  
		  dim outcome as Boolean = db.Connect
		  
		  if outcome = false then
		    mLastError = "Could not open settings file: Corrupt or wrong password"
		    Return  // fail
		  end if
		  
		  // connected ok
		  
		  file = initfile  // keep it there for reopeing the db -- localconfSession does not keep an open connection to the database file
		  passwd = EncodeBase64(password , 0)  // don't keep the plaintext password in memory
		  db.Close
		  
		  Return
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function Count(recordObject as localconfRecord) As integer
		  // counts records that match the (application/user/section/key) criterion
		  
		  mLastError = ""
		  
		  if validateCachedParams = false then
		    mLastError = "Localconf session no longer valid, please restart!"
		  ElseIf IsNull(recordObject) then
		    mLastError = "Invalid search parameters"
		  ElseIf recordObject.key.Trim = "" then
		    mLastError = "No key entered!"
		  end if
		  if mLastError <> "" then return -1
		  
		  dim db as new SQLiteDatabase
		  db.DatabaseFile = file
		  db.EncryptionKey = preparePassword(DecodeBase64(passwd))
		  
		  if db.Connect = false then 
		    mLastError = "Error accessing settings file: " + db.ErrorMessage
		    Return -1
		  end if
		  
		  dim WHERE as string
		  WHERE = "application = " + if(recordObject.application.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.application.Trim.Uppercase + "'") + " AND "
		  WHERE = WHERE + "user = " + if(recordObject.user.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.user.Trim.Uppercase + "'") + " AND "
		  WHERE = WHERE + "section = " + if(recordObject.section.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.section.Trim.Uppercase + "'") + " AND "
		  WHERE = WHERE + "key = '" + recordObject.key.Trim.Uppercase + "'"
		  
		  dim countdata as RecordSet = db.SQLSelect("SELECT COUNT(*) FROM localconf WHERE " + WHERE + " ORDER BY objidx ASC")
		  
		  if db.error = true then 
		    mLastError = "Error querying settings file: " + db.ErrorMessage
		    Return -1
		  end if
		  
		  Return countdata.IdxField(1).IntegerValue
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Shared Function create(file as FolderItem, password as string) As string
		  // returns empty string if OK, error message if error
		  
		  if IsNull(file) then return "Configuration file is null!"
		  if file.Exists then Return "Configuration file already exists!"
		  if password.Trim = "" then return "No password set for new configuration file!"
		  
		  dim db as new SQLiteDatabase
		  
		  db.DatabaseFile = file
		  db.EncryptionKey = preparePassword(password)
		  
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
		  
		  dim CREATETABLE as string = "CREATE TABLE localconf (objidx INTEGER PRIMARY KEY AUTOINCREMENT , application VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , user VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , section VARCHAR NOT NULL DEFAULT '" + GlobalName + "' , key VARCHAR NOT NULL , value VARCHAR , comment VARCHAR)"
		  db.SQLExecute(CREATETABLE)
		  
		  dim INSERTINITRECORD as String = "INSERT INTO localconf (section , key , value , comment) VALUES ('LOCALCONF' , 'INITSTAMP' , '" + date(new date).SQLDateTime + "' , 'Automatically generated by localconf')"
		  db.SQLExecute(INSERTINITRECORD)
		  
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
		Function Delete(objidx as Integer) As Boolean
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function FreeQuery(optional WHERE as String = "TRUE") As localconfRecord()
		  // if error then output is 1-element array with .error = true and .errorMessage hold reason for error
		  // if localconf file is empty then output is an array having UBound = -1
		  
		  dim dump(-1) as localconfRecord
		  
		  if validateCachedParams = false then
		    dump.Append new localconfRecord("Localconf session no longer valid, please restart!")
		    Return dump
		  end if
		  
		  dim db as new SQLiteDatabase
		  db.DatabaseFile = file
		  db.EncryptionKey = preparePassword(DecodeBase64(passwd))
		  
		  if db.Connect = false then 
		    dump.Append new localconfRecord("Error accessing settings file: " + db.ErrorMessage)
		    Return dump
		  end if
		  
		  if WHERE.Trim = "" then WHERE = "TRUE"
		  
		  dim dumpdata as RecordSet = db.SQLSelect("SELECT * FROM localconf WHERE " + WHERE + " ORDER BY objidx ASC")
		  
		  if db.error = true then 
		    dump.Append new localconfRecord("Error querying settings file: " + db.ErrorMessage)
		    Return dump
		  end if
		  
		  dim record as localconfRecord
		  
		  while not dumpdata.EOF
		    record = new localconfRecord(True)
		    
		    record.objidx = dumpdata.Field("objidx").IntegerValue
		    record.application = dumpdata.Field("application").StringValue
		    record.user = dumpdata.Field("user").StringValue
		    record.section = dumpdata.Field("section").StringValue
		    record.key = dumpdata.Field("key").StringValue
		    
		    record.value = dumpdata.Field("value").StringValue
		    record.comment = dumpdata.Field("comment").StringValue
		    
		    dump.Append record
		    
		    dumpdata.MoveNext
		  wend
		  
		  Return dump
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
		Function ReadArray(recordObject as localconfRecord) As localconfRecord()
		  // reads all elements that matche the (application/user/section/key) criterion
		  // objidx is ignored if exists
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function ReadSingle(recordObject as localconfRecord) As localconfRecord
		  // reads the first element that matches the (application/user/section/key) OR (objidx) criterion
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Shared Function Salt() As string
		  // replace it with your own salt if needed
		  
		  Return "dEfAu1Ts@LT!"
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h0
		Function Upsert(recordObject as localconfRecord) As localconfRecord
		  // match criterion is either the (application/user/section/key) combination or objidx (only update)
		  // objidx holds priority
		  
		  if validateCachedParams = false then
		    Return new localconfRecord("Localconf session no longer valid, please restart!")
		  ElseIf IsNull(recordObject) then
		    Return new localconfRecord("Invalid search parameters")
		  end if
		  
		  dim db as new SQLiteDatabase
		  db.DatabaseFile = file
		  db.EncryptionKey = preparePassword(DecodeBase64(passwd))
		  
		  if db.Connect = false then 
		    Return new localconfRecord("Error accessing settings file: " + db.ErrorMessage)
		  end if
		  
		  dim rs as RecordSet
		  dim WHERE as String
		  
		  
		  if recordObject.objidx > 0 then  // objidx is our match criterion -- it can only mean update
		    
		    rs = db.SQLSelect("SELECT * FROM localconf WHERE objidx = " + str(recordObject.objidx))
		    if db.Error then Return new localconfRecord("Error accessing settings file: " + db.ErrorMessage)
		    if rs.RecordCount = 0 then Return new localconfRecord("Configuration record " + str(recordObject.objidx) + " does not exist!")
		    
		    rs.Edit
		    rs.Field("value").StringValue = recordObject.value
		    rs.Field("comment").StringValue = recordObject.comment
		    rs.Update
		    if db.Error then Return new localconfRecord("Error updating settings file: " + db.ErrorMessage)
		    
		    db.Close
		    
		  else // we look for the (application/user/section/key) combination
		    
		    if recordObject.key.Trim = "" then Return new localconfRecord("Key should not be empty!")
		    
		    select case Count(recordObject)
		    case -1  // error
		      Return new localconfRecord("Error reading config file: " + LastError)
		    case 0  // insert
		      
		      dim newRecord as new DatabaseRecord
		      
		      newRecord.Column("application") = if(recordObject.application.Trim = "" , GlobalName , recordObject.application.Trim.Uppercase)
		      newRecord.Column("user") = if(recordObject.user.Trim = "" , GlobalName , recordObject.user.Trim.Uppercase)
		      newRecord.Column("section") = if(recordObject.section.Trim = "" , GlobalName , recordObject.section.Trim.Uppercase)
		      newRecord.Column("key") = recordObject.key.Trim.Uppercase
		      
		      if recordObject.value <> "" then newRecord.Column("value") = recordObject.value
		      if recordObject.comment.Trim <> "" then newRecord.Column("comment") = recordObject.comment
		      
		      db.InsertRecord("localconf" , newRecord)
		      if db.Error then Return new localconfRecord("Error writing settings file: " + db.ErrorMessage)
		      
		    case 1  // update
		      WHERE = "application = " + if(recordObject.application.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.application.Trim.Uppercase + "'") + " AND "
		      WHERE = WHERE + "user = " + if(recordObject.user.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.user.Trim.Uppercase + "'") + " AND "
		      WHERE = WHERE + "section = " + if(recordObject.section.Trim = "" , "'" + GlobalName + "'" , "'" + recordObject.section.Trim.Uppercase + "'") + " AND "
		      WHERE = WHERE + "key = '" + recordObject.key.Trim.Uppercase + "'"
		      
		      rs = db.SQLSelect("SELECT * FROM localconf WHERE " + WHERE)
		      if db.Error then Return new localconfRecord("Error accessing settings file: " + db.ErrorMessage)
		      if rs.RecordCount <> 1 then Return new localconfRecord("Internal data inconsistency!")
		      
		      rs.Edit
		      rs.Field("value").StringValue = recordObject.value
		      rs.Field("comment").StringValue = recordObject.comment
		      rs.Update
		      if db.Error then Return new localconfRecord("Error updating settings file: " + db.ErrorMessage)
		      
		      db.Close
		      
		    else // it's an array --this method can't act on arrays
		      Return new localconfRecord("Cannot use upsert on an array!")
		    end select
		    
		  end if
		  
		  Return recordObject
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function validateCachedParams() As Boolean
		  if file = nil then return false
		  if file.Exists = false then return false
		  if passwd = "" then return false
		  if DecodeBase64(passwd) = "" then return False
		  
		  Return true
		  
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
