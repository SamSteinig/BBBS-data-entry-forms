<% 
'query the DB for the ID of the record it just wrote

If modtype = "new" Then
	modquery = "SELECT @@IDENTITY AS " & form & "ID FROM tbl_frm" & form
	Set GetMod = Con.Execute(modquery)
	gMod = form & "ID"
	vMod = GetMod(gMod)
	GetMod.Close
	Set GetMod = Nothing
Else
	vMod = jMod
End If

'response.write "vmod=" & vmod & "<P>"

'write that ID to the modify table
Set ModRST = Server.CreateObject("ADODB.Recordset")
ModRST.Open "SELECT * FROM tbl_ModifyLog", Con, 1, 3
ModRST.AddNew
ModRST("Form") = form
ModRST("FormModified") = vMod
ModRST("Year") = Request("year")
If form = "Performance" Then
	ModRST("Month") = m
End If
ModRST("ModifyType") = modtype
ModRST("UserName") = Session("Login")
ModRST("AgencyID") = Session("AgencyIDN")
ModRST("ModifyDate") = Now
ModRST.Update
Set ModRST = Nothing

If modtype <> "delete" Then
	'get the ID of the modify record it just wrote
	modquery2 = "SELECT @@IDENTITY AS ModifyID FROM tbl_ModifyLog"
	Set GetMod2 = Con.Execute(modquery2)
	iMod = GetMod2("ModifyID")
	GetMod2.Close
	Set GetMod2 = Nothing
	
	'response.write "imod=" & imod & "<P>"

	'write that modify id back to the first record
	Set ModRST = Server.CreateObject("ADODB.Recordset")
	qqmod = "SELECT * FROM tbl_frm" & form & " WHERE " & form & "ID=" & vMod
	ModRST.Open qqmod, Con, 1, 3
	ModRST("ModifyLogID") = iMod
	ModRST.Update
	Set ModRST = Nothing
End If
 %>