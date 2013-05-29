<% 
m = 12
y = 2000
form = "Performance"
modtype= "new"

Server.ScriptTimeout = 100000

Set Con = Server.CreateObject("ADODB.Connection")
Con.Open "BBBSAforms", "sa","12sist12"
x = 0
q = "SELECT * FROM tbl_frmPerformance WHERE month=" & m & " AND year=" & y
Set GetAll = Con.Execute(q)
GetAll.MoveFirst
Do Until GetAll.EOF
	If IsNull(GetAll("ModifyLogID")) Then
		Set ModRST = Server.CreateObject("ADODB.Recordset")
		ModRST.Open "SELECT * FROM tbl_ModifyLog", Con, 1, 3
		ModRST.AddNew
		ModRST("Form") = form
		ModRST("FormModified") = GetAll("PerformanceID")
		ModRST("Year") = y
		ModRST("Month") = m
		ModRST("ModifyType") = modtype
		ModRST("UserName") = "bbbsa"
		ModRST("AgencyID") = GetAll("AgencyID")
		ModRST("ModifyDate") = Now
		ModRST.Update
		Set ModRST = Nothing
		
		q2 = "SELECT ModifyID FROM tbl_ModifyLog WHERE formModified=" & GetAll("PerformanceID")
		Set GetMod2 = Con.Execute(q2)
		iMod = GetMod2("ModifyID")
		GetMod2.Close
		Set GetMod2 = Nothing
		
		'response.write "imod=" & imod & "<br>"
		
		Set ModRST = Server.CreateObject("ADODB.Recordset")
		q3 = "SELECT * FROM tbl_frmPerformance WHERE PerformanceID=" & GetAll("PerformanceID")
		ModRST.Open q3, Con, 1, 3
		ModRST("ModifyLogID") = iMod
		ModRST.Update
		Set ModRST = Nothing
		
		x = x + 1
	End If
	GetAll.MoveNext
Loop
GetAll.Close
Set GetAll = Nothing
Con.Close
Set Con = Nothing
 %>
<P>
<b><%= x %></b> records fixed/modified