		<% 
		mquery = "SELECT ModifyDate FROM tbl_ModifyLog WHERE FormModified=" & Int(gid) & " ORDER BY ModifyDate DESC"
		Set GetModify = Con.Execute(mquery)
		
		If((GetModify.BOF) and (GetModify.EOF)) Then %>
			Last Modified:	unavailable<br>
			Agency ID# <% Response.write Session("agencyidn") 
			
		Else
			GetModify.MoveFirst
			 %>
			Last Modified: <%= GetModify("ModifyDate") %><br><%=gid%><br>
			<% 
			GetModify.Close
			Set GetModify = Nothing
			 %>
			<!-- Agency ID# <%' Response.write Session("agencyidn")%> -->
	
		<% End If %>
