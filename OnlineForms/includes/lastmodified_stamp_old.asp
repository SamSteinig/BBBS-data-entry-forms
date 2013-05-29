		<% 
		query2 = "SELECT UserName FROM tbl_ModifyLog WHERE AgencyID='" & Session("agencyidn") & "' AND (ModifyType='new') AND FormModified=" & gid
		Set GetModify = Con.Execute(query2)
		 %>
		Created By: <%= GetModify("UserName") %><br>
		<% 
		GetModify.Close
		Set GetModify = Nothing
		query3 = "SELECT UserName,ModifyDate FROM tbl_ModifyLog WHERE AgencyID='" & Session("agencyidn") & "' AND FormModified=" & gid & " ORDER BY ModifyDate DESC"
		Set GetModify = Con.Execute(query3)
		GetModify.MoveFirst
		 %>
		Last Modified: <%= GetModify("ModifyDate") %><br>
		Last Modified By: <%= GetModify("UserName") %></td>
		<% 
		GetModify.Close
		Set GetModify = Nothing
		 %>