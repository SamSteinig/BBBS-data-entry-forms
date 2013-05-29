<% 
y = Request("y")
 %>
<input type="hidden" name="year" value="<%= y %>">
<% 
If Request("m") > 0 Then
	m = Request("m")
 %>
<input type="hidden" name="month" value="<%= m %>">
<% 
End If
 %>
<input type="hidden" name="User" value="<%= Session("login") %>">
<input type="hidden" name="AgencyIDN" value="<%= Session("agencyidn") %>">

