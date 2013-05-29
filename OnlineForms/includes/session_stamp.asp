<% Response.Expires = 0 %>

<% 
If Len(Session("login")) = 0 Then
	Response.Redirect("login.asp?go=expired")
End IF
 %>