<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
<link rel="stylesheet" type="text/css" href="../../bbbsa.css">
</head>

<body>

<%= now %>

<P>
<%= Session.SessionID %>
<P>

<% form = "Expenses" %> 
<% q = "SELECT UserName,ModifyDate FROM tbl_ModifyLog WHERE AgencyID=" & Session("agencyidn") & " AND FormModified=" & "Get" & form & "(" & chr(34) & form & "ID" & chr(34) & ")" & " ORDER BY ModifyDate DESC" %>
<%= q %>
</body>
</html>
