<%@LANGUAGE="VBSCRIPT"%>


<!--#include file="../Connections/NAD_BE.asp" -->
<!--#INCLUDE FILE="../security/passwordpro/check_user_inc.asp"-->

<%
templateURL="http://agencyconnection.bbbs.org/site/lookup.asp?c=9dJGKRNqFmG&b=1809973"
%>
<!--#INCLUDE FILE="../media/inc/kinterawrapper.asp"-->
<% Response.Write(LeftContent(myStr))%>


<%
set AgencySummary = Server.CreateObject("ADODB.Recordset")
AgencySummary.ActiveConnection = ConnStr
Dim SQL
SQL = "SELECT * FROM vw_2006AAI_No_Expenses_or_Board_Data ORDER BY AgencyID"


AgencySummary.Source = SQL
AgencySummary.CursorType = 0
AgencySummary.CursorLocation = 2
AgencySummary.Open()
if (AgencySummary.EOF) then
	recordCount = -1
else
	AgencySummaryRS = AgencySummary.GetRows
	recordCount = UBound(AgencySummaryRS,2)
	if isEmpty(recordCount) then
		recordCount = -1
	end if
end if
recordCount = recordCount+1
%>
		
<TABLE WIDTH="100%"  CELLPADDING="2" CELLSPACING="2" BORDER="0">
	<TR>
		<TD VALIGN="TOP"></TD>
		<TD VALIGN="TOP"></TD>
		<TD VALIGN="TOP"></TD>
		<TD VALIGN="TOP"></TD>

	</TR>
	<TR>
		<td align="center" colspan="3" BGCOLOR="#c0c0c0">NO 2006 AAI DATA ENTERED</td>
	</TR>
	<TR>
		
		<TD VALIGN="TOP" CLASS="title"><strong>Agency ID</strong></TD>
		<TD VALIGN="TOP" CLASS="title"><strong>Agency Name</strong></TD>
		<TD VALIGN="TOP" CLASS="title"><strong>Region</strong></TD>
	</TR>
	<%
	for i = 0 to recordCount-1
	%>
	<TR> 
	  <TD VALIGN="TOP"><%=AgencySummaryRS(0,i)%></TD>
	  <TD VALIGN="TOP" CLASS="results"><%=AgencySummaryRS(1,i)%></TD>
	  <TD VALIGN="TOP" CLASS="results"><%=AgencySummaryRS(2,i)%></TD>
	</TR>
	<%
	next
	%>
</TABLE	>
		

<%Response.Write(RightContent(myStr))%>