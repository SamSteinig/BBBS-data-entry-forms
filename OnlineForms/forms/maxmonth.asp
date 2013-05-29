
<% 


Set Con = Server.CreateObject("ADODB.Connection")
	Con.Open "BBBSAforms", "sa","12sist12"

monthquery = "SELECT Max(Month) as MaxMonth from tbl_frmPerformance where agencyid = '9999'"

' monthquery = "select month from tbl_frmPerformance where agencyid = '9999'"

' maxmonth = 6
set getmonth = con.execute(monthquery)



response.write getmonth("MaxMonth")

' while not getmonth.eof
'	response.write getmonth("Month") & "<br>"
'	getmonth.MoveNext
' Wend
	
getmonth.close
set getmonth = nothing




%>