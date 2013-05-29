<%
Dim HelpID
if (Request.QueryString("HelpID") <> "") then HelpID = Request.QueryString("HelpID")

Dim SixMonthsAgo
if (Request.QueryString("SixMonthsAgo")<>"") then SixMonthsAgo = Request.QueryString("SixMonthsAgo")

Dim Now
if (Request.QueryString("Now")<>"") then Now = Request.QueryString("Now")

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<html>
<head>
	<title>Fee Calculation Form Help</title>
</head>
<link rel="STYLESHEET" type="text/css" href="../includes/bbbsa_forms.css">

<body>

<% if HelpID = "total_expenditures" then %>	
	<p>
	<span class = "formIndex">(A) TOTAL EXPENDITURES</span>
	</p>
	
	<span class = "formMain">
	<p>This amount must agree with your last completed financial audit.  If your audit is not yet completed, use your unaudited figures.  Adjustments, if material, will be made when the audited figures are sent to us.  Send your most current audited financial statement and note on the front of the form when your fiscal year ends.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_fees_paid_to_bbbsa" then %>	
	<p>
	<span class = "formIndex">(B) Less: Prior Year Fees Paid to BBBSA</span>
	</p>
	
	<span class = "formMain">
	<p>Includes affiliation fees only.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_capital_purchases" then %>	
	<p>
	<span class = "formIndex">(C) Less: Prior Year Capital Purchases</span>
	</p>
	
	<span class = "formMain">
	<p>Land, buildings, and equipment purchased <strong>(only if included in your Total Expenditure amount)</strong>.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_depreciation" then %>	
	<p>
	<span class = "formIndex">(D) Less: Prior Year Depreciation</span>
	</p>
	
	<span class = "formMain">
	<p>Do not include depreciation expenses for the capital purchases included above.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

<% if HelpID = "prior_year_fundraising_expenses" then %>	
	<p>
	<span class = "formIndex">(E) Less: Prior Year Fundraising Expenses</span>
	</p>
	
	<span class = "formMain">
	<p>Include in this line all expenses related to fundraising activities.  This can include direct salaries and fringes paid to fundraising personnel.  Since this may have a significant impact on your fees, these expenses must be direct fundraising expenses only, not allocated.
	</p>
	<p>All deductions should be substantiated and will be verified against a copy of your last audit, which must be submitted to BBBSA after the audit has been finalized.
	</p>
	
	</span>
<% end if %>

</body>
</html>
