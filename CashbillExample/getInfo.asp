<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	 '팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"		 '팝빌 회원 아이디
	mgtKey = "20150201-01"		 '연동관리번호

	On Error Resume Next

	Set Presponse = m_CashbillService.GetInfo(testCorpNum, mgtKey, userID)

	If Err.Number <> 0 then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팝빌 현금영수증 상태/요약 정보확인 </legend>
				<ul>
					<% If code = 0 Then %>
						<li>itemKey : <%=Presponse.itemKey%></li>
						<li>mgtKey : <%=Presponse.mgtKey%></li>
						<li>tradeDate : <%=Presponse.tradeDate%></li>
						<li>issueDT : <%=Presponse.issueDT%></li>
						<li>customerName : <%=Presponse.customerName%></li>
						<li>itemName : <%=Presponse.itemName%></li>
						<li>identityNum : <%=Presponse.identityNum%></li>
						<li>taxactionType : <%=Presponse.taxationType%></li>
						<li>totalAmount : <%=Presponse.totalAmount%></li>
						<li>tradeUsage : <%=Presponse.tradeUsage%></li>
						<li>tradeType : <%=Presponse.tradeType%></li>
						<li>stateCode : <%=Presponse.stateCode%></li>
						<li>stateDT : <%=Presponse.stateDT%></li>
						<li>printYN : <%=Presponse.printYN%></li>
						<li>confirmNum : <%=Presponse.confirmNum%></li>
						<li>orgTradeDate : <%=Presponse.orgTradeDate%></li>
						<li>orgConfirmNum : <%=Presponse.orgConfirmNum%></li>
						<li>ntssendDT : <%=Presponse.ntssendDT%></li>
						<li>ntsPresponse : <%=Presponse.ntsResult%></li>
						<li>ntsPresponseDT : <%=Presponse.ntsResultDT%></li>
						<li>ntsPresponseCode : <%=Presponse.ntsResultCode%></li>
						<li>ntsPresponseMessage : <%=Presponse.ntsResultMessage%></li>
						<li>regDT : <%=Presponse.regDT%></li>
					<% Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
				</ul>
			</fieldset>
		 </div>
	</body>
</html>