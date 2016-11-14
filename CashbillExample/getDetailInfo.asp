<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 현금영수증 1건의 상세정보를 조회합니다.
	' - 응답항목에 대한 자세한 사항은 "[현금영수증 API 연동매뉴얼] > 4.1.
	'   현금영수증 구성" 을 참조하시기 바랍니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'문서관리번호
	mgtKey = "20161114-01"       

	On Error Resume Next

	Set Presponse = m_CashbillService.GetDetailInfo(testCorpNum, mgtKey, userID)

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
				<legend>현금영수증 상세정보 확인</legend>
				<ul>
					<% If code = 0 Then %>
						<fieldset class="fieldset2">
							<ul>
								<li>mgtKey : <%=Presponse.mgtKey%></li>
								<li>tradeDate : <%=Presponse.tradeDate%></li>
								<li>tradeUsage : <%=Presponse.tradeUsage%></li>
								<li>tradeType : <%=Presponse.tradeType %></li>
								<li>taxationType : <%=Presponse.taxationType%></li>
								<li>supplyCost : <%=Presponse.supplyCost%></li>
								<li>tax : <%=Presponse.tax %></li>
								<li>serviceFee : <%=Presponse.serviceFee%></li>
								<li>totalAmount : <%=Presponse.totalAmount%></li>

								<li>franchiseCorpNum : <%=Presponse.franchiseCorpNum%></li>
								<li>franchiseCorpName : <%=Presponse.franchiseCorpName%></li>
								<li>franchiseCEOName : <%=Presponse.franchiseCEOName%></li>
								<li>franchiseAddr : <%=Presponse.franchiseAddr%></li>
								<li>franchiseTEL : <%=Presponse.franchiseTEL %></li>

								<li>identityNum : <%=Presponse.identityNum%></li>
								<li>customerName : <%=Presponse.customerName%></li>
								<li>itemName : <%=Presponse.itemName%></li>
								<li>orderNumber : <%=Presponse.orderNumber%></li>
								
								<li>email : <%=Presponse.email%></li>
								<li>hp : <%=Presponse.hp%></li>
								<li>fax : <%=Presponse.fax%></li>
								<li>smssendYN : <%=Presponse.smssendYN%></li>
								<li>faxsendYN : <%=Presponse.faxsendYN %></li>
								
								<li>confirmNum : <%=Presponse.confirmNum%></li>
								
								<li>orgConfirmNum : <%=Presponse.orgConfirmNum%></li>
								<li>orgTradeDate : <%=Presponse.orgTradeDate%></li>
							</ul>
						</fieldset>
					<%	Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
					
				</ul>
			</fieldset>
		 </div>
	</body>
</html>