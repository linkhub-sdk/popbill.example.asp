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
	mgtKey = "20170718-04"       

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
								<li>confirmNum (국세청승인번호) : <%=Presponse.confirmNum%></li>
								<li>mgtKey (문서관리번호) : <%=Presponse.mgtKey%></li>
								<li>tradeDate (거래일자) : <%=Presponse.tradeDate%></li>
								<li>tradeUsage (거래유형) : <%=Presponse.tradeUsage%></li>
								<li>tradeType (현금영수증 형태) : <%=Presponse.tradeType %></li>
								<li>taxationType (과세형태) : <%=Presponse.taxationType%></li>
								<li>supplyCost (공급가액) : <%=Presponse.supplyCost%></li>
								<li>tax (세액) : <%=Presponse.tax %></li>
								<li>serviceFee (봉사료) : <%=Presponse.serviceFee%></li>
								<li>totalAmount (거래금액) : <%=Presponse.totalAmount%></li>

								<li>franchiseCorpNum (발행자 사업자번호) : <%=Presponse.franchiseCorpNum%></li>
								<li>franchiseCorpName (발행자 상호) : <%=Presponse.franchiseCorpName%></li>
								<li>franchiseCEOName (발행자 대표자명) : <%=Presponse.franchiseCEOName%></li>
								<li>franchiseAddr (가맹점 주소) : <%=Presponse.franchiseAddr%></li>
								<li>franchiseTEL (가맹점 전화번호) : <%=Presponse.franchiseTEL %></li>

								<li>identityNum (거래처 식별번호) : <%=Presponse.identityNum%></li>
								<li>customerName (고개명) : <%=Presponse.customerName%></li>
								<li>itemName (상품명) : <%=Presponse.itemName%></li>
								<li>orderNumber (가맹점 주문번호) : <%=Presponse.orderNumber%></li>
								
								<li>email (고객 이메일) : <%=Presponse.email%></li>
								<li>hp (고객 휴대폰번호) : <%=Presponse.hp%></li>
								<li>smssendYN (알림문자 전송여부) : <%=Presponse.smssendYN%></li>
							
								<li>orgConfirmNum (원본 현금영수증 국세청승인번호) : <%=Presponse.orgConfirmNum%></li>
								<li>orgTradeDate (원본 현금영수증 거래일자) : <%=Presponse.orgTradeDate%></li>
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