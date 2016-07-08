<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'팝빌회원 사업자번호, "-" 제외
	itemCode = "121"						'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	UserID = "testkorea"					'팝빌회원 아이디
		
	Set result = m_StatementService.GetChargeInfo ( testCorpNum, ItemCode, UserID )

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>과금정보 확인</legend>
				<%
					If code = 0 Then
				%>
						<ul>
							<li> unitCost (단가) : <%=result.unitCost%></li>
							<li> chargeMethod (과금유형) : <%=result.chargeMethod%></li>
							<li> rateSystem (과금제도) : <%=result.rateSystem%></li>
						</ul>
				<%
					Else
				%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>
