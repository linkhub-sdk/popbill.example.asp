<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "4108600477"		'팝빌회원 사업자번호, "-" 제외
	sendType = "MMS"					 '전송유형 (SMS - 단문, LMS - 장문, MMS-포토)
	UserID = "innoposttest"					'팝빌회원 아이디
	
	On Error Resume Next

	Set result = m_MessageService.GetChargeInfo ( testCorpNum, sendType, UserID )

	If Err.Number <> 0 Then
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
				<legend>과금정보 조회</legend>
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
