<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 1건의 전자명세서 상태/요약 정보를 확인합니다.
	' - https://docs.popbill.com/statement/asp/api#GetInfo
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"	
	
	'팝빌 회원 아이디
	userID = "testkorea"				

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"					

	'문서번호
	mgtKey = "20211201-001"				

	On Error Resume Next

	Set result = m_StatementService.GetInfo(testCorpNum, itemCode, mgtKey, userID)

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
				<legend>전자명세서 상태/요약 정보확인</legend>
				<ul>
					<% If code = 0 Then %>
						<li> itemKey(아이템키) : <%=result.itemKey %></li>
						<li> itemCode(문서종류코드) : <%=result.itemCode %></li>
						<li> stateCode(상태코드) : <%=result.stateCode %></li>
						<li> taxType(세금형태) : <%=result.taxType %></li>
						<li> purposeType(영수/청구) : <%=result.purposeType %></li>
						<li> writeDate(작성일자) : <%=result.writeDate %></li>
						<li> senderCorpName(발신자 상호) : <%=result.senderCorpName %></li>
						<li> senderCorpNum(발신자 사업자번호) : <%=result.senderCorpNum %></li>
						<li> senderPrintYN(발신자 인쇄여부) : <%=result.senderPrintYN %></li>
						<li> receiverCorpName(수신자 상호) : <%=result.receiverCorpName %></li>
						<li> receiverCorpNum(수신자 사업자번호) : <%=result.receiverCorpNum %></li>
						<li> receiverPrintYN(수신자 인쇄여부) : <%=result.receiverPrintYN %></li>
						<li> supplyCostTotal(공급가액 합계) : <%=result.supplyCostTotal %></li>
						<li> taxTotal(세액 합계) : <%=result.taxTotal %></li>
						<li> issueDT(발행일시) : <%=result.issueDT %></li>
						<li> stateDT(상태 변경일시) : <%=result.stateDT %></li>
						<li> openYN(메일 개봉 여부) : <%=result.openYN %></li>
						<li> openDT(개봉 일시) : <%=result.openDT %></li>
						<li> stateMemo(상태메모) : <%=result.stateMemo %></li>
						<li> regDT(등록일시) : <%=result.regDT %></li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>