<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'팝빌 회원 사업자번호, "-"제외 10자리
	userID = "testkorea"				'팝빌 회원 아이디
	itemCode = "121"					'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	mgtKey = "20150202-03"				'연동관리번호

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
				<legend>전자명세서 요약/상태 정보확인</legend>
				<ul>
					<% If code = 0 Then %>
						<li>itemKey : <%=result.itemKey%> </li>
						<li>stateCode : <%=result.stateCode%> </li>
						<li>taxType : <%=result.taxType%> </li>
						<li>purposeType : <%=result.purposeType%> </li>
						<li>writeDate : <%=result.writeDate%> </li>
						<li>senderCorpName : <%=result.senderCorpName%> </li>
						<li>senderCorpNum : <%=result.senderCorpNum%> </li>
						<li>receiverCorpName : <%=result.receiverCorpName%> </li>
						<li>receiverCorpNum : <%=result.receiverCorpNum%> </li>
						<li>supplyCostTotal : <%=result.supplyCostTotal%> </li>
						<li>taxTotal : <%=result.taxTotal%> </li>
						<li>issueDT : <%=result.issueDT%> </li>
						<li>stateDT : <%=result.stateDT%> </li>
						<li>openYN : <%=result.openYN%> </li>
						<li>stateMemo : <%=result.stateMemo%> </li>
						<li>regDT : <%=result.regDT%> </li>
					<% Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>