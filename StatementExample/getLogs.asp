<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 전자명세서 상태변경 이력을 확인합니다.
	' - 상태 변경이력 확인(GetLogs API) 응답항목에 대한 자세한 정보는
	'   "[전자명세서 API 연동매뉴얼] > 3.3.4 GetLogs (상태 변경이력 확인)"
	'   을 참조하시기 바랍니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-"제외 10자리
	testCorpNum = "1234567890"			

	'팝빌 회원 아이디
	userID = "testkorea"				

	'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	itemCode = "121"					

	'문서관리번호
	mgtKey = "20161114-06"				

	On Error Resume Next

	Set result = m_StatementService.GetLogs(testCorpNum, itemCode, mgtKey, userID)

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
				<legend>전자명세서 상태변경 이력 </legend>
				<ul>
					<% If code = 0 Then 
						For i=0 To result.Count-1%>
						<fieldset class="fieldset2">
						<legend> 전자명세서 상태변경 이력 [<%=i+1%>]</legend>
							<ul>
								<li>docLogType : <%=result.Item(i).docLogType%> </li>
								<li>log : <%=result.Item(i).log%> </li>
								<li>procType : <%=result.Item(i).procType%> </li>
								<li>procCorpName : <%=result.Item(i).procCorpName%> </li>
								<li>procMemo : <%=result.Item(i).procMemo%> </li>
								<li>regDT : <%=result.Item(i).regDT%> </li>
								<li>ip : <%=result.Item(i).ip%> </li>
							</ul>
						</fieldset>
					<%
						Next
						Else %>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>