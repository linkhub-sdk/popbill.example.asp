<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		  '팝빌 회원 사업자번호, "-" 제외
	userID = "testkorea"			  '팝빌 회원 아이디
	receiptNum = "015012713201000001" '팩스 전송시 발급받은 전송번호
 
 	'전송결과코드는 [팝빌 FAX API 연동매뉴얼 5.부록] 참조
	
	On Error Resume Next

	Set result = m_FaxService.GetFaxDetail(testCorpNum, receiptNum, userID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>팩스전송 전송결과 확인 </legend>
				<% If code = 0 Then %>
					<ul>
						<li>sendState(전송상태) : <%=result.sendState%> </li>
						<li>convState(변환상태) : <%=result.convState%> </li>
						<li>sendNum(발신번호) : <%=result.sendNum%> </li>
						<li>receiveNum(수신번호) : <%=result.receiveNum%> </li>
						<li>receiveName(수신자명) : <%=result.receiveName%> </li>
						<li>sendPageCnt(페이지수) : <%=result.sendPageCnt%></li>
						<li>successPageCnt(성공 페이지수) : <%=result.successPageCnt%></li>
						<li>failPageCnt(실패 페이지수) : <%=result.failPageCnt%></li>
						<li>refundPageCnt(환불 페이지수) : <%=result.refundPageCnt%></li>
						<li>cancelPageCnt(취소 페이지수) : <%=result.cancelPageCnt%></li>
						<li>reserveDT(예약시간) : <%=result.reserveDT%></li>
						<li>sendDT(발송시간) : <%=result.sendDT%></li>
						<li>sendResult(통신사 처리결과) : <%=result.sendResult%></li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>