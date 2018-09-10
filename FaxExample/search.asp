
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 팩스전송 내역을 조회합니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'시작일자, yyyyMMdd
	SDate = "20180601"					

	'종료일자, yyyyMMdd
	EDate = "20180730"					
	
	' 전송상태값 배열, 1-대기, 2-성공, 3-실패, 4-취소
	Dim State(4)
	State(0) = "1"
	State(1) = "2"
	State(2) = "3"
	State(3) = "4"
	
	'예약전송 검색여부
	ReserveYN = False
	
	'개인조회 여부
	SenderOnlyYN = False

	'정렬발향, A-오름차순, D-내림차순
	Order = "D"

	'페이지 번호
	Page = 1

	'페이지당 검색개수
	PerPage = 20
	
	On Error Resume Next

	Set result = m_FaxService.Search(testCorpNum, SDate, EDate, State, ReserveYN, SenderOnlyYN, Order, Page, PerPage)
	
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
				<legend>팩스전송 전송내역 조회 </legend>
					<ul>
						<li> code : <%=result.code%></li>
						<li> total : <%=result.total%></li>
						<li> pageNum : <%=result.pageNum%></li>
						<li> perPage : <%=result.perPage%></li>
						<li> pageCount : <%=result.pageCount%></li>
						<li> message : <%=result.message%></li>
					</ul>
				<% If code = 0 Then 
						For i=0 To UBound(result.list)-1
				%>
					<fieldset class="fieldset2">
							<legend> 팩스 전송결과 [ <%=i+1%> /  <%=UBound(result.list)%> ] </legend>
							<ul>
								<li>state (전송상태 코드) : <%=result.list(i).state%> </li>
								<li>result (전송결과 코드) : <%=result.list(i).result%> </li>
								<li>sendNum (발신번호) : <%=result.list(i).sendNum%> </li>
								<li>senderName (발신자명) : <%=result.list(i).senderName%> </li>
								<li>receiveNum (수신번호) : <%=result.list(i).receiveNum%> </li>
								<li>receiveName (수신자명) : <%=result.list(i).receiveName%> </li>
								<li>title (팩스 제목) : <%=result.list(i).title %> </li>
								<li>sendPageCnt (페이지수) : <%=result.list(i).sendPageCnt%></li>
								<li>successPageCnt (성공 페이지수) : <%=result.list(i).successPageCnt%></li>
								<li>failPageCnt (실패 페이지수) : <%=result.list(i).failPageCnt%></li>
								<li>refundPageCnt (환불 페이지수) : <%=result.list(i).refundPageCnt%></li>
								<li>cancelPageCnt (취소 페이지수) : <%=result.list(i).cancelPageCnt%></li>
								<li>reserveDT (예약시간) : <%=result.list(i).reserveDT%></li>
								<li>sendDT (발송시간) : <%=result.list(i).sendDT%></li>
								<li>receiptDT (전송 접수시간) : <%=result.list(i).receiptDT%></li>
								<li>fileNames (전송파일명 배열) : <%=result.list(i).fileNames%></li>
							</ul>
						</fieldset>
				<%	
					Next
					Else  
				%>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>

			</fieldset>
		 </div>
	</body>
</html>