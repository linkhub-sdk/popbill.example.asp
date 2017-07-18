
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 문자전송내역 목록을 조회합니다.
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'시작일자
	SDate = "20170601"

	'종료일자
	EDate = "20170731"					
	
	'전송상태값 배열, 1-대기, 2-성공, 3-실패, 4-취소
	Dim State(4)
	State(0) = "1"
	State(1) = "2"
	State(2) = "3"
	State(3) = "4"

	'검색대상 배열, SMS., LMS, MMS
	Dim Item(3)
	Item(0) = "SMS"
	Item(1) = "LMS"
	Item(2) = "MMS"

	' 예약전송여부
	ReserveYN = False	

	' 개인조회여부 
	SenderYN = False		

	' 정렬방향, D-내림차순, A-오름차순
	Order = "D"				

	' 페이지 번호 
	Page = 1					

	' 페이지당 검색개수 
	PerPage = 30			
	
	On Error Resume Next

	Set resultObj = m_MessageService.Search(testCorpNum, SDate, EDate, Item, ReserveYN, SenderYN, Order, Page, PerPage)
	
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
				<legend>문자메시지 전송내역 조회 </legend>
				<ul>
						<li> code : <%=resultObj.code%></li>
						<li> total : <%=resultObj.total%></li>
						<li> pageNum : <%=resultObj.pageNum%></li>
						<li> perPage : <%=resultObj.perPage%></li>
						<li> pageCount : <%=resultObj.pageCount%></li>
						<li> message : <%=resultObj.message%></li>
				</ul>
					<% If code = 0 Then
						For i=0 To UBound(resultObj.list) -1
					%>

						<fieldset class="fieldset2">
							<legend> 문자메시지 전송결과 [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
							<ul>
								<li>state : <%=resultObj.list(i).state%> </li>

								<li>state (전송상태 코드) : <%=resultObj.list(i).state%> </li>
								<li>result (전송결과 코드) : <%=resultObj.list(i).result%> </li>
								<li>subject (메시지 제목) : <%=resultObj.list(i).subject%> </li>
								<li>content (메시지 내용) : <%=resultObj.list(i).content%> </li>
								<li>type (메시지 유형) : <%=resultObj.list(i).msgType%> </li>
								<li>sendnum (발신번호) : <%=resultObj.list(i).sendnum%> </li>
								<li>senderName (발신자명) : <%=resultObj.list(i).senderName%> </li>
								<li>receiveNum (수신번호) : <%=resultObj.list(i).receiveNum%> </li>
								<li>receiveName (수신자명) : <%=resultObj.list(i).receiveName%> </li>
								<li>receiptDT (접수일시) : <%=resultObj.list(i).receiptDT%> </li>
								<li>sendDT (전송일시) : <%=resultObj.list(i).sendDT%> </li>
								<li>resultDT (전송결과 수신일시) : <%=resultObj.list(i).resultDT%> </li>
								<li>reserveDT (예약일시) : <%=resultObj.list(i).reserveDT%> </li>
								<li>tranNet (전송처리 이동통신사명) : <%=resultObj.list(i).tranNet%> </li>
							</ul>
						</fieldset>

					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message : <%=message%> </li>
					<% End If %>

			</fieldset>
		 </div>
	</body>
</html>