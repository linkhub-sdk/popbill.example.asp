
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
	' - 최대 검색기간 : 6개월 이내
	' - https://docs.popbill.com/message/asp/api#Search
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'시작일자
	SDate = "20210501"

	'종료일자
	EDate = "20210601"					
	
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
	
	'조회 검색어.
	'문자 전송시 입력한 발신자명 또는 수신자명 기재.
	'조회 검색어를 포함한 발신자명 또는 수신자명을 검색합니다.
	QString = ""
	
	' On Error Resume Next

	Set resultObj = m_MessageService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)
	
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
				<legend>문자메세지 전송내역 조회 </legend>
				<ul>
						<li> code (응답코드) : <%=resultObj.code%></li>
						<li> total (총 검색결과 건수) : <%=resultObj.total%></li>
						<li> pageNum (페이지 번호) : <%=resultObj.pageNum%></li>
						<li> perPage (페이지당 목록개수) : <%=resultObj.perPage%></li>
						<li> pageCount (페이지 개수) : <%=resultObj.pageCount%></li>
						<li> message (응답메시지) : <%=resultObj.message%></li>
				</ul>
					<% If code = 0 Then
						For i=0 To UBound(resultObj.list) -1
					%>

						<fieldset class="fieldset2">
							<legend> 문자메시지 전송결과 [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
							<ul>

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
								<li>receiptNum (접수번호) : <%=resultObj.list(i).receiptNum%> </li>
								<li>requestNum (요청번호) : <%=resultObj.list(i).requestNum%> </li>
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