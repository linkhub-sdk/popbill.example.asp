
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 카카오톡 전송내역 목록을 조회합니다. (조회기간 단위 : 최대 2개월)
	' - 카카오톡 접수일시로부터 6개월 이내 접수건만 조회할 수 있습니다.
	' - https://docs.popbill.com/kakao/asp/api#Search
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"		

	'시작일자
	SDate = "20210601"

	'종료일자
	EDate = "20210624"					
	
	'전송상태값 배열, 0-전송대기, 1-전송중, 2-전송성공, 3-대체문자 전송, 4-전송실패, 5-전송취소
	Dim State(6)
	State(0) = "0"
	State(1) = "1"
	State(2) = "2"
	State(3) = "3"
	State(4) = "4"
	State(5) = "5"

	'검색대상 배열, ATS-알림톡, FTS-친구톡 텍스트, FMS-친구톡 이미지
	Dim Item(3)
	Item(0) = "ATS"
	Item(1) = "FTS"
	Item(2) = "FMS"

	' 예약전송여부, 공백-전체조회, 1-예약전송건 조회, 0-즉시전송건 조회
	ReserveYN = ""	

	' 개인조회여부 (True-개인조회 / False-전체조회)
	SenderYN = False		

	' 정렬방향, D-내림차순, A-오름차순
	Order = "D"				

	' 페이지 번호 
	Page = 1					

	PerPage = 30			

	'조회 검색어.
	'카카오톡 전송시 입력한 발신자명 또는 수신자명 기재.
	'조회 검색어를 포함한 발신자명 또는 수신자명을 검색합니다.
	QString = ""

	On Error Resume Next

	Set resultObj = m_KakaoService.Search(testCorpNum, SDate, EDate, State, Item, ReserveYN, SenderYN, Order, Page, PerPage, QString)

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
				
					<legend>카카오톡 전송내역 조회 </legend>
					<ul>
					<% If code = 0 Then %>
							<li> code (응답코드) : <%=resultObj.code%></li>
							<li> message (응답메시지) : <%=resultObj.message%></li>
							<li> total (총 검색결과 건수) : <%=resultObj.total%></li>
							<li> pageNum (페이지 번호) : <%=resultObj.pageNum%></li>
							<li> pageCount (페이지 개수) : <%=resultObj.pageCount%></li>
							<li> perPage (페이지당 검색개수) : <%=resultObj.perPage%></li>
					</ul>
						<% 
							For i=0 To UBound(resultObj.list) -1
						%>
							<fieldset class="fieldset2">
								<legend> 카카오톡 전송결과 [ <%=i+1%> / <%= UBound(resultObj.list)%> ] </legend>
								<ul>
									<li>state (전송상태 코드) : <%=resultObj.list(i).state%> </li>
									<li>sendDT (전송일시) : <%=resultObj.list(i).sendDT%> </li>
									<li>result (전송결과 코드) : <%=resultObj.list(i).result%> </li>
									<li>resultDT (전송결과 수신일시) : <%=resultObj.list(i).resultDT%> </li>
									<li>contentType (카카오톡 유형) : <%=resultObj.list(i).contentType%> </li>
									<li>receiveNum (수신번호) : <%=resultObj.list(i).receiveNum%> </li>
									<li>receiveName (수신자명) : <%=resultObj.list(i).receiveName%> </li>
									<li>content (알림톡/친구톡 내용) : <%=resultObj.list(i).content%> </li>
									<li>altContentType (대체문자 전송타입) : <%=resultObj.list(i).altContentType%> </li>
									<li>altSendDT (대체문자 전송일시) : <%=resultObj.list(i).altSendDT%> </li>
									<li>altResult (대체문자 전송결과 코드) : <%=resultObj.list(i).altResult%> </li>
									<li>altResultDT (대체문자 전송결과 수신일시) : <%=resultObj.list(i).altResultDT%> </li>
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