<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 수집 작업이 완료된 계좌의 거래내역을 조회합니다.
	' - https://docs.popbill.com/easyfinbank/asp/api#Search
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	

	'팝빌회원 아이디
	UserID = "testkorea"
	
	'수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
	JobID = "019123114000000010"

	'거래유형 배열, I-입금, O-출금
	Dim TradeType(2) 
	TradeType(0) = "I"
	TradeType(1) = "O"

	'조회 검색어, 입금/출금액, 메모, 적요 like 검색
	SearchString = ""

	'페이지 번호 
	Page  = 1

	'페이지당 목록개수
	PerPage = 10

	'정렬방항, D-내림차순, A-오름차순
	Order = "D"

	On Error Resume Next

	Set result = m_EasyFinBankService.Search(testCorpNum, JobID, TradeType, SearchString, _	
								Page, PerPage, Order, UserID)

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
				<legend>수집 결과 조회</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> code (응답코드) : <%=result.code%> </li>
						<li> message  (응답메시지) : <%=result.message%> </li>
						<li> total (총 검색결과 건수) : <%=result.total%> </li>
						<li> perPage (페이지당 검색개수) : <%=result.perPage%> </li>
						<li> pageNum (페이지 번호) : <%=result.pageNum%> </li>
						<li> pageCount (페이지 개수) : <%=result.pageCount%> </li>
					</ul>

				<%
					For i=0 To UBound(result.list) -1 
				%>
					<fieldset class="fieldset2">					
						<legend>거래내역 정보 [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
							<ul>								
								<li> tid (거래내역 아이디) : <%= result.list(i).tid %></li>
								<li> trdate (거래일자) : <%= result.list(i).trdate %></li>
								<li> trserial (거래일자별 거래내역 순번) : <%= result.list(i).trserial %></li>
								<li> trdt (거래일시) : <%= result.list(i).trdt %></li>
								<li> accIn (입금액) : <%= result.list(i).accIn %></li>
								<li> accOut (출금액) : <%= result.list(i).accOut %></li>
								<li> balance (잔액) : <%= result.list(i).balance %></li>
								<li> remark1 (비고1) : <%= result.list(i).remark1 %></li>
								<li> remark2 (비고2) : <%= result.list(i).remark2 %></li>
								<li> remark3 (비고3) : <%= result.list(i).remark3 %></li>
								<li> memo (메모) : <%= result.list(i).memo %></li>
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
				<%	
					End If
				%>
			</fieldset>
		 </div>
	</body>
</html>

