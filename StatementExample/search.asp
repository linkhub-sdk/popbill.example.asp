<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 전자명세서 목록을 조회합니다.
	' - 응답항목에 대한 자세한 사항은 "[전자명세서 API 연동매뉴얼] >
	'   3.3.3. Search (목록 조회)" 를 참조하시기 바랍니다.
	'**************************************************************

	'팝빌회원 사업자번호
	testCorpNum = "1234567890"

	'검색일자 유형, R-등록일자, W-작성일자, I-발행일자
	DType = "W"

	'시작일자, yyyyMMdd
	SDate = "20160901"				

	'종료일자, yyyyMMdd
	EDate = "20161131"				

	' 전송상태값 배열, 미기지새 전체조회, 문서상태값 3자리 배열, 2,3번째 자리 와일드카드 사용가능
	Dim State(2)
	State(0) = "2**"
	State(1) = "3**"

	'명세서 종류코드배열 
	' - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	Dim ItemCode(6)
	ItemCode(0) = "121"
	ItemCode(1) = "122"
	ItemCode(2) = "123"
	ItemCode(3) = "124"
	ItemCode(4) = "125"
	ItemCode(5) = "126"
	
	'정렬방향, A-오름차순, D-내림차순
	Order = "D"			

	'페이지 번호
	Page = 1				

	'페이지당 검색개수
	PerPage = 20		

	'거래처정보, 거래처상호 또는 사업자등록번호 기재, 미기지시 전체조회
	SQuery = ""	

	On Error Resume Next

	Set result = m_StatementService.Search(testCorpNum, DType, SDate, EDate, State, ItemCode, Order, Page, PerPage, SQuery)

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
				<legend>전자명세서 목록조회</legend>
					<% If code = 0 Then %>
					<ul>
						<li> code : <%=result.code%></li>
						<li> total : <%=result.total%></li>
						<li> pageNum : <%=result.pageNum%></li>
						<li> perPage : <%=result.perPage%></li>
						<li> pageCount : <%=result.pageCount%></li>
						<li> message : <%=result.message%></li>
					</ul>
					
					<% For i=0 To UBound(result.list)-1 %>

						<fieldset class="fieldset2">
							<legend> 전자명세서 조회결과 [ <%=i+1%> / <%=UBound(result.list)%> ] </legend>
							<ul>
								<li>itemKey : <%=result.list(i).itemKey%> </li>
								<li>stateCode : <%=result.list(i).stateCode%> </li>
								<li>taxType : <%=result.list(i).taxType%> </li>
								<li>purposeType : <%=result.list(i).purposeType%> </li>
								<li>writeDate : <%=result.list(i).writeDate%> </li>
								<li>senderCorpName : <%=result.list(i).senderCorpName%> </li>
								<li>senderCorpNum : <%=result.list(i).senderCorpNum%> </li>
								<li>senderPrintYN : <%=result.list(i).senderPrintYN%> </li>
								<li>receiverCorpName : <%=result.list(i).receiverCorpName%> </li>
								<li>receiverCorpNum : <%=result.list(i).receiverCorpNum%> </li>
								<li>receiverPrintYN : <%=result.list(i).receiverPrintYN%> </li>
								<li>supplyCostTotal : <%=result.list(i).supplyCostTotal%> </li>
								<li>taxTotal : <%=result.list(i).taxTotal%> </li>
								<li>issueDT : <%=result.list(i).issueDT%> </li>
								<li>stateDT : <%=result.list(i).stateDT%> </li>
								<li>openYN : <%=result.list(i).openYN%> </li>
								<li>stateMemo : <%=result.list(i).stateMemo%> </li>
								<li>regDT : <%=result.list(i).regDT%> </li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
			</fieldset>
		 </div>
	</body>
</html>