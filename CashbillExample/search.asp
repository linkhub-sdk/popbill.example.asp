<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'팝빌 회원 사업자번호, "-" 제외
	DType = "R"							'검색일자 유형, R-등록일자, T-거래일자, I-발행일자
	SDate = "20150101"				'시작일자, yyyyMMdd
	EDate = "20160127"				'종료일자, yyyyMMdd

	' 전송상태값 배열, 미기지새 전체조회, 문서상태값 3자리 배열, 2,3번째 자리 와일드카드 사용가능
	Dim State(3)
	State(0) = "100"
	State(1) = "2**"
	State(2) = "3**"
	
	Dim TradeType(2)
	TradeType(0) = "N"
	TradeType(1) = "C"

	Dim TradeUsage(2)
	TradeUsage(0) = "P"
	TradeUsage(1) = "C"

	Dim TaxationType(2)
	TaxationType(0) = "T"
	TaxationType(1) = "N"

	Order = "D"			'정렬방향, A-오름차순, D-내림차순
	Page = 1
	PerPage = 20

	On Error Resume Next
	
	Set SearchResult = m_CashbillService.Search(testCorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TaxationType, Order, Page, PerPage)

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
				<legend>현금영수증 목록조회</legend>
					<ul>
						<li> code : <%=SearchResult.code%></li>
						<li> total : <%=SearchResult.total%></li>
						<li> pageNum : <%=SearchResult.pageNum%></li>
						<li> perPage : <%=SearchResult.perPage%></li>
						<li> pageCount : <%=SearchResult.pageCount%></li>
						<li> message : <%=SearchResult.message%></li>
					</ul>
					<% If code = 0 Then 
						For i=0 To UBound(SearchResult.list)-1 %>
						<fieldset class="fieldset2">
							<legend> 현금영수증 조회 결과 [<%=i+1%> <%=UBound(SearchResult.list)%>]</legend>
							<ul>
								<li>itemKey : <%=SearchResult.list(i).itemKey%></li>
								<li>mgtKey : <%=SearchResult.list(i).mgtKey%></li>
								<li>tradeDate : <%=SearchResult.list(i).tradeDate%></li>
								<li>issueDT : <%=SearchResult.list(i).issueDT%></li>
								<li>customerName : <%=SearchResult.list(i).customerName%></li>
								<li>itemName : <%=SearchResult.list(i).itemName%></li>
								<li>identityNum : <%=SearchResult.list(i).identityNum%></li>
								<li>taxactionType : <%=SearchResult.list(i).taxationType%></li>
								<li>totalAmount : <%=SearchResult.list(i).totalAmount%></li>
								<li>tradeUsage : <%=SearchResult.list(i).tradeUsage%></li>
								<li>tradeType : <%=SearchResult.list(i).tradeType%></li>
								<li>stateCode : <%=SearchResult.list(i).stateCode%></li>
								<li>stateDT : <%=SearchResult.list(i).stateDT%></li>
								<li>printYN : <%=SearchResult.list(i).printYN%></li>
								<li>confirmNum : <%=SearchResult.list(i).confirmNum%></li>
								<li>orgTradeDate : <%=SearchResult.list(i).orgTradeDate%></li>
								<li>orgConfirmNum : <%=SearchResult.list(i).orgConfirmNum%></li>
								<li>ntssendDT : <%=SearchResult.list(i).ntssendDT%></li>
								<li>ntsPresponse : <%=SearchResult.list(i).ntsResult%></li>
								<li>ntsPresponseDT : <%=SearchResult.list(i).ntsResultDT%></li>
								<li>ntsPresponseCode : <%=SearchResult.list(i).ntsResultCode%></li>
								<li>ntsPresponseMessage : <%=SearchResult.list(i).ntsResultMessage%></li>
								<li>regDT : <%=SearchResult.list(i).regDT%></li>
							</ul>
						</fieldset>
					<%	Next
						Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
				</fieldset>
		 </div>
	</body>
</html>