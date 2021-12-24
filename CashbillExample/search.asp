<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 현금영수증 목록을 조회합니다. (조회기간 단위 : 최대 6개월)
	' - https://docs.popbill.com/cashbill/asp/api#Search
	'**************************************************************

	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	


	'검색일자 유형, R-등록일자, T-거래일자, I-발행일자
	DType = "T"						
	
	'시작일자, yyyyMMdd
	SDate = "20211201"				

	'종료일자, yyyyMMdd
	EDate = "20211230"				

	' 전송상태값 배열, 미기지새 전체조회, 문서상태값 3자리 배열, 2,3번째 자리 와일드카드 사용가능
	Dim State(3)
	State(0) = "2**"
	State(1) = "3**"
	State(2) = "4**"
	
	'문서형태, N-일반현금영수증, C-취소현금영수증
	Dim TradeType(2)			
	TradeType(0) = "N"
	TradeType(1) = "C"

	'거래구분, P-소득공제용, C-지출증빙용
	Dim TradeUsage(2)		
	TradeUsage(0) = "P"
	TradeUsage(1) = "C"

	'거래유형, N-일반, B-도서공연, T-대중교통
	Dim TradeOpt(3)		
	TradeOpt(0) = "N"
	TradeOpt(1) = "B"
	TradeOpt(2) = "T"

	'과세형태 배열, T-과세,  N-비과세
	Dim TaxationType(2)		
	TaxationType(0) = "T"
	TaxationType(1) = "N"


	'정렬방향, A-오름차순, D-내림차순
	Order = "D"			

	'페이지번호
	Page = 1				

	'페이지당 검색개수, 최대 1000
	PerPage = 20		

	'식별번호 기재, 공백처리시 전체조회
	QString = ""		

	On Error Resume Next
	
	Set SearchResult = m_CashbillService.Search(testCorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TradeOpt, TaxationType, Order, Page, PerPage, QString)

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
						<li> code (응답 코드) : <%=SearchResult.code%></li>
						<li> message (응답 메시지) : <%=SearchResult.message%></li>
						<li> total (총 검색결과 건수) : <%=SearchResult.total%></li>
						<li> pageNum (페이지 번호) : <%=SearchResult.pageNum%></li>
						<li> perPage (페이지당 검색개수) : <%=SearchResult.perPage%></li>
						<li> pageCount (페이지 개수) : <%=SearchResult.pageCount%></li>
					</ul>
					<% If code = 0 Then 
						For i=0 To UBound(SearchResult.list)-1 %>
						<fieldset class="fieldset2">
							<legend> 현금영수증 조회 결과 [<%= i+1 %> / <%= SearchResult.total %>]</legend>
							<ul>
								<li>itemKey (현금영수증 아이템키) : <%=SearchResult.list(i).itemKey%></li>
								<li>confirmNum (국세청 승인번호) : <%=SearchResult.list(i).confirmNum%></li>
								<li>mgtKey (문서번호) : <%=SearchResult.list(i).mgtKey%></li>
								<li>tradeDate (거래일자) : <%=SearchResult.list(i).tradeDate%></li>
								<li>issueDT (발행일시) : <%=SearchResult.list(i).issueDT%></li>
								<li>regDT (등록일시) : <%=SearchResult.list(i).regDT%></li>
								<li>taxationType (과세형태) : <%=SearchResult.list(i).taxationType%></li>
								<li>totalAmount (거래금액) : <%=SearchResult.list(i).totalAmount%></li>
								<li>tradeUsage (거래구분) : <%=SearchResult.list(i).tradeUsage%></li>
								<li>tradeOpt (거래유형) : <%=SearchResult.list(i).tradeOpt%></li>
								<li>tradeType (문서형태) : <%=SearchResult.list(i).tradeType%></li>
								<li>stateCode (상태코드) : <%=SearchResult.list(i).stateCode%></li>
								<li>stateDT (상태변경일시) : <%=SearchResult.list(i).stateDT%></li>
								<li>stateMemo (상태메모) : <%=SearchResult.list(i).stateMemo%></li>
								<li>identityNum (거래처 식별번호) : <%=SearchResult.list(i).identityNum%></li>
								<li>itemName (상품명) : <%=SearchResult.list(i).itemName%></li>
								<li>customerName (고객명) : <%=SearchResult.list(i).customerName%></li>
								<li>ntssendDT (국세청 전송일시) : <%=SearchResult.list(i).ntssendDT%></li>
								<li>ntsresultDT (국세청 처리결과 수신일시) : <%=SearchResult.list(i).ntsResultDT%></li>
								<li>ntsresultCode (국세청 처리결과 상태코드) : <%=SearchResult.list(i).ntsResultCode%></li>
								<li>ntsresultMessage (국세청 처리결과 메시지) : <%=SearchResult.list(i).ntsResultMessage%></li>
								<li>orgConfirmNum (원본 현금영수증 국세청승인번호) : <%=SearchResult.list(i).orgConfirmNum%></li>
								<li>orgTradeDate (원본 현금영수증 거래일자) : <%=SearchResult.list(i).orgTradeDate%></li>
								<li>printYN (인쇄여부) : <%=SearchResult.list(i).printYN%></li>
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