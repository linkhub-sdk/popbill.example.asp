<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "4108600477"		' [필수] 팝빌회원 사업자번호, "-" 제외
	UserID = "innoposttest"
	KeyType= "SELL"						' [필수] 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	DType = "R"								' [필수] 검색일자 유형, R-등록일자, W-작성일자, I-발행일자
	SDate = "20160501"					' [필수] 시작일자, yyyyMMdd
	EDate = "20160731"					' [필수] 종료일자, yyyyMMdd
	
	' 전송상태값 배열, 미기지새 전체조회, 문서상태값 3자리 배열, 2,3번째 자리 와일드카드 사용가능
	Dim State(2)
	State(0) = "2**"
	State(1) = "3**"

	
	' 문서유형 배열, N-일반세금계산서, M-수정세금계산서  중 선택배열
	Dim TIType(2)
	TIType(0) = "N"
	TIType(1) = "M"

	' 과세형태 배열, T-과세, N-면세, Z-영세 중 선택 배열
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"

	LateOnly = null		' 지연발행여부,  null- 전체조회, False-정상발행분 조회, True-지연발행분 조회

	Order = "D"			' 정렬방향, A-오름차순, D-내림차순
	Page = 1				' 페이지 번호
	PerPage = 100		' 페이지당 검색갯수, 최대 1000

	'종사업장번호 사업자유형, S-매출, B-매입, T-수탁
	TaxRegIDType = "S"

	'종사업장번호 유무, 공백-전체조회, 0-종사업장번호 없음, 1-종사업장번호 있음
	TaxRegIDYN = ""
	
	'종사업장번호, 콤마(",")로 구분하여 구성 ex) "1234,0001"
	TaxRegID = ""

	Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, TIType, TaxType, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, TaxRegID, UsreID)

	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>세금계산서 목록조회 결과</legend>
						<ul>
							<li> code : <%=result.code%></li>
							<li> total : <%=result.total%></li>
							<li> pageNum : <%=result.pageNum%></li>
							<li> perPage : <%=result.perPage%></li>
							<li> pageCount : <%=result.pageCount%></li>
							<li> message : <%=result.message%></li>
						</ul>

				<%
					If code = 0 Then
						For i=0 To UBound(result.list) -1
				%>
							<fieldset class="fieldset2">					
								<legend>  Search.List [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
									<ul>
										<li> itemKey :  <%=result.list(i).itemKey%> </li>
										<li> stateCode :  <%=result.list(i).stateCode%> </li>
										<li> taxType :  <%=result.list(i).taxType%> </li>
										<li> purposeType :  <%=result.list(i).purposeType%> </li>
										<li> issueType :  <%=result.list(i).issueType %> </li>
										<li> writeDate :  <%=result.list(i).writeDate%> </li>
										<li> invoicerCorpName :  <%=result.list(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum :  <%=result.list(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey :  <%=result.list(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN :  <%=result.list(i).invoicerPrintYN%> </li>
										<li> invoiceeCorpName :  <%=result.list(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum :  <%=result.list(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey :  <%=result.list(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN :  <%=result.list(i).invoiceePrintYN%> </li>
										<li> trusteeCorpName :  <%=result.list(i).trusteeCorpName%> </li>
										<li> trusteeCorpNum :  <%=result.list(i).trusteeCorpName%> </li>
										<li> trusteeMgtKey :  <%=result.list(i).trusteeMgtKey%> </li> 
										<li> trusteePrintYN :  <%=result.list(i).trusteePrintYN%> </li>
										<li> supplyCostTotal :  <%=result.list(i).supplyCostTotal%> </li>
										<li> taxTotal :  <%=result.list(i).taxTotal%> </li>
										<li> issueDT :  <%=result.list(i).issueDT%> </li>
										<li> preIssueDT :  <%=result.list(i).preIssueDT%> </li>
										<li> stateDT :  <%=result.list(i).stateDT%> </li>
										<li> openYN :  <%=result.list(i).openYN%> </li>
										<li> openDT :  <%=result.list(i).openDT%> </li>
										<li> ntsresult :  <%=result.list(i).ntsresult%> </li>
										<li> ntsconfirmNum :  <%=result.list(i).ntsconfirmNum %> </li>
										<li> ntssendDT :  <%=result.list(i).ntssendDT%> </li>
										<li> ntsresultDT :  <%=result.list(i).ntsresultDT%> </li>
										<li> ntssendErrCode :  <%=result.list(i).ntssendErrCode%> </li>
										<li> stateMemo :  <%=result.list(i).stateMemo%> </li>
										<li> regDT :  <%=result.list(i).regDT%> </li>
										<li> lateIssueYN :  <%=result.list(i).lateIssueYN%> </li>
									</ul>
								</fieldset>
				<%
						Next
					Else
				%>
				</fieldset>
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
