<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 세금계산서 목록을 조회합니다.
	' - 응답항목에 대한 자세한 사항은 "[전자세금계산서 API 연동매뉴얼] >
	'   4.2. (세금)계산서 상태정보 구성" 을 참조하시기 바랍니다.
	'**************************************************************

	' 팝빌회원 사업자번호, "-" 제외 10자리
	testCorpNum = "1234567890"
	
	' 팝빌회원 아이디
	UserID = "testkorea"

	' [필수] 발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	KeyType = "SELL"

	' [필수] 검색일자 유형, R-등록일자, W-작성일자, I-발행일자
	DType = "W"
	
	' [필수] 시작일자, yyyyMMdd
	SDate = "20181201"

	' [필수] 종료일자, yyyyMMdd
	EDate = "20190103"
	
	' 전송상태값 배열, 미기지새 전체조회, 문서상태값 3자리 배열, 2,3번째 자리 와일드카드 사용가능
	Dim State(2)
	State(0) = "3**"
	State(1) = "6**"

	
	' 문서유형 배열, N-일반세금계산서, M-수정세금계산서  중 선택배열
	Dim TIType(2)
	TIType(0) = "N"
	TIType(1) = "M"

	' 과세형태 배열, T-과세, N-면세, Z-영세 중 선택 배열
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"

	' 과세형태 배열, T-과세, N-면세, Z-영세 중 선택 배열
	Dim IssueType(3)
	IssueType(0) = "N"
	IssueType(1) = "R"
	IssueType(2) = "T"

	' 지연발행여부,  null- 전체조회, False-정상발행분 조회, True-지연발행분 조회
	LateOnly = null		

	' 정렬방향, A-오름차순, D-내림차순
	Order = "D"

	' 페이지 번호
	Page = 1

	' 페이지당 검색갯수, 최대 1000
	PerPage = 5

	'종사업장번호 사업자유형, S-매출, B-매입, T-수탁
	TaxRegIDType = "S"

	'종사업장번호 유무, 공백-전체조회, 0-종사업장번호 없음, 1-종사업장번호 있음
	TaxRegIDYN = ""
	
	'종사업장번호, 콤마(",")로 구분하여 구성 ex) "1234,0001"
	TaxRegID = ""

	'거래처 정보, 거래처 상호 또는 사업자등록번호 기재, 공백처리시 전체조회
	QString = ""

	'연동문서 조회여부, 공백-전체조회, 0-일반문서 조회, 1-연동문서 조회
	InterOPYN = ""

	On Error Resume Next

	Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, _ 
						TIType, TaxType, IssueType, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, _
						TaxRegID, QString, InterOPYN, UsreID)

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
				<%
					If code = 0 Then
				%>
						<legend>세금계산서 목록조회</legend>
						<ul>
							<li> code (응답코드) : <%=result.code%></li>
							<li> message (응답메시지) : <%=result.message%></li>
							<li> total (총 검색결과 건수) : <%=result.total%></li>
							<li> pageNum (페이지 번호) : <%=result.pageNum%></li>
							<li> perPage (페이지당 목록개수) : <%=result.perPage%></li>
							<li> pageCount (페이지 개수) : <%=result.pageCount%></li>
						</ul>
						<%
							For i=0 To UBound(result.list) -1
						%>
							<fieldset class="fieldset2">					
								<legend>  세금계산서 상태/요약정보 [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
									<ul>
										<li> itemKey (세금계산서 아이템키) :  <%=result.list(i).itemKey%> </li>
										<li> stateCode (상태코드) :  <%=result.list(i).stateCode%> </li>
										<li> taxType (과세형태) :  <%=result.list(i).taxType%> </li>
										<li> purposeType (영수/청구) :  <%=result.list(i).purposeType%> </li>
										<li> issueType (발행형태) :  <%=result.list(i).issueType %> </li>
										<li> writeDate (작성일자) :  <%=result.list(i).writeDate%> </li>

										<li> invoicerCorpName (공급자 상호) :  <%=result.list(i).invoicerCorpName%> </li>
										<li> invoicerCorpNum (공급자 사업자번호) :  <%=result.list(i).invoicerCorpNum%> </li>
										<li> invoicerMgtKey (공급자 문서관리번호) :  <%=result.list(i).invoicerMgtKey%> </li>
										<li> invoicerPrintYN (공급자 인쇄여부) :  <%=result.list(i).invoicerPrintYN%> </li>
										
										<li> invoiceeCorpName (공급받는자 상호) :  <%=result.list(i).invoiceeCorpName%> </li>
										<li> invoiceeCorpNum (공급받는자 사업자번호) :  <%=result.list(i).invoiceeCorpNum%> </li>
										<li> invoiceeMgtKey (공급받는자 문서관리번호) :  <%=result.list(i).invoiceeMgtKey%> </li>
										<li> invoiceePrintYN (공급받는자 인쇄여부) :  <%=result.list(i).invoiceePrintYN%> </li>
										<li> closeDownState (공급받는자 휴폐업상태) :  <%=result.list(i).closeDownState%> </li>
										<li> closeDownStateDate (공급받는자 휴폐업일자) :  <%=result.list(i).closeDownStateDate%> </li>

										<li> interOPYN (연동문서 여부) :  <%=result.list(i).interOPYN%> </li>
										<li> supplyCostTotal (공급가액 합계) :  <%=result.list(i).supplyCostTotal%> </li>
										<li> taxTotal (세액 합계) :  <%=result.list(i).taxTotal%> </li>
										<li> issueDT (발행일시) :  <%=result.list(i).issueDT%> </li>

										<li> stateDT (상태 변경일시) :  <%=result.list(i).stateDT%> </li>
										<li> openYN (개봉 여부) :  <%=result.list(i).openYN%> </li>
										<li> openDT (개봉 일시) :  <%=result.list(i).openDT%> </li>
										<li> ntsresult (국세청 전송결과) :  <%=result.list(i).ntsresult%> </li>
										<li> ntsconfirmNum (국세청 승인번호) :  <%=result.list(i).ntsconfirmNum %> </li>
										<li> ntssendDT (국세청 전송일시) :  <%=result.list(i).ntssendDT%> </li>
										<li> ntsresultDT (국세청 결과 수신일시) :  <%=result.list(i).ntsresultDT%> </li>
										<li> ntssendErrCode (전송실패 사유코드) :  <%=result.list(i).ntssendErrCode%> </li>

										<li> stateMemo (상태메모) :  <%=result.list(i).stateMemo%> </li>
										<li> regDT (등록일시) :  <%=result.list(i).regDT%> </li>
										<li> lateIssueYN (지연발행 여부) :  <%=result.list(i).lateIssueYN%> </li>
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
