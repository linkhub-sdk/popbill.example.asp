<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 검색조건을 사용하여 수집 결과 요약정보를 조회합니다.
	' - 응답항목에 관한 정보는 "[홈택스 전자(세금)계산서 연계 API 연동매뉴얼]
	'   > 3.3.2. Summary (수집 결과 요약정보 조회)" 을 참고하시기 바랍니다.
	'**************************************************************

	'팝빌회원 사업자번호, "-" 제외
	testCorpNum = "6798700433"	
	
	'팝빌회원 아이디
	UserID = "testkorea_linkhub"	
	
	'수집 요청(requestJob) 시 반환받은 작업아이디(jobID)
	JobID = "016111416000000024"

	'문서형태 배열, N-일반 전자세금계산서, M-수정 전자세금계산서 
	Dim TIType(2) 
	TIType(0) = "N"
	TIType(1) = "M"

	'과세형태 배열,  T-과세, N-면세, Z-영세
	Dim TaxType(3)
	TaxType(0) = "T"
	TaxType(1) = "N"
	TaxType(2) = "Z"
	
	'영수/청구 배열, R-영수, C-청구, N-없음
	Dim PurposeType(3)
	PurposeType(0) = "R"
	PurposeType(1) = "C"
	PurposeType(2) = "N"

	'종사업장 유무, 공백-전체조회, 0-종사업장번호 없음, 1-종사업장번호 조회
	TaxRegIDYN = ""

	'종사업장 사업자 유형, S-공급자, B-공급받는자, T-수탁자
	TaxRegIDType = "S"

	'종사업장번호, 콤마(",")로 구분하여 구성 ex) 1234,1001
	TaxRegID = ""
	
	On Error Resume Next

	Set result = m_HTTaxinvoiceService.Summary(testCorpNum, JobID, TIType, TaxType,  _
							PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID)

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
						<li> count (수집 결과 건수) : <%=result.count%> </li>
						<li> supplyCostTotal (공급가액 합계) : <%=result.supplyCostTotal%> </li>
						<li> taxTotal (세액 합계) : <%=result.taxTotal%> </li>
						<li> amountTotal (합계 금액) : <%=result.amountTotal%> </li>
					</ul>
				<%
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