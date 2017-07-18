<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' 다수건의 현금영수증 상태/요약 정보를 확인합니다. (최대 1000건)
	' - 응답항목에 대한 자세한 정보는 "[현금영수증 API 연동매뉴얼] > 4.2.
	'   현금영수증 상태정보 구성"을 참조하시기 바랍니다.
	'**************************************************************
	
	'팝빌 회원 사업자번호, "-" 제외
	testCorpNum = "1234567890"	 

	'팝빌 회원 아이디
	userID = "testkorea"		 

	'조회할 현금영수증 문서관리번호 배열, 최대 1000건
	Dim mgtKeyList(3) 
	MgtKeyList(0) = "20170718-04"
	MgtKeyList(1) = "20150129-05"
	MgtKeyList(2) = "20150129-06"

	On Error Resume Next
	
	Set Presponse = m_CashbillService.GetInfos(testCorpNum, mgtKeyList, userID)

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
				<legend>현금영수증 상태 대량 확인</legend>
				<ul>
					<% If code = 0 Then 
						For i=0 To Presponse.Count-1 %>
						<fieldset class="fieldset2">
							<legend> 현금영수증 조회 결과 [<%=i+1%>]</legend>
							<ul>
								<li>itemKey (현금영수증 아이템키) : <%=Presponse.Item(i).itemKey%></li>
								<li>mgtKey (문서관리번호) : <%=Presponse.Item(i).mgtKey%></li>
								<li>tradeDate (거래일자) : <%=Presponse.Item(i).tradeDate%></li>
								<li>issueDT (발행일시) : <%=Presponse.Item(i).issueDT%></li>
								<li>regDT (등록일시) : <%=Presponse.Item(i).regDT%></li>
								<li>taxationType (과세형태) : <%=Presponse.Item(i).taxationType%></li>
								<li>totalAmount (거래금액) : <%=Presponse.Item(i).totalAmount%></li>
								<li>tradeUsage (거래용도) : <%=Presponse.Item(i).tradeUsage%></li>
								<li>tradeType (현금영수증 형태) : <%=Presponse.Item(i).tradeType%></li>
								<li>stateCode (상태코드) : <%=Presponse.Item(i).stateCode%></li>
								<li>stateDT (상태변경일시) : <%=Presponse.Item(i).stateDT%></li>

								<li>identityNum (거래처 식별번호) : <%=Presponse.Item(i).identityNum%></li>
								<li>customerName (고객명) : <%=Presponse.Item(i).customerName%></li>
								<li>itemName (상품명) : <%=Presponse.Item(i).itemName%></li>

								<li>confirmNum (국세청승인번호) : <%=Presponse.Item(i).confirmNum%></li>
								<li>ntssendDT (국세청 전송일시) : <%=Presponse.Item(i).ntssendDT%></li>
								<li>ntsresultDT (국세청 처리결과 수신일시) : <%=Presponse.Item(i).ntsResultDT%></li>
								<li>ntsresultCode (국세청 처리결과 상태코드) : <%=Presponse.Item(i).ntsResultCode%></li>
								<li>ntsresultMessage (국세청 처리결과 메시지) : <%=Presponse.Item(i).ntsResultMessage%></li>
								<li>orgTradeDate (원본 현금영수증 거래일자) : <%=Presponse.Item(i).orgTradeDate%></li>
								<li>orgConfirmNum (원본 현금영수증 국세청승인번호) : <%=Presponse.Item(i).orgConfirmNum%></li>

								<li>printYN (인쇄여부) : <%=Presponse.Item(i).printYN%></li>
							</ul>
						</fieldset>
					<%	Next
						Else %>
						<li> Response.code : <%=code%> </li>
						<li> Response.message : <%=message%> </li>
					<% End If%> 
					
				</ul>
			</fieldset>
		 </div>
	</body>
</html>