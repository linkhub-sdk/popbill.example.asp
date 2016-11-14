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
	MgtKeyList(0) = "20161114-01"
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
								<li>itemKey : <%=Presponse.Item(i).itemKey%></li>
								<li>mgtKey : <%=Presponse.Item(i).mgtKey%></li>
								<li>tradeDate : <%=Presponse.Item(i).tradeDate%></li>
								<li>issueDT : <%=Presponse.Item(i).issueDT%></li>
								<li>customerName : <%=Presponse.Item(i).customerName%></li>
								<li>itemName : <%=Presponse.Item(i).itemName%></li>
								<li>identityNum : <%=Presponse.Item(i).identityNum%></li>
								<li>taxactionType : <%=Presponse.Item(i).taxationType%></li>
								<li>totalAmount : <%=Presponse.Item(i).totalAmount%></li>
								<li>tradeUsage : <%=Presponse.Item(i).tradeUsage%></li>
								<li>tradeType : <%=Presponse.Item(i).tradeType%></li>
								<li>stateCode : <%=Presponse.Item(i).stateCode%></li>
								<li>stateDT : <%=Presponse.Item(i).stateDT%></li>
								<li>printYN : <%=Presponse.Item(i).printYN%></li>
								<li>confirmNum : <%=Presponse.Item(i).confirmNum%></li>
								<li>orgTradeDate : <%=Presponse.Item(i).orgTradeDate%></li>
								<li>orgConfirmNum : <%=Presponse.Item(i).orgConfirmNum%></li>
								<li>ntssendDT : <%=Presponse.Item(i).ntssendDT%></li>
								<li>ntsPresponse : <%=Presponse.Item(i).ntsResult%></li>
								<li>ntsPresponseDT : <%=Presponse.Item(i).ntsResultDT%></li>
								<li>ntsPresponseCode : <%=Presponse.Item(i).ntsResultCode%></li>
								<li>ntsPresponseMessage : <%=Presponse.Item(i).ntsResultMessage%></li>
								<li>regDT : <%=Presponse.Item(i).regDT%></li>
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