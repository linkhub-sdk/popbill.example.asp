<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'회원 사업자번호, "-" 제외
	KeyType= "SELL"             '발행유형 SELL(매출), BUY(매입), TRUSTEE(위수탁)
	UserID = "testkorea"	  '회원아이디

	'세금계산서 연동관리번호 배열, 최대 1000건
	Dim MgtKeyList(3) 
	MgtKeyList(0) = "20150121-01"
	MgtKeyList(1) = "20150121-02"
	MgtKeyList(2) = "20150121-03"
	
	Set result = m_TaxinvoiceService.GetInfos(testCorpNum, KeyType, MgtKeyList, UserID)

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
				<legend>팝빌 세금계산서 상태/요약 다량 요청</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend> TaxinvoiceResult : <%=i+1%> </legend>
									<ul>
										<li> itemKey : <%=result.Item(i).itemKey%></li>
										<li> stateCode : <%=result.Item(i).stateCode%></li>
										<li> taxType : <%=result.Item(i).taxType%></li>
										<li> writeDate : <%=result.Item(i).writeDate%></li>
										<li> regDT : <%=result.Item(i).regDT%></li>
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
