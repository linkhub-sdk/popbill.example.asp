<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>팝빌 SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'팝빌 회원 사업자번호, "-"제외 10자리
	userID = "testkorea"				'팝빌 회원 아이디
	itemCode = "121"					'명세서 코드 - 121(거래명세서), 122(청구서), 123(견적서) 124(발주서), 125(입금표), 126(영수증)
	mgtKey = "20150201-01"				'연동관리번호

	On Error Resume Next
	
	Set result = m_StatementService.GetDetailInfo(testCorpNum, itemCode, mgtKey, userID)
	
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
				<legend>전자명세서 상세정보</legend>
				<ul>
					<% If code = 0 Then %>
						<li>writeDate : <%=result.writeDate%> </li>
						<li>taxType : <%=result.taxType%> </li>
						<li>senderCorpName : <%=result.senderCorpName%> </li>
						<li>senderCorpNum : <%=result.senderCorpNum%> </li>
						<li>senderAddr : <%=result.senderAddr%> </li>
						<li>senderBizClass : <%=result.senderBizClass%> </li>
						<li>senderBizType : <%=result.senderBizType%> </li>
						<li>sendercontactName : <%=result.sendercontactName%> </li>
						<li>senderDeptName : <%=result.senderDeptName%> </li>
						<li>senderTEL : <%=result.senderTEL%> </li>
						<li>senderHP : <%=result.senderHP%> </li>
						<li>senderEmail : <%=result.senderEmail%> </li>

						<li>receiverCorpName : <%=result.receiverCorpName%> </li>
						<li>receiverCorpNum : <%=result.receiverCorpNum%> </li>
						<li>receiverAddr : <%=result.receiverAddr%> </li>
						<li>receiverBizClass : <%=result.receiverBizClass%> </li>
						<li>receiverBizType : <%=result.receiverBizType%> </li>
						<li>receivercontactName : <%=result.receivercontactName%> </li>
						<li>receiverDeptName : <%=result.receiverDeptName%> </li>
						<li>receiverTEL : <%=result.receiverTEL%> </li>
						<li>receiverHP : <%=result.receiverHP%> </li>
						<li>receiverEmail : <%=result.receiverEmail%> </li>
						<li>taxTotal : <%=result.taxTotal %> </li>
						<li>supplyCostTotal : <%=result.supplyCostTotal %> </li>
						<li>totalAmount : <%=result.totalAmount %> </li>
						<li>purposeType : <%=result.purposeType %> </li>
						<li>serialNum : <%=result.serialNum %> </li>
						<li>remark1 : <%=result.remark1 %> </li>
						<li>remark2 : <%=result.remark2 %> </li>
						<li>remark3 : <%=result.remark3 %> </li>
						
						<!--기타 상세항목 생략-->

						<fieldset class="fieldset2">
							<legend>추가속성</legend>
							<ul>
							<% For Each propertyKey In result.propertyBag.keys() %>
								<li> <%=propertyKey%> : <%=result.propertyBag.get(propertyKey)%></li>
							<% Next %>
							</ul>
						</fieldset>
						<% For i=0 To Ubound(result.detailList)-1%>
								<fieldset class="fieldset2">
								<legend> 상세항목 <%=i+1%> </legend>
									<ul>
										<li> serialNum : <%=result.detailList(i).serialNum%> </li>
										<li> itemName : <%=result.detailList(i).itemName%> </li>
										<li> supplyCost : <%=result.detailList(i).supplyCost%> </li>
									</ul>
								</fieldset>
							<% 
								Next
								Else
							%>
		
							<li>Response.code : <%=code%> </li>
							<li>Response.message: <%=message%> </li>
						<% 
							End If
						%>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>