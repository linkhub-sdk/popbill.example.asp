<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"			'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	userID = "testkorea"				'�˺� ȸ�� ���̵�
	itemCode = "121"					'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	mgtKey = "20150201-01"				'����������ȣ

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
				<legend>���ڸ��� ������</legend>
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
						
						<!--��Ÿ ���׸� ����-->

						<fieldset class="fieldset2">
							<legend>�߰��Ӽ�</legend>
							<ul>
							<% For Each propertyKey In result.propertyBag.keys() %>
								<li> <%=propertyKey%> : <%=result.propertyBag.get(propertyKey)%></li>
							<% Next %>
							</ul>
						</fieldset>
						<% For i=0 To Ubound(result.detailList)-1%>
								<fieldset class="fieldset2">
								<legend> ���׸� <%=i+1%> </legend>
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