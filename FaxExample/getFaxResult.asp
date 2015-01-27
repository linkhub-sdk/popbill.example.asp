<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		  '�˺� ȸ�� ����ڹ�ȣ, "-" ����
	userID = "testkorea"			  '�˺� ȸ�� ���̵�
	receiptNum = "015012713201000001" '�ѽ� ���۽� �߱޹��� ���۹�ȣ
 
 	'���۰���ڵ�� [�˺� FAX API �����Ŵ��� 5.�η�] ����
	
	On Error Resume Next

	Set result = m_FaxService.GetFaxDetail(testCorpNum, receiptNum, userID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	Else 
		code = Presponse.code
		message = Presponse.message
	End If

	On Error GoTo 0
%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>�ѽ����� ���۰�� Ȯ�� </legend>
				<% If code = 0 Then %>
					<ul>
						<li>sendState(���ۻ���) : <%=result.sendState%> </li>
						<li>convState(��ȯ����) : <%=result.convState%> </li>
						<li>sendNum(�߽Ź�ȣ) : <%=result.sendNum%> </li>
						<li>receiveNum(���Ź�ȣ) : <%=result.receiveNum%> </li>
						<li>receiveName(�����ڸ�) : <%=result.receiveName%> </li>
						<li>sendPageCnt(��������) : <%=result.sendPageCnt%></li>
						<li>successPageCnt(���� ��������) : <%=result.successPageCnt%></li>
						<li>failPageCnt(���� ��������) : <%=result.failPageCnt%></li>
						<li>refundPageCnt(ȯ�� ��������) : <%=result.refundPageCnt%></li>
						<li>cancelPageCnt(��� ��������) : <%=result.cancelPageCnt%></li>
						<li>reserveDT(����ð�) : <%=result.reserveDT%></li>
						<li>sendDT(�߼۽ð�) : <%=result.sendDT%></li>
						<li>sendResult(��Ż� ó�����) : <%=result.sendResult%></li>
					</ul>
				<%	Else  %>
					<ul>
						<li>Response.code: <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					</ul>	
				<%	End If	%>
			</fieldset>
		 </div>
	</body>
</html>