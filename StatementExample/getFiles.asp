<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ڸ����� ÷�ε� ������ ����� Ȯ���մϴ�.
	' - �����׸� �� ���Ͼ��̵�(AttachedFile) �׸��� ���ϻ���(DeleteFile API)
	'   ȣ��� �̿��� �� �ֽ��ϴ�.
	'**************************************************************

	'�˺� ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
	testCorpNum = "1234567890"			

	'�˺� ȸ�� ���̵�
	userID = "testkorea"				

	'���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������) 124(���ּ�), 125(�Ա�ǥ), 126(������)
	itemCode = "121"					

	'����������ȣ
	mgtKey = "20161114-06"				

	On Error Resume Next

	Set result = m_StatementService.GetFiles(testCorpNum, itemCode, mgtKey, userID)

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
				<legend>÷������ ��� Ȯ��</legend>
				<ul>
					<% If code = 0 Then 
						   For i=0 To result.length-1
					%>
						<fieldset class="fieldset2">
							<legend>÷������ [<%=i+1%>] </legend>
							<ul>
								<li>SerialNum : <%=result.Get(i).SerialNum%></li>
								<li>AttachedFile : <%=result.Get(i).AttachedFile%></li>
								<li>DisplayName : <%=result.Get(i).DisplayName%></li>
								<li>regDT : <%=result.Get(i).regDT%></li>
							</ul>
						</fieldset>
					<% 
						Next
						Else
					%>
						<li>Response.code : <%=code%> </li>
						<li>Response.message: <%=message%> </li>
					<% End If %>
				</ul>
			</fieldset>
		 </div>
	</body>
</html>