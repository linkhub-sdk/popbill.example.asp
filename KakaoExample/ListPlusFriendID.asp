<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' īī���� �÷���ģ�� ���� ����� Ȯ���մϴ�.
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		 
	
	On Error Resume Next

	Set Presponse = m_KakaoService.ListPlusFriendID(testCorpNum)

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
				<legend>īī���� �÷���ģ�� ���� ��� Ȯ��</legend>
				<%
					For i=0 To Presponse.length -1
				%>
				<fieldset class="fieldset2">
				<ul>
					<li>�÷���ģ�� ���̵� (plusFriendID) : <%=Presponse.Get(i).plusFriendID%> </li>
					<li>�÷���ģ�� �̸�(plusFriendName) : <%=Presponse.Get(i).plusFriendName%> </li>
					<li>����Ͻ�(regDT) : <%=Presponse.Get(i).regDT%> </li>
				</ul>
				</fieldset>
				<%
					Next
				%>

		 </div>
	</body>
</html>