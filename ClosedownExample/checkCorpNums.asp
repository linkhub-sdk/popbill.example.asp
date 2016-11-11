<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
		<title>�������ȸ API SDK ASP Example.</title>
	</head>
	<!--#include file="common.asp"--> 
	<%
		'**************************************************************
		' �ټ��� ����ڿ� ���� ��������θ� ��ȸ�մϴ�. (�ִ� 1000��)
		'**************************************************************

		'�˺�ȸ�� ����ڹ�ȣ
		UserCorpNum = "1234567890"		

		'��ȸ�� ����ڹ�ȣ �迭, �ִ� 1000��
		Dim CorpNumList(3)
		CorpNumList(0) = "1234567890"
		CorpNumList(1) = "4108600477"
		CorpNumList(2) = "110-04-45791"
						
		On Error Resume Next

		Set result = m_ClosedownService.checkCorpNums(UserCorpnum, CorpNumList)
		
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
				<legend>�������ȸ - �뷮</legend>
				<br/>
				<p class="info">> state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�</p>
				<p class="info">> type (��� ����) : null-�˼�����, 1-�Ϲݰ�����, 2-�鼼������, 3-���̰�����, 4-�񿵸�����, �������</p>
				<br/>
			<%
				If Not IsEmpty(result) Then  
					For i=0 To result.Count-1
			%>
					<fieldset class="fieldset2">
						<legend>�������ȸ - �뷮</legend>
						<ul>
								<li>����ڹ�ȣ(corpNum) : <%= result.Item(i).corpNum%></li>		
								<li>���������(state) : <%= result.Item(i).state%></li>
								<li>���������(type) : <%= result.Item(i).ctype%></li>	
								<li>���������(stateDate) : <%= result.Item(i).stateDate%></li>	
								<li>����û Ȯ������(checkDate) : <%= result.Item(i).checkDate%></li>	
						</ul>
					</fieldset>
			<%
					Next
				End If 
				If Not IsEmpty(code) then
			%>
				<fieldset class="fieldset2">
					<legend>�������ȸ - �ܰ�</legend>
					<ul>
						<li>Response.code : <%= code %> </li>
						<li>Response.message : <%= message %></li>
					</ul>
				</fieldset>
			<%
				End If
			%>		

			</fieldset>

		<script type ="text/javascript">
			 window.onload=function(){
				 document.getElementById('CorpNum').focus();
			 }
			 
			 function search(){
				document.getElementById('corpnum_form').submit();
			 }		 
		 </script>
	</body>
</html>