<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ������ �Ϸ�� ������ �ŷ����� ��������� ��ȸ�մϴ�.
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	
	
	'�˺�ȸ�� ���̵�
	UserID = "testkorea"	
	
	'���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
	JobID = "019123114000000010"

	'�ŷ����� �迭, I-�Ա�, O-���
	Dim TradeType(2) 
	TradeType(0) = "I"
	TradeType(1) = "O"

	'��ȸ �˻���, �Ա�/��ݾ�, �޸�, ���� like �˻�
	SearchString = ""

	On Error Resume Next

	Set result = m_EasyFinBankService.Summary(testCorpNum, JobID, TradeType, SearchString, UserID)

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
				<legend>���� ��� ������� ��ȸ</legend>
				<%
					If code = 0 Then
				%>
					<ul>
						<li> count (���� ��� �Ǽ�) : <%=result.count%> </li>
						<li> cntAccIn (�Աݰŷ� �Ǽ�) : <%=result.cntAccIn%> </li>
						<li> cntAccOut (��ݰŷ� �Ǽ�) : <%=result.cntAccOut%> </li>
						<li> totalAccIn (�Աݾ� �հ�) : <%=result.totalAccIn%> </li>
						<li> totalAccOut (��ݾ� �հ�) : <%=result.totalAccOut%> </li>
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