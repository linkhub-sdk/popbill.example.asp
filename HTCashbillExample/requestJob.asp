<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ���ݿ����� ����/���� ���� ������ ��û�մϴ�
	' - ����/���� ���� ���μ����� "[Ȩ�ý� ���ݿ����� ���� API �����Ŵ���]
	'   > 1.2. ���μ��� �帧��" �� �����Ͻñ� �ٶ��ϴ�.
	' - ���� ��û�� ��ȯ���� �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	'�������� SELL(����), BUY(����)
	KeyType= "BUY"						

	'��������, ǥ������(yyyyMMdd)
	SDate = "20161001"

	'��������, ǥ������(yyyyMMdd)
	EDate =	"20161131"					

	'�˺�ȸ�� ���̵�
	testUserID = "testkorea"			
	
	On Error Resume Next

	jobID = m_HTCashbillService.requestJob(testCorpNum, KeyType, SDate, EDate, testUserID)

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
				<legend>���� ��û</legend>
				<% If code = 0 Then %>
					<ul>
						<li>jobID(�۾����̵�) : <%=jobID%> </li>
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