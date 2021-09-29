<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	'  ���� �ŷ����� ������ ��û�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 1����)
	' - ��ȸ�Ϸκ��� �ִ� 3���� ���� �������� ��ȸ�� �� �ֽ��ϴ�.
	' - ���� ��û�� ��ȯ���� �۾����̵�(JobID)�� ��ȿ�ð��� 1�ð� �Դϴ�.
	' - https://docs.popbill.com/easyfinbank/asp/api#RequestJob
	'**************************************************************

	'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"	

	'�˺�ȸ�� ���̵�
	testUserID = "testkorea"		
	
	'�����ڵ�
	BankCode = "0039"

	'���¹�ȣ
	AccountNumber = "2070064402404"		

	'��������, ǥ������(yyyyMMdd)
	SDate = "20200701"			
	
	'��������, ǥ������(yyyyMMdd)
	EDate =	"20200724"				

	
	On Error Resume Next

	jobID = m_EasyFinBankService.requestJob(testCorpNum, BankCode, AccountNumber, SDate, EDate, testUserID)

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