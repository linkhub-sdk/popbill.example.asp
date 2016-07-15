<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'ȸ�� ����ڹ�ȣ, "-" ����
	KeyType= SELL						'�������� SELL(����), BUY(����), TRUSTEE(����Ź)
	DType = "W"							'�˻� ��������, W-�ۼ�����, I-��������, S-��������
	SDate = "20160601"					'��������, ǥ������(yyyyMMdd)
	EDate =	"20160831"					'��������, ǥ������(yyyyMMdd)
	testUserID = "testkorea"		'ȸ�� ���̵�
	
	On Error Resume Next

	'������û�� ��ȯ�Ǵ� jobID�� ��ȿ�ð��� 1�ð� �Դϴ�.
	jobID = m_HTTaxinvoiceService.requestJob(testCorpNum, KeyType, DType, SDate, EDate, testUserID)

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