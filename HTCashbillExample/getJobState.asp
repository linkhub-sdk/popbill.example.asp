<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'�˺�ȸ�� ����ڹ�ȣ, "-" ����
	JobID = "016071511000000009"	'������û�� ��ȯ�����۾����̵�(jobID)
	UserID = "testkorea"					'�˺�ȸ�� ���̵�
	
	On Error Resume Next

	Set result = m_HTCashbillService.GetJobState(testCorpNum, JobID, UserID)

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
				<legend>���� ���� Ȯ��</legend>
				<%
					If code = 0 Then
				%>
						<ul>
							<li> jobID (�۾����̵�) : <%=result.jobID%></li>
							<li> jobState (��������) : <%=result.jobState%></li>
							<li> queryType (��������) : <%=result.queryType%></li>
							<li> queryDateType (��������) : <%=result.queryDateType%></li>
							<li> queryStDate (��������) : <%=result.queryStDate%></li>
							<li> queryEnDate (��������) : <%=result.queryEnDate%></li>
							<li> errorCode (�����ڵ�) : <%=result.errorCode%></li>
							<li> errorReason (�����޽���) : <%=result.errorReason%></li>
							<li> jobStartDT (�۾� �����Ͻ�) : <%=result.jobStartDT%></li>
							<li> jobEndDT (�۾� �����Ͻ�) : <%=result.jobEndDT%></li>
							<li> collectCount (��������) : <%=result.collectCount%></li>
							<li> regDT (���� ��û�Ͻ�) : <%=result.regDT%></li>
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
