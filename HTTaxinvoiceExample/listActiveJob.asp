<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"		'����ȸ�� ����ڹ�ȣ, "-" ����
	UserID = "testkorea"					'����ȸ�� ���̵�
	
	'������û�� ��ȯ�Ǵ� jobID�� ��ȿ�ð��� 1�ð� �Դϴ�.
	Set result = m_HTTaxinvoiceService.ListActiveJob(testCorpNum, UserID)
	
	If Err.Number <> 0 Then
		code = Err.Number
		message = Err.Description
		Err.Clears
	End If

%>
	<body>
		<div id="content">
			<p class="heading1">Response</p>
			<br/>
			<fieldset class="fieldset1">
				<legend>���� ��� ��ȸ</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend>ListActiveJob [ <%=i+1%> / <%=result.Count%> ] </legend>
									<ul>
										<li> jobID (�۾����̵�) : <%=result.Item(i).jobID%></li>
										<li> jobState (��������) : <%=result.Item(i).jobState%></li>
										<li> queryType (��������) : <%=result.Item(i).queryType%></li>
										<li> queryDateType (��������) : <%=result.Item(i).queryDateType%></li>
										<li> queryStDate (��������) : <%=result.Item(i).queryStDate%></li>
										<li> queryEnDate (��������) : <%=result.Item(i).queryEnDate%></li>
										<li> errorCode (�����ڵ�) : <%=result.Item(i).errorCode%></li>
										<li> errorReason (�����޽���) : <%=result.Item(i).errorReason%></li>
										<li> jobStartDT (�۾� �����Ͻ�) : <%=result.Item(i).jobStartDT%></li>
										<li> jobEndDT (�۾� �����Ͻ�) : <%=result.Item(i).jobEndDT%></li>
										<li> collectCount (��������) : <%=result.Item(i).collectCount%></li>
										<li> regDT (���� ��û�Ͻ�) : <%=result.Item(i).regDT%></li>
									</ul>
								</fieldset>
				<%
						Next
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