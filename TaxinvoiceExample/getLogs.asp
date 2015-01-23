<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	testCorpNum = "1234567890"	'ȸ�� ����ڹ�ȣ, "-" ����
	KeyType= "SELL"             '�������� SELL(����), BUY(����), TRUSTEE(����Ź)
	MgtKey = "20150122-13"      '����������ȣ 

	Set result = m_TaxinvoiceService.GetLogs(testCorpNum, KeyType, MgtKey)
	
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
				<legend> �����̷�Ȯ�� </legend>
				<%
					If code = 0 Then
						For i=0 To result.Count -1 %>
						 <fieldset class="fieldset2">
							<ul>
								<li> DocLogType :  <%=result.Item(i).DocLogType%> </li>
								<li> Log : <%=result.Item(i).Log %> </li>
								<li> ProcType : <%=result.Item(i).ProcType%> </li>
								<li> ProcCorpName : <%=result.Item(i).ProcCorpName%></li>
								<li> ProcMemo : <%=result.Item(i).ProcMemo %></li>
								<li> regDT : <%=result.Item(i).regDT %></li>
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