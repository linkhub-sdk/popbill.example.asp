<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
		<link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
		<title>�˺� SDK ASP Example.</title>
	</head>
<!--#include file="common.asp"--> 
<%
	'**************************************************************
	' ����ȸ���� ����� ����� Ȯ���մϴ�.	
	'**************************************************************

	' �˺�ȸ�� ����ڹ�ȣ, "-" ����
	testCorpNum = "1234567890"		

	' �˺�ȸ�� ���̵�
	UserID = "testkorea"					
	
	Set result = m_TaxinvoiceService.ListContact(testCorpNum, UserID)

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
				<legend>����� ��� ��ȸ</legend>
				<%
					If code = 0 Then
						For i=0 To result.Count-1
				%>
							<fieldset class="fieldset2">					
								<legend> ContactInfoList [ <%=i+1%> / <%=result.Count%> ] </legend>
									<ul>
										<li> id : <%=result.Item(i).id%></li>
										<li> email : <%=result.Item(i).email%></li>
										<li> hp : <%=result.Item(i).hp%></li>
										<li> personName : <%=result.Item(i).personName%></li>
										<li> searchAllAllowYN : <%=result.Item(i).searchAllAllowYN%></li>
										<li> tel : <%=result.Item(i).tel%></li>
										<li> fax : <%=result.Item(i).fax%></li>
										<li> mgrYN : <%=result.Item(i).mgrYN%></li>
										<li> regDT : <%=result.Item(i).regDT%></li>
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
