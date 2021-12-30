<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �뷮�� ���ڸ��� �μ��˾� URL�� ��ȯ�մϴ�. (�ִ� 100��)
    ' - ������å���� ���� ��ȯ�� URL�� ��ȿ�ð��� 30���Դϴ�.
    ' - https://docs.popbill.com/statement/asp/api#GetMassPrintURL
    '**************************************************************

    '�˺� ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	 

    '�˺� ȸ�� ���̵�
    userID = "testkorea"		 

    '���� �ڵ� - 121(�ŷ�����), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = "121"			 

    '������ȣ �迭, �ִ� 100��
    Dim mgtKeyList(2)  
    mgtKeyList(0) = "20211201-001"
    mgtKeyList(1) = "20211201-002"

    On Error Resume Next	

    url = m_StatementService.GetMassPrintURL(testCorpNum, itemCode, mgtKeyList, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    End If

    On Error GoTo 0 
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�ٷ� �μ� URL ��û</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>URL : <%=CStr(url)%> </li>
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