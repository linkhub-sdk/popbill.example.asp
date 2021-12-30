<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �˺��� �α��� ���·� ������ �� �ִ� �˾� URL�� ��ȯ�մϴ�.
    ' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
    ' - https://docs.popbill.com/easyfinbank/asp/api#GetAccessURL
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	 
    
    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    On Error Resume Next
    url = m_EasyFinBAnkService.GetAccessURL(testCorpNum, userID)
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
                <legend>�˺� �α��� �˾� URL</legend>
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