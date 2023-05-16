<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ���� �ܿ�����Ʈ�� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/closedown/asp/api/point#GetBalance
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    On Error Resume Next

    remainPoint = m_ClosedownService.getBalance(testCorpNum)

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
                <legend>����ȸ�� �ܿ�����Ʈ Ȯ��</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>�ܿ�����Ʈ : <%=CStr(remainpoint)%> </li>
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