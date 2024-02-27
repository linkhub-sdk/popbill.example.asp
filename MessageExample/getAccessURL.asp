<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺� ����Ʈ�� �α��� ���·� ������ �� �ִ� �������� �˾� URL�� ��ȯ�մϴ�.
    ' - ��ȯ�Ǵ� URL�� ���� ��å�� 30�� ���� ��ȿ�ϸ�, �ð��� �ʰ��� �Ŀ��� �ش� URL�� ���� ������ ������ �Ұ��մϴ�.
    ' - https://developers.popbill.com/reference/sms/asp/api/member#GetAccessURL
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    url = m_MessageService.GetAccessURL(CorpNum, UserID)

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
