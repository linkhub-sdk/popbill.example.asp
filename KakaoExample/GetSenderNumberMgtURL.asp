<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �߽Ź�ȣ ���� �˾� URL�� ��ȯ�մϴ�.
    ' - ������å�� ���� ��ȯ�� URL�� 30���� ��ȿ�ð��� �����ϴ�.
    ' - https://docs.popbill.com/kakao/asp/api#GetSenderNumberMgtURL
    '**************************************************************

    '�˺� ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺� ȸ�� ���̵�
    userID = "testkorea"

    On Error Resume Next

    url = m_KakaoService.GetSenderNumberMgtURL(testCorpNum, userID)

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
                <legend>�߽Ź�ȣ ���� �˾� URL</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li>URL : <%=url%> </li>
                    <% Else %>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>