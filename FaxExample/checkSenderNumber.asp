<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ѽ� �߽Ź�ȣ ��Ͽ��θ� Ȯ���մϴ�.
    ' - �߽Ź�ȣ ���°� '����'�� ��쿡�� ���ϰ� 'Response'�� ���� 'code'�� 1�� ��ȯ�˴ϴ�.
    ' - https://developers.popbill.com/reference/fax/asp/api/sendnum#CheckSenderNumber
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' Ȯ���� �߽Ź�ȣ
    SenderNumber = ""

    On Error Resume Next

    Set result = m_FaxService.CheckSenderNumber(testCorpNum, SenderNumber, userID)

    If Err.Number <> 0 then
        code = Err.Number
        message =  Err.Description
        Err.Clears
    Else
        code = result.code
        message = result.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�߽Ź�ȣ ��Ͽ��� Ȯ��</legend>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
         </div>
    </body>
</html>