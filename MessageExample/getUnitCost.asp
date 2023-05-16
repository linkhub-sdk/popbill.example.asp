<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���� ���۽� ���ݵǴ� ����Ʈ �ܰ��� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/sms/asp/api/point#GetUnitCost
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' �������� (SMS - �ܹ�, LMS - �幮, MMS - ����)
    sendType = "SMS"

    On Error Resume Next

    unitCost = m_MessageService.GetUnitCost(testCorpNum, sendType)

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
                <legend>���ڸ޽��� ���۴ܰ� Ȯ��</legend>
                <% If code = 0 Then %>
                    <ul>
                        <li><%=sendType%> ���۴ܰ� : <%=CInt(unitCost)%> </li>
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