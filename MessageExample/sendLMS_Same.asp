<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �ִ� 2,000byte�� �幮(LMS) �޽��� �ټ��� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
    '  - https://developers.popbill.com/reference/sms/asp/api/send#SendLMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' �߽Ź�ȣ
    senderNum = ""

    ' �޽��� ����
    subject = "�������� �޽��� ����"

    ' �޽��� ����, �ִ� 2000byte �ʰ��� ���̰� �����Ǿ� ���۵�
    content = "�������� �޽��� ����"

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    ' ������ �޽��� ���� ( true , false �� �� 1)
    ' �� true = ���� , false = �Ϲ�
    adsYN = False

    ' ���������迭, �ִ� 1000��
    Set msgList = CreateObject("Scripting.Dictionary")

    For i =0 To 15
        Set message = New Messages

        ' ���Ź�ȣ
        message.receiver = ""

        ' �����ڸ�
        message.receivername = " �������̸�"+CStr(i)

        msgList.Add i, message
    Next

    On Error Resume Next

    ' ���ۿ�û��ȣ
    ' �˺��� ���� ������ �ĺ��� �� �ֵ��� ��Ʈ�ʰ� �Ҵ��� �ĺ���ȣ.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    requestNum = ""

    receiptNum = m_MessageService.SendLMS(testCorpNum, senderNum, subject, content, msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>�幮 ���ڸ޽��� �������� </legend>
                <% If code = 0 Then %>
                    <ul>
                        <li>ReceiptNum(������ȣ) : <%=receiptNum%> </li>
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