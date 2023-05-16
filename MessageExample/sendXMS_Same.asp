<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �޽��� ũ��(90byte)�� ���� �ܹ�/�幮(SMS/LMS)�� �ڵ����� �ν��Ͽ� 1���� �޽����� ������ �˺��� �����ϸ�, ��� �����ڿ��� ���� ������ �����մϴ�. (�ִ� 1,000��)
    ' - �ܹ�(SMS) = 90byte ������ �޽���, �幮(LMS) = 2000byte ������ �޽���.
    ' - https://developers.popbill.com/reference/sms/asp/api/send#SendXMS
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �޽��� ���� ( true , false �� �� 1)
    ' �� true = ���� , false = �Ϲ�
    adsYN = False

    ' �������۽ð� yyyyMMddHHmmss, reserveDT���� ���� ��� �������
    reserveDT = ""

    ' �߽Ź�ȣ
    senderNum = ""

    ' �޽��� ����
    subject = "�������� ����"

    ' �޽��� ����, 90byte �������� ��/�幮 �޽����� �ڵ��νĵǾ� ����
    content = "�������� ����, 90Byte�ʰ��� LMS(�幮) �޽����� ���۵�"


    ' �������� �迭, �ִ� 1000��
    Set msgList = CreateObject("Scripting.Dictionary")

    For i =0 To 49
        Set message = New Messages

        ' ���Ź�ȣ
        message.receiver = "000111222"

        ' �����ڸ�
        message.receivername = " �������̸�"+CStr(i)

        msgList.Add i, message
    Next

    For i =50 To 99
        Set message = New Messages

        ' ���Ź�ȣ
        message.receiver = "000111222"

        ' �����ڸ�
        message.receivername = " �������̸�"+CStr(i)

        msgList.Add i, message
    Next

    ' ���ۿ�û��ȣ
    ' �˺��� ���� ������ �ĺ��� �� �ֵ��� ��Ʈ�ʰ� �Ҵ��� �ĺ���ȣ.
    ' 1~36�ڸ��� ����. ����, ����, ������(-), �����(_)�� �����Ͽ� �˺� ȸ������ �ߺ����� �ʵ��� �Ҵ�.
    requestNum = ""

    On Error Resume Next

    receiptNum = m_MessageService.SendXMS(testCorpNum, senderNum, subject,content, msgList, reserveDT, adsYN, requestNum, userID)

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
                <legend>��/�幮 �ڵ��ν� ���ڸ޽��� �������� </legend>
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