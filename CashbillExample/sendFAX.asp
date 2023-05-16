<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݿ������� �ѽ��� �����ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
    ' - �ѽ� ���� ��û�� ����Ʈ�� �����˴ϴ�. (���۽��н� ȯ��ó��)
    ' - https://developers.popbill.com/reference/cashbill/asp/api/etc#SendFAX
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺�ȸ�� ���̵�
    userID = "testkorea"

    '������ȣ
    mgtKey = "20220720-ASP-001"

    '�߽Ź�ȣ
    sender = ""

    '�����ѽ���ȣ
    receiver = ""

    On Error Resume Next

    Set Presponse = m_CashbillService.SendFAX(testCorpNum, mgtKey, Sender, Receiver, UserID)

    If Err.Number <> 0 then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݿ����� �ѽ����� </legend>
                <ul>
                    <li>Response.code : <%=code%></li>
                    <li>Response.message : <%=message%></li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>