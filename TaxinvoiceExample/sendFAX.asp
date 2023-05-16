<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭�� �ѽ��� �����ϴ� �Լ���, �˺� ����Ʈ [���ڡ��ѽ�] > [�ѽ�] > [���۳���] �޴����� ���۰���� Ȯ�� �� �� �ֽ��ϴ�.
    ' - �Լ� ȣ��� ����Ʈ�� ���ݵ˴ϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#SendFAX
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    ' �߽Ź�ȣ
    Sender = ""

    ' �����ѽ���ȣ
    Receiver = ""

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.SendFAX(testCorpNum, KeyType, MgtKey, Sender, Receiver, testUserID)

    If Err.Number <> 0 Then
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
                <legend>�ѽ� ������</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message : <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>