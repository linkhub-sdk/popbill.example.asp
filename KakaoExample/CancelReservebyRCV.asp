<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺����� ��ȯ���� ������ȣ�� ���� ���� �ĺ��Ͽ� ���Ź�ȣ�� ����� īī������ ���� ����մϴ�. (����ð� 10�� ������ ����)
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#CancelReservebyRCV
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    'īī���� �������� ������ �˺��κ��� ��ȯ ���� ������ȣ
    ReceiptNum = "018031513173900001"

    'īī���� �������� ������ �˺��� ��û�� ���Ź�ȣ
    ReceiveNum = "010111222"

    On Error Resume Next

    Set result = m_KakaoService.CancelReservebyRCV(CorpNum, ReceiptNum, ReceiveNum, UserID)

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
                <legend>�������� �Ϻ� ��� (������ȣ)</legend>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
        </div>
    </body>
</html>
