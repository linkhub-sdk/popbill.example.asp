<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ��Ʈ�ʰ� �Ҵ��� ���� ��û��ȣ�� ���� ���� �ĺ��Ͽ� ���Ź�ȣ�� ����� īī������ ���� ����մϴ�. (����ð� 10�� ������ ����)
    ' - https://developers.popbill.com/reference/kakaotalk/asp/api/send#CancelReserveRNbyRCV
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    'īī���� �������� ������ ��Ʈ�ʰ� �Ҵ��� ���� ��û��ȣ
    RequestNum = "20221221123456"

    'īī���� �������� ������ �˺��� ��û�� ���Ź�ȣ
    ReceiveNum = "010222333"

    On Error Resume Next

    Set result = m_KakaoService.CancelReserveRNbyRCV(CorpNum, RequestNum,ReceiveNum, UserID)

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
                <legend>�������� �Ϻ� ��� (���� ��û��ȣ)</legend>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
            </fieldset>
        </div>
    </body>
</html>
