<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ�� ����Ʈ ������ ���� �������Ա��� ��û�մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/point#PaymentRequest
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    Dim m_PaymentForm : Set m_PaymentForm = New PaymentForm

    '����ڸ�
    m_PaymentForm.SettlerName = "�����"

    '����� �̸���
    m_PaymentForm.SettlerEmail = "email_damdang@email.com"


    '����� �޴���
    m_PaymentForm.NotifyHP = "010-1234-1234"

    '�Ա��ڸ�
    m_PaymentForm.PaymentName = "�Ա���"

    '�����ݾ�
    m_PaymentForm.SettleCost = "10000"

    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_CashbillService.PaymentRequest(CorpNum, m_PaymentForm, UserID)

    If Err.Number <> 0 Then
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
                <legend>����ȸ�� ������ �Աݽ�û</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> PaymentResponse </legend>
                            <ul>
                                <li> code (�����ڵ�) : <%=result.code%></li>
                                <li> message (����޽���) : <%=result.message%></li>
                                <li> settleCode (�����ڵ�) : <%=result.settleCode%></li>
                            </ul>
                        </fieldset>
                <%
                    Else
                %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
        </div>
    </body>
</html>
