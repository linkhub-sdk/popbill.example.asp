<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����ȸ�� ����Ʈ�� ȯ�� ��û�մϴ�.
    ' - https://developers.popbill.com/reference/easyfinbank/asp/api/point#Refund
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    CorpNum = "1234567890"

    Dim m_RefundForm : Set m_RefundForm = New RefundForm
    '����ڸ�
    m_RefundForm.ContactName = "����� �̸�"

    '����� ����ó
    m_RefundForm.TEL = "010-1234-1234"

    'ȯ�� ��û ����Ʈ
    m_RefundForm.RequestPoint = "1000"

    '�����
    m_RefundForm.AccountBank = "����"

    '���¹�ȣ
    m_RefundForm.AccountNum = "110-1234-12345"

    '�����ָ�
    m_RefundForm.AccountName = "������_�׽�Ʈ"

    'ȯ�һ���
    m_RefundForm.Reason = "ȯ���ϰڽ��ϴ�"


    '�˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_EasyFinBankService.Refund(CorpNum, m_RefundForm, UserID)

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
                <legend>����ȸ�� ����Ʈ ȯ�ҽ�û</legend>
                <%
                    If code = 0 Then
                %>
                    <fieldset class="fieldset2">
                        <legend> result </legend>
                            <ul>
                                <li> code (���� �ڵ�) : <%=result.code%></li>
                                <li> message (���� �޽���) : <%=result.message%></li>
                                <li> refundCode (ȯ���ڵ�) : <%=result.refundCode%></li>
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
