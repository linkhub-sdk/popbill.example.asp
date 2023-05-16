<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' "�ӽ�����" �Ǵ� "(��)������" ������ ���ݰ�꼭�� ����(���ڼ���)�ϸ�, "����Ϸ�" ���·� ó���մϴ�.
    ' - ���ݰ�꼭 ����û ������å [https://developers.popbill.com/guide/taxinvoice/asp/introduction/policy-of-send-to-nts]
    ' - "����Ϸ�" �� ���ڼ��ݰ�꼭�� ����û ���� ������ �������(CancelIssue API) �Լ��� ����û �Ű� ��󿡼� ������ �� �ֽ��ϴ�.
    ' - ���ݰ�꼭 ������ ���ؼ� �������� �������� �˺� ���������� ������� �Ǿ�� �մϴ�.
    '   �� ����Ź������ ���, ��Ź���� ������ ����� �ʿ��մϴ�.
    ' - ���ݰ�꼭 ���� �� ���޹޴��ڿ��� ���� ������ �߼۵˴ϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/issue#Issue
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    testUserID = "testkorea"

    ' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-002"

    ' �޸�
    Memo = "���� �޸�"

    ' ���� �ȳ����� ����, �̱���� �⺻������� ����
    EmailSubject = ""

    ' �������� ��������  (true / false �� �� 1)
    ' �� true = ���� , false = �Ұ���
    ' - ���ึ������ ���� ���ݰ�꼭�� �����ϴ� ���, ���꼼�� �ΰ��� �� �ֽ��ϴ�.
    ' - ���꼼�� �ΰ��Ǵ��� ������ �ؾ��ϴ� ��쿡�� forceIssue�� ����
    '   true�� �����Ͽ� ����(Issue API)�� ȣ���Ͻø� �˴ϴ�.
    ForceIssue = False

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.Issue(testCorpNum, KeyType ,MgtKey, Memo ,EmailSubject, ForceIssue, testUserID)

    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        ntsConfirmNum = ""
        Err.Clears
    Else
        code = Presponse.code
        message = Presponse.message
        ntsConfirmNum = Presponse.ntsConfirmNum
    End If

    On Error GoTo 0
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>���ݰ�꼭 ����</legend>
                <ul>
                    <li>�����ڵ� (Response.code) : <%=code%> </li>
                    <li>����޽��� (Response.message) : <%=message%> </li>
                    <% If ntsConfirmNum <> "" Then %>
                    <li>����û���ι�ȣ (Response.ntsConfirmNum) : <%=ntsConfirmNum%> </li>
                    <% End If %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>