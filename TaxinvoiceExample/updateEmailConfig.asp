<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************'
    ' ���ڼ��ݰ�꼭 ���� �������� �׸� ���� ���ۿ��θ� �����Ѵ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#UpdateEmailConfig
    '
    ' ������������
    ' [������]
    ' TAX_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_CHECK : �����ڿ��� ���ڼ��ݰ�꼭�� ����Ȯ�� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
    '
    ' [������]
    ' TAX_REQUEST : �����ڿ��� ���ݰ�꼭�� ���ڼ��� �Ͽ� ������ ��û�ϴ� �����Դϴ�.
    ' TAX_CANCEL_REQUEST : ���޹޴��ڿ��� ���ݰ�꼭�� ��� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_REFUSE : ���޹޴��ڿ��� ���ݰ�꼭�� �ź� �Ǿ����� �˷��ִ� �����Դϴ�.
    '
    ' [����Ź����]
    ' TAX_TRUST_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_TRUST_ISSUE_TRUSTEE : ��Ź�ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_TRUST_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_TRUST_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' TAX_TRUST_CANCEL_ISSUE_INVOICER : �����ڿ��� ���ڼ��ݰ�꼭�� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
    '
    ' [ó�����]
    ' TAX_CLOSEDOWN : �ŷ�ó�� ����� ���θ� Ȯ���Ͽ� �ȳ��ϴ� �����Դϴ�.
    ' TAX_NTSFAIL_INVOICER : ���ڼ��ݰ�꼭 ����û ���۽��и� �ȳ��ϴ� �����Դϴ�.
    '
    ' [����߼�]
    ' ETC_CERT_EXPIRATION : �˺��� ��ϵ� �������� ���Ό���� �ȳ��ϴ� �����Դϴ�.
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    '�˺�ȸ�� ���̵�
    userID = "testkorea"

    '���� ���� ����
    emailType = "TAX_ISSUE"

    '���� ���� (true = ����, false = ������)
    sendYN = true

    On Error Resume Next

    Set Presponse = m_TaxinvoiceService.updateEmailConfig(testCorpNum, emailType, sendYN, UserID)

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
                <legend>�˸����� ���ۼ��� ����</legend>
                <ul>
                    <li> Response.code : <%=code%> </li>
                    <li> Response.message : <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>