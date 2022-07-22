<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************'
    ' ���ڸ��� ���� �������� �׸� ���� ���ۿ��θ� �����Ѵ�.
    ' - https://docs.popbill.com/statement/asp/api#UpdateEmailConfig
    '
    ' ������������
    ' SMT_ISSUE : ���޹޴��ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' SMT_ACCEPT : �����ڿ��� ���ڸ����� ���� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' SMT_DENY : �����ڿ��� ���ڸ����� �ź� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' SMT_CANCEL : ���޹޴��ڿ��� ���ڸ����� ��� �Ǿ����� �˷��ִ� �����Դϴ�.
    ' SMT_CANCEL_ISSUE : ���޹޴��ڿ��� ���ڸ����� ������� �Ǿ����� �˷��ִ� �����Դϴ�.
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"	 

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"		 

    ' ���� ���� ����
    emailType = "SMT_ISSUE"

    ' ���� ���� (true = ����, false = ������)
    sendYN = true

    On Error Resume Next

    Set Presponse = m_StatementService.updateEmailConfig(testCorpNum, emailType, sendYN, UserID)

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