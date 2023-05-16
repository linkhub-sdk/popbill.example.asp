<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' Ȩ�ý����� ������ ���� �˺��� ���ڼ��ݰ�꼭�� �μ������ ������ ����մϴ�.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/cert#RegistDeptUser
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' Ȩ�ý����� ������ ���ڼ��ݰ�꼭 �μ������ ���̵�
    deptUserID = "userid_test"

    ' Ȩ�ý����� ������ ���ڼ��ݰ�꼭 �μ������ ��й�ȣ
    deptUserPWD = "passwd_test"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    On Error Resume Next

    Set Presponse = m_HTTaxinvoiceService.RegistDeptUser(testCorpNum, deptUserID, deptUserPWD, userID)

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
                <legend>�μ������ �������</legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>