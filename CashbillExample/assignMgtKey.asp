<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' �˺� ����Ʈ�� ���� �����Ͽ� ������ȣ�� �ο����� ���� ���ݿ������� ������ȣ�� �Ҵ��մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/etc#AssignMgtKey
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"		 
    
    ' ���ݿ����� ������Ű, ���� �����ȸ(Search) API�� ��ȯ�׸��� ItemKey ����
    itemKey = "020010910360800001"

    ' �Ҵ��� ������ȣ, ����, ���� '-', '_' �������� �ִ� 24�ڸ�����
    ' ����ڹ�ȣ�� �ߺ����� ������ȣ �Ҵ�
    mgtKey = "20220720-ASP-004"

    On Error Resume Next

    Set Presponse = m_CashbillService.AssignMgtKey(testCorpNum, itemKey, mgtKey)
    
    If Err.Number <> 0 Then
        code = Err.Number
        message = Err.Description
        Err.Clears
    Else
        code = Presponse.code
        message =Presponse.message
    End If

    On Error GoTo 0 
%>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>������ȣ �Ҵ� </legend>
                <ul>
                    <li> Response.code : <%=code%></li>
                    <li> Response.message : <%=message%></li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>