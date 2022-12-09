<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ����ϰ��� �ϴ� ���̵��� �ߺ����θ� Ȯ���մϴ�.
    ' - https://docs.popbill.com/bizinfocheck/asp/api#CheckID
    '**************************************************************

    '�ߺ�Ȯ���� ���̵� 
    testID = "testkorea"    

    On Error Resume Next

    Set Presponse = m_BizInfoCheckService.CheckID(testID)
    
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
                <legend>���̵� �ߺ�Ȯ�� </legend>
                <ul>
                    <li>Response.code : <%=code%> </li>
                    <li>Response.message: <%=message%> </li>
                </ul>
            </fieldset>
         </div>
    </body>
</html>