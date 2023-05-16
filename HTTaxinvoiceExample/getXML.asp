<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����û ���ι�ȣ�� ���� ������ ���ڼ��ݰ�꼭 1���� �������� XML ������ ���ڿ��� ��ȯ�մϴ�.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#GetXML
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' ����û���ι�ȣ
    NTSConfirmNum = "201611104100020300000cb2"

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.GetXML ( testCorpNum, NTSConfirmNum, UserID )

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
                <legend> ������ ��ȸ - XML</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> ResultCode (��û�� ���� ���� �����ڵ�) : <%=result.ResultCode%></li>
                        <li> Message (����û���ι�ȣ) : <%=result.Message%></li>
                        <li> retObject (���ڼ��ݰ�꼭 XML ����) : <%=Replace(result.retObject, "<" ,"&lt")%></li>
                    </ul>
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
