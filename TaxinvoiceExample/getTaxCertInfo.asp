<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˺� ���������� ��ϵ� ������������ ������ Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/cert#GetTaxCertInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    On Error Resume Next

    Set resultObj = m_TaxinvoiceService.GetTaxCertInfo(testCorpNum)

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
                <legend>������ ���� Ȯ��</legend>
                <%

                    If code = 0 Then
                %>
                <ul>
                    <li>regDT (����Ͻ�) : <%=resultObj.regDT %></li>
                    <li>expireDT (�����Ͻ�) : <%=resultObj.expireDT %></li>
                    <li>issuerDN (������ �߱��� DN) : <%=resultObj.issuerDN %></li>
                    <li>subjectDN (��ϵ� ������ DN) : <%=resultObj.subjectDN %></li>
                    <li>issuerName (������ ����) : <%=resultObj.issuerName %></li>
                    <li>oid (OID) : <%=resultObj.oid %></li>
                    <li>regContactName (��� ����� ����) : <%=resultObj.regContactName %></li>
                    <li>regContactID (��� ����� ���̵�) : <%=resultObj.regContactID %></li>
                </ul>

                <%
                    Else
                %>
                    <ul>
                        <li>Response.code : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>