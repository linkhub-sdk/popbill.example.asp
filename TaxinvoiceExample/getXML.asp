<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭 1���� �������� XML�� ��ȯ�մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetXML
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    ' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType = "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set taxXML = m_TaxinvoiceService.GetXML(testCorpNum, KeyType, MgtKey)

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
                <legend>������ Ȯ�� - XML</legend>
                <%

                    If code = 0 Then
                %>
                <ul>
                    <li>code (�����ڵ�) : <%=taxXML.code%></li>
                    <li>message (����޽���) : <%=taxXML.message%></li>
                    <li>retObject (���ڼ��ݰ�꼭 XML����) : <%=Replace(taxXML.retObject, "<", "&lt;")%></li>
                    <!-- Browser���� xml������ ����ϱ� ���� '<' &lt�� ġȯ�Ͽ����ϴ�. -->
                </ul>

                <%
                    Else
                %>
                    <ul>
                        <li>Response.dcode : <%=code%> </li>
                        <li>Response.message : <%=message%> </li>
                    </ul>
                <%
                    End If
                %>
            </fieldset>
         </div>
    </body>
</html>