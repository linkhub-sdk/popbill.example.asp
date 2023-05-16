<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭�� ���¿� ���� �����̷��� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetLogs
    '**************************************************************

    '  �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetLogs(testCorpNum, KeyType, MgtKey)

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
                <legend> �����̷�Ȯ�� </legend>
                <%
                    If code = 0 Then
                        For i=0 To result.Count -1 %>
                         <fieldset class="fieldset2">
                            <ul>
                                <li> DocLogType(�α�Ÿ��) :  <%=result.Item(i).DocLogType%> </li>
                                <li> Log(�̷�����) : <%=result.Item(i).Log %> </li>
                                <li> ProcType(ó������) : <%=result.Item(i).ProcType%> </li>
                                <li> ProcCorpName(ó��ȸ���) : <%=result.Item(i).ProcCorpName%></li>
                                <li> procContactName(ó�������) : <%=result.Item(i).procContactName%></li>
                                <li> ProcMemo(ó���޸�) : <%=result.Item(i).ProcMemo %></li>
                                <li> regDT(����Ͻ�) : <%=result.Item(i).regDT %></li>
                                <li> ip(������) : <%=result.Item(i).ip %></li>
                            </ul>
                        </fieldset>
                <%
                    Next
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