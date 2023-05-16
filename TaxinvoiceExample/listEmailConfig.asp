<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭 ���� ���� �׸� ���� �߼ۼ����� Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/etc#ListEmailConfig
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    On Error Resume Next

    Set emailObj = m_TaxinvoiceService.listEmailConfig(testCorpNum, UserID)

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
                <legend>�˸����� ���۸�� ��ȸ</legend>
                        <ul>
                        <%
                            If code = 0 Then
                            For i=0 To emailObj.Count-1
                        %>
                            <% If emailObj.Item(i).emailType = "TAX_ISSUE" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(���޹޴��ڿ��� ���ڼ��ݰ�꼭 ���� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_ISSUE_INVOICER" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(�����ڿ��� ���ڼ��ݰ�꼭 ���� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CHECK" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(�����ڿ��� ���ڼ��ݰ�꼭 ����Ȯ�� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CANCEL_ISSUE" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_REQUEST" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(�����ڿ��� ���ݰ�꼭�� �����û ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CANCEL_REQUEST" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(���޹޴��ڿ��� ���ݰ�꼭 ��� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_REFUSE" Then %>
                                    <li>[������] <%= emailObj.Item(i).emailType %>(���޹޴��ڿ��� ���ݰ�꼭 �ź� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE" Then %>
                                    <li>[����Ź����] <%= emailObj.Item(i).emailType %>(��Ź�ڿ��� ���ڼ��ݰ�꼭 ���� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE_TRUSTEE" Then %>
                                    <li>[����Ź����] <%= emailObj.Item(i).emailType %>(��Ź�ڿ��� ���ڼ��ݰ�꼭 ���� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_ISSUE_INVOICER" Then %>
                                    <li>[����Ź����] <%= emailObj.Item(i).emailType %>(�����ڿ��� ���ڼ��ݰ�꼭 ���� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_CANCEL_ISSUE" Then %>
                                    <li>[����Ź����] <%= emailObj.Item(i).emailType %>(���޹޴��ڿ��� ���ڼ��ݰ�꼭 ������� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_TRUST_CANCEL_ISSUE_INVOICER" Then %>
                                    <li>[����Ź����] <%= emailObj.Item(i).emailType %>(�����ڿ��� ���ڼ��ݰ�꼭 ������� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_CLOSEDOWN" Then %>
                                    <li>[ó����� ]<%= emailObj.Item(i).emailType %>(�ŷ�ó�� ����� ���� Ȯ�� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "TAX_NTSFAIL_INVOICER" Then %>
                                    <li>[ó�����] <%= emailObj.Item(i).emailType %>(���ڼ��ݰ�꼭 ����û ���۽��� �ȳ� ���� ���� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                            <% If emailObj.Item(i).emailType = "ETC_CERT_EXPIRATION" Then %>
                                    <li>[ó�����] <%= emailObj.Item(i).emailType %>(�˺��� ��ϵ� �������� ���Ό���� �ȳ��ϴ� ����) : <%= emailObj.Item(i).sendYN %></li>
                            <% End If %>
                        <%
                            Next
                            Else
                        %>
                        </ul>
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
