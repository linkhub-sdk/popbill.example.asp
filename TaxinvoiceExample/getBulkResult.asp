<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' ������ ������ SubmitID�� ����Ͽ� ���ݰ�꼭 ��������� Ȯ���մϴ�.
    ' - ���� ���ݰ�꼭 ó�����´� ��������(txState)�� �Ϸ�(2) �� ��ȯ�˴ϴ�.
    ' - https://docs.popbill.com/taxinvoice/asp/api#GetBulkResult
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' ������̵�, �ִ� 36�ڸ� (����, ����, "-" ����)
    SubmitID = "20220720-ASP-BULK001"

    ' �˺�ȸ�����̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetBulkResult(testCorpNum, SubmitID, UserID)

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
                <legend>�ʴ뷮 ���� ��� Ȯ��</legend>
                <% 
                    If code = 0 Then 
                %>
                    <ul>
                        <li> code (�����ڵ�) :  <%=result.code%> </li>
                        <li> message (����޽���) :  <%=result.message%> </li>
                        <li> submitID (������̵�) :  <%=result.submitID%> </li>
                        <li> submitCount (���ݰ�꼭 ���� �Ǽ�) :  <%=result.submitCount%> </li>
                        <li> successCount (���ݰ�꼭 ���� ���� �Ǽ�) : <%=result.successCount%></li> 
                        <li> failCount (���ݰ�꼭 ���� ���� �Ǽ�) :  <%=result.failCount %> </li>
                        <li> txState (���������ڵ�) :  <%=result.txState%> </li>
                        <li> txResultCode (���� ����ڵ�) :  <%=result.txResultCode%> </li>
                        <li> txStartDT (����ó�� �����Ͻ�) :  <%=result.txStartDT%> </li>
                        <li> txEndDT (����ó�� �Ϸ��Ͻ�	) :  <%=result.txEndDT%> </li>
                        <li> receiptDT (�����Ͻ�) :  <%=result.receiptDT%> </li>
                        <li> receiptID (�������̵�) :  <%=result.receiptID%> </li>
                    </ul>
                    <%   Dim i
                        For i=0 To UBound(result.issueResult) -1
                     %>	
                     <fieldset class="fieldset2">					
                        <legend>  issueResult (���� ���) [ <%=i+1%> / <%=UBound(result.issueResult)%> ]</legend>
                        <ul>
                            <li> invoicerMgtKey (������ ������ȣ) : <%=result.issueResult(i).invoicerMgtKey %>
                            <li> trusteeMgtKey (��Ź�� ������ȣ	) : <%=result.issueResult(i).trusteeMgtKey %>
                            <li> code (�����ڵ�) : <%=result.issueResult(i).code %>
                            <li> message (����޽���) : <%=result.issueResult(i).message %>
                            <li> ntsconfirmNum (����û���ι�ȣ) : <%=result.issueResult(i).ntsconfirmNum %>
                            <li> issueDT (�����Ͻ�) : <%=result.issueResult(i).issueDT %>
                        </ul>
                    </fieldset>
                     <% Next %>
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