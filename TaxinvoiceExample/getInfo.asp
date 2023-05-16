<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݰ�꼭 1���� ���� �� ��������� Ȯ���մϴ�.
    ' - ���ϰ� 'TaxinvoiceInfo'�� ���� 'stateCode'�� ���� ���ݰ�꼭�� �����ڵ带 Ȯ���մϴ�.
    ' - ���ݰ�꼭 �����ڵ� [https://developers.popbill.com/reference/taxinvoice/asp/response-code#state-code]
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#GetInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType= "SELL"

    ' ������ȣ
    MgtKey = "20220720-ASP-001"

    ' �˺�ȸ�����̵�
    UserID = "testkorea"

    On Error Resume Next

    Set result = m_TaxinvoiceService.GetInfo(testCorpNum, KeyType, MgtKey, UserID)

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
                <legend>���ݰ�꼭 ����/��� ���� Ȯ�� </legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> itemKey (���ݰ�꼭 ������Ű) :  <%=result.itemKey%> </li>
                        <li> taxType (��������) :  <%=result.taxType%> </li>
                        <li> writeDate (�ۼ�����) :  <%=result.writeDate%> </li>
                        <li> regDT (�ӽ����� ����) :  <%=result.regDT%> </li>
                        <li> issueType (��������) :  <%=result.issueType %> </li>
                        <li> supplyCostTotal (���ް��� �հ�) :  <%=result.supplyCostTotal%> </li>
                        <li> taxTotal (���� �հ�) :  <%=result.taxTotal%> </li>
                        <li> purposeType (����/û��) :  <%=result.purposeType%> </li>
                        <li> issueDT (�����Ͻ�) :  <%=result.issueDT%> </li>
                        <li> lateIssueYN (�������� ����) :  <%=result.lateIssueYN%> </li>
                        <li> preIssueDT (���࿹���Ͻ�) :  <%=result.preIssueDT%> </li>
                        <li> openYN (���� ����) :  <%=result.openYN%> </li>
                        <li> openDT (�����Ͻ�) :  <%=result.openDT%> </li>
                        <li> stateMemo (���¸޸�) :  <%=result.stateMemo%> </li>
                        <li> stateCode (�����ڵ�) :  <%=result.stateCode%> </li>
                        <li> stateDT (���� �����Ͻ�) :  <%=result.stateDT%> </li>
                        <li> ntsconfirmNum (����û ���ι�ȣ) :  <%=result.ntsconfirmNum %> </li>
                        <li> ntsresult (����û ���۰��) :  <%=result.ntsresult%> </li>
                        <li> ntssendDT (����û �����Ͻ�) :  <%=result.ntssendDT%> </li>
                        <li> ntsresultDT  (����û ��� �����Ͻ�) :  <%=result.ntsresultDT%> </li>
                        <li> ntssendErrCode (���۽��� �����ڵ�) :  <%=result.ntssendErrCode%> </li>
                        <li> modifyCode (���������ڵ�) : <%=result.modifyCode%></li>
                        <li> interOPYN (������������) :  <%=result.interOPYN%> </li>
                        <li> invoicerCorpName (������ ��ȣ) :  <%=result.invoicerCorpName%> </li>
                        <li> invoicerCorpNum (������ ����ڹ�ȣ) :  <%=result.invoicerCorpNum%> </li>
                        <li> invoicerMgtKey (������ ������ȣ) :  <%=result.invoicerMgtKey%> </li>
                        <li> invoicerPrintYN (������ �μ⿩��) :  <%=result.invoicerPrintYN%> </li>
                        <li> invoiceeCorpName (���޹޴��� ��ȣ) :  <%=result.invoiceeCorpName%> </li>
                        <li> invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) :  <%=result.invoiceeCorpNum%> </li>
                        <li> invoiceeMgtKey (���޹޴��� ������ȣ) :  <%=result.invoiceeMgtKey%> </li>
                        <li> invoiceePrintYN (���޹޴��� �μ⿩��) :  <%=result.invoiceePrintYN%> </li>
                        <li> closeDownState (���޹޴��� ���������) :  <%=result.closeDownState%> </li>
                        <li> closeDownStateDate (���޹޴��� ���������) :  <%=result.closeDownStateDate%> </li>

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