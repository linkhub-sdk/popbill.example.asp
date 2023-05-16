<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ����û ���ι�ȣ�� ���� ������ ���ڼ��ݰ�꼭 1���� �������� ��ȯ�մϴ�.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#GetTaxinvoice
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' ���ڼ��ݰ�꼭 ����û���ι�ȣ
    NTSConfirmNum = "201611104100020300000cb2"

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.GetTaxinvoice(testCorpNum, NTSConfirmNUm, UserID)

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
                <% If code = 0 Then %>
                <legend>������ ��ȸ</legend>
                <ul>
                    <li>ntsconfirmNum (����û���ι�ȣ) : <%=result.ntsconfirmNum%> </li>
                    <li>writeDate (�ۼ�����) : <%=result.writeDate%> </li>
                    <li>issueDT (�����Ͻ�) : <%=result.issueDT%> </li>
                    <li>invoiceType (���ڼ��ݰ�꼭 ����) : <%=result.invoiceType%> </li>
                    <li>taxType (��������) : <%=result.taxType%> </li>
                    <li>taxTotal (���� �հ�) : <%=result.taxTotal%> </li>
                    <li>supplyCostTotal (���ް��� �հ�) : <%=result.supplyCostTotal%> </li>
                    <li>totalAmount (�հ�ݾ�) : <%=result.totalAmount%> </li>
                    <li>purposeType (����/û��) : <%=result.purposeType%> </li>
                    <li>serialNum (�Ϸù�ȣ) : <%=result.serialNum%> </li>
                    <li>cash (����) : <%=result.cash%> </li>
                    <li>chkBill (��ǥ) : <%=result.chkBill%> </li>
                    <li>credit (�ܻ�) : <%=result.credit%> </li>
                    <li>note (����) : <%=result.note%> </li>
                    <li>remark1 (���1) : <%=result.remark1%> </li>
                    <li>remark2 (���2) : <%=result.remark2%> </li>
                    <li>remark3 (���3) : <%=result.remark3%> </li>

                    <li>modifyCode (���� �����ڵ� ) : <%=result.modifyCode%> </li>
                    <li>orgNTSConfirmNum (���� ���ڼ��ݰ�꼭 ����û���ι�ȣ) : <%=result.orgNTSConfirmNum%> </li>

                    <li>invoicerCorpNum (������ ����ڹ�ȣ) : <%=result.invoicerCorpNum%> </li>
                    <li>invoicerMgtKey (������ ������ȣ) : <%=result.invoicerMgtKey%> </li>
                    <li>invoicerTaxRegID (������ ��������ȣ ) : <%=result.invoicerTaxRegID%> </li>
                    <li>invoicerCorpName (������ ��ȣ) : <%=result.invoicerCorpName%> </li>
                    <li>invoicerCEOName (������ ��ǥ�ڼ���) : <%=result.invoicerCEOName%> </li>
                    <li>invoicerAddr (������ �ּ�) : <%=result.invoicerAddr%> </li>
                    <li>invoicerBizType (������ ����) : <%=result.invoicerBizType%> </li>
                    <li>invoicerBizClass (������ ����) : <%=result.invoicerBizClass%> </li>
                    <li>invoicerContactName (������ ����� ����) : <%=result.invoicerContactName%> </li>
                    <li>invoicerTEL (������ ����ó) : <%=result.invoicerTEL%> </li>
                    <li>invoicerEmail (������ �̸���) : <%=result.invoicerEmail%> </li>

                    <li>invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : <%=result.invoiceeCorpNum%> </li>
                    <li>invoiceeType (���޹޴��� ����) : <%=result.invoiceeType%> </li>
                    <li>invoiceeMgtKey (���޹޴��� ������ȣ) : <%=result.invoiceeMgtKey%> </li>
                    <li>invoiceeTaxRegID (���޹޴��� ��������ȣ) : <%=result.invoiceeTaxRegID%> </li>
                    <li>invoiceeCorpName (���޹޴��� ��ȣ) : <%=result.invoiceeCorpName%> </li>
                    <li>invoiceeCEOName (���޹޴��� ��ǥ�ڼ���) : <%=result.invoiceeCEOName%> </li>
                    <li>invoiceeAddr (���޹޴��� �ּ�) : <%=result.invoiceeAddr%> </li>
                    <li>invoiceeBizType (���޹޴��� ����) : <%=result.invoiceeBizType%> </li>
                    <li>invoiceeBizClass (���޹޴��� ����) : <%=result.invoiceeBizClass%> </li>
                    <li>invoiceeContactName1 (���޹޴��� ����� ����) : <%=result.invoiceeContactName1%> </li>
                    <li>invoiceeTEL1 (���޹޴��� ����� ����ó) : <%=result.invoiceeTEL1%> </li>
                    <li>invoiceeEmail1 (���޹޴��� ����� �̸���) : <%=result.invoiceeEmail1%> </li>
                </ul>
                <fieldset class="fieldset2">
                <%
                    For i=0 To UBound(result.detailList) -1
                %>
                        <legend>ǰ������ [<%=i+1%>]</legend>
                        <ul>
                            <li> serialNum (�Ϸù�ȣ) : <%= result.detailList(i).serialNum %></li>
                            <li> purchaseDT (�ŷ�����) : <%= result.detailList(i).purchaseDT %></li>
                            <li> itemName (ǰ��) : <%= result.detailList(i).itemName %></li>
                            <li> spec (�԰�) : <%= result.detailList(i).spec %></li>
                            <li> qty (����) : <%= result.detailList(i).qty %></li>
                            <li> unitCost (�ܰ�) : <%= result.detailList(i).unitCost %></li>
                            <li> supplyCost (���ް���) : <%= result.detailList(i).supplyCost %></li>
                            <li> tax (����) : <%= result.detailList(i).tax %></li>
                            <li> remark (���) : <%= result.detailList(i).remark %></li>
                        </ul>
                <%
                        Next
                %>
                </fieldset>
                <%	Else  %>
                    <ul>
                        <li>Response.code: <%=code%> </li>
                        <li>Response.message: <%=message%> </li>
                    </ul>
                <%	End If	%>
            </fieldset>
         </div>
    </body>
</html>