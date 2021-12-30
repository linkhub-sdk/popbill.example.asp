<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"--> 
<%
    '**************************************************************
    ' 1���� ���ݰ�꼭 ���׸��� Ȯ���մϴ�.
    ' - https://docs.popbill.com/taxinvoice/asp/api#GetDetailInfo
    '**************************************************************
    
    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    ' ���ݰ�꼭 �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType = "SELL"

    ' ������ȣ
    MgtKey = "20211201-002"

    On Error Resume Next

    Set taxInfo = m_TaxinvoiceService.GetDetailInfo(testCorpNum, KeyType, MgtKey)

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
                <legend>���ݰ�꼭 ������ Ȯ�� </legend>
                <% 

                    If code = 0 Then
                %>
                <ul>
                    <li>ntsconfirmNum (����û ���ι�ȣ) : <%=taxInfo.ntsconfirmNum%></li>
                    <li>writeDate (�ۼ�����) : <%=taxInfo.writeDate%></li>
                    <li>chargeDirection (���ݹ���) : <%=taxInfo.chargeDirection%></li>
                    <li>issueType (��������) : <%=taxInfo.issueType%></li>
                    <li>taxType (��������) : <%=taxInfo.taxType%></li>
                    <li>supplyCostTotal (���ް��� �հ�) : <%=taxInfo.supplyCostTotal%></li>
                    <li>taxTotal (���� �հ�) : <%=taxInfo.taxTotal%></li>
                    <li>totalAmount (�հ�ݾ�) : <%=taxInfo.totalAmount%></li>
                    <li>cash (����) : <%=taxInfo.cash%></li>
                    <li>chkBill (��ǥ) : <%=taxInfo.chkBill%></li>
                    <li>credit (�ܻ�) : <%=taxInfo.credit%></li>
                    <li>note (����) : <%=taxInfo.note%></li>
                    <li>remark1 (���1) : <%=taxInfo.remark1%></li>
                    <li>remark2 (���2) : <%=taxInfo.remark2%></li>
                    <li>remark3 (���3) : <%=taxInfo.remark3%></li>

                    <li>invoicerCorpNum (������ ����ڹ�ȣ) : <%=taxInfo.invoicerCorpNum%> </li>
                    <li>invoicerMgtKey (������ ������ȣ) : <%=taxInfo.invoicerMgtKey%></li>
                    <li>invoicerTaxRegID (������ ������� �ĺ���ȣ) : <%=taxInfo.invoicerTaxRegID%></li>
                    <li>invoicerCorpName (������ ��ȣ) : <%=taxInfo.invoicerCorpName%></li>
                    <li>invoicerCEOName (������ ��ǥ�ڸ�) : <%=taxInfo.invoicerCEOName%></li>
                    <li>invoicerAddr (������ �ּ�) : <%=taxInfo.invoicerAddr%></li>
                    <li>invoicerBizClass (������ ����) : <%=taxInfo.invoicerBizClass%></li>
                    <li>invoicerBizType (������ ����) : <%=taxInfo.invoicerBizType%></li>
                    <li>invoicerContactName (������ ����ڸ�) : <%=taxInfo.invoicerContactName%></li>
                    <li>invoicerTEL (������ ����ó) : <%=taxInfo.invoicerTEL%></li>
                    <li>invoicerHP (������ �޴�����ȣ) : <%=taxInfo.invoicerHP%></li>
                    <li>invoicerEmail (������ ����) : <%=taxInfo.invoicerEmail%></li>
                    <li>invoicerSMSSendYN (�˸����� ���ۿ���) : <%=taxInfo.invoicerSMSSendYN%></li>

                    <li>invoiceeType (���޹޴��� ����) : <%=taxInfo.invoiceeType%></li>
                    <li>invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) : <%=taxInfo.invoiceeCorpNum%></li>
                    <li>invoiceeMgtKey (���޹޴��� ������ȣ) : <%=taxInfo.invoiceeMgtKey%></li>
                    <li>invoiceeTaxRegID (���޹޴��� ������� �ĺ���ȣ) : <%=taxInfo.invoiceeTaxRegID%></li>
                    <li>invoiceeCorpName (���޹޴��� ��ȣ) : <%=taxInfo.invoiceeCorpName%></li>
                    <li>invoiceeCEOName (���޹޴��� ��ǥ�ڸ�) : <%=taxInfo.invoiceeCEOName%></li>
                    <li>invoiceeAddr (���޹޴��� �ּ�) : <%=taxInfo.invoiceeAddr%></li>
                    <li>invoiceeBizClass (���޹޴��� ����) : <%=taxInfo.invoiceeBizClass%></li>
                    <li>invoiceeBizType (���޹޴��� ����) : <%=taxInfo.invoiceeBizType%></li>
                    <li>invoiceeContactName1 (���޹޴��� ����ڸ�) : <%=taxInfo.invoiceeContactName1%></li>
                    <li>closeDownState (���޹޴��� ���������) : <%=taxInfo.closeDownState%></li>
                    <li>closeDownStateDate (���޹޴��� ���������) : <%=taxInfo.closeDownStateDate%></li>

                    <%
                        For i=0 To UBound(taxInfo.detailList)-1
                    %>
                        <fieldset class="fieldset2">
                        <legend>���׸�(ǰ��) ���� <%=i+1%> </legend>
                        <ul>
                            <li>serialNum (�Ϸù�ȣ) : <%=taxInfo.detailList(i).serialNum%></li>
                            <li>purchaseDT (�ŷ�����) : <%=taxInfo.detailList(i).purchaseDT%></li>
                            <li>itemName (ǰ��) : <%=taxInfo.detailList(i).itemName%></li>
                            <li>spec (�԰�) : <%=taxInfo.detailList(i).spec%></li>
                            <li>qty (����) : <%=taxInfo.detailList(i).qty%></li>
                            <li>unitCost (�ܰ�) : <%=taxInfo.detailList(i).unitCost%></li>
                            <li>supplyCost (���ް���) : <%=taxInfo.detailList(i).supplyCost%></li>
                            <li>tax (����) : <%=taxInfo.detailList(i).tax%></li>
                            <li>remark (���) : <%=taxInfo.detailList(i).remark%></li>
                        </ul>
                        </fieldset>
                    <%
                        Next
                    %>
                    <%
                        For i=0 To UBound(taxInfo.addContactList)-1
                    %>
                        <fieldset class="fieldset2">
                            <legend>�߰������ ���� <%=i+1%> </legend>
                                <ul>
                                    <li>serialNum (�Ϸù�ȣ) : <%=taxInfo.addContactList(i).serialNum%></li>
                                    <li>email (����� ����) : <%=taxInfo.addContactList(i).email%></li>
                                    <li>contactName (����ڸ�) : <%=taxInfo.addContactList(i).contactName%></li>
                                </ul>
                            </fieldset>
                    <%
                        Next
                    %>
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