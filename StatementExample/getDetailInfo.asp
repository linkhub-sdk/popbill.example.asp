<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ڸ����� 1���� ������ Ȯ���մϴ�.
    ' - https://developers.popbill.com/reference/statement/asp/api/info#GetDetailInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-"���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ �ڵ� - 121(�ŷ�������), 122(û����), 123(������), 124(���ּ�), 125(�Ա�ǥ), 126(������)
    itemCode = "121"

    ' ������ȣ
    mgtKey = "20220720-ASP-002"

    On Error Resume Next

    Set result = m_StatementService.GetDetailInfo(testCorpNum, itemCode, mgtKey, userID)

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
                <legend>���ڸ����� ������</legend>
                <ul>
                    <% If code = 0 Then %>
                        <li> itemCode(������ �ڵ�) : <%=result.itemCode%> </li>
                        <li> mgtKey(������ȣ) : <%=result.mgtKey%> </li>
                        <li> invoiceNum(�˺� ���ι�ȣ) : <%=result.invoiceNum%> </li>
                        <li> formCode(������ �ڵ�) : <%=result.formCode%> </li>
                        <li> writeDate(�ۼ�����) : <%=result.writeDate%> </li>
                        <li> taxType(��������) : <%=result.taxType %> </li>
                        <li> purposeType(����/û��) : <%=result.purposeType%> </li>
                        <li> serialNum(����� �Ϸù�ȣ) : <%=result.serialNum%> </li>
                        <li> taxTotal(���� �հ�) : <%=result.taxTotal%> </li>
                        <li> supplyCostTotal(���ް��� �հ�) : <%=result.supplyCostTotal%> </li>
                        <li> totalAmount(�հ�ݾ�) : <%=result.totalAmount%> </li>
                        <li> remark1(���1) : <%=result.remark1%> </li>
                        <li> remark2(���2) : <%=result.remark2%> </li>
                        <li> remark3(���3) : <%=result.remark3%> </li>
                        <li> senderCorpNum(�߽��� ����ڹ�ȣ) : <%=result.senderCorpNum%> </li>
                        <li> senderTaxRegID(�߽��� ��������ȣ) : <%=result.senderTaxRegID%> </li>
                        <li> senderCorpName(�߽��� ��ȣ) : <%=result.senderCEOName%> </li>
                        <li> senderCEOName(�߽��� ��ǥ�ڼ���) : <%=result.senderCEOName%> </li>
                        <li> senderAddr(�߽��� �ּ�) : <%=result.senderAddr%> </li>
                        <li> senderBizClass(�߽��� ����) : <%=result.senderBizClass%> </li>
                        <li> senderBizType(�߽��� ����) : <%=result.senderBizType%> </li>
                        <li> senderContactName(�߽��� ����ڸ�) : <%=result.senderContactName%> </li>
                        <li> senderTEL(�߽��� ����ó) : <%=result.senderTEL%> </li>
                        <li> senderHP(�߽��� �޴�����ȣ) : <%=result.senderHP%> </li>
                        <li> senderEmail(�߽��� �����ּ�) : <%=result.senderEmail%> </li>
                        <li> receiverCorpNum(������ ����ڹ�ȣ) : <%=result.receiverCorpNum%> </li>
                        <li> receiverTaxRegID(������ ��������ȣ) : <%=result.receiverTaxRegID%> </li>
                        <li> receiverCorpName(������ ��ȣ) : <%=result.receiverCorpName%> </li>
                        <li> receiverCEOName(������ ��ǥ�ڼ���) : <%=result.receiverCEOName%> </li>
                        <li> receiverAddr(������ �ּ�) : <%=result.receiverAddr%> </li>
                        <li> receiverBizClass(������ ����) : <%=result.receiverBizClass%> </li>
                        <li> receiverBizType(������ ����) : <%=result.receiverBizType%> </li>
                        <li> receiverContactName(������ ����ڸ�) : <%=result.receiverContactName%> </li>
                        <li> receiverTEL(������ ����ó) : <%=result.receiverTEL%> </li>
                        <li> receiverHP(������ �޴�����ȣ) : <%=result.receiverHP%> </li>
                        <li> receiverEmail(������ �����ּ�) : <%=result.receiverEmail%> </li>
                        <li> businessLicenseYN(����ڵ���� ÷�ο���) : <%=result.businessLicenseYN%> </li>
                        <li> bankBookYN(����纻 ÷�ο���) : <%=result.bankBookYN%> </li>
                        <li> smssendYN(�˸����� ���ۿ���) : <%=result.smssendYN%> </li>
                        <li> autoacceptYN(����� �ڵ����� ����) : <%=result.autoacceptYN%> </li>

                        <!--��Ÿ ���׸� ����-->

                        <fieldset class="fieldset2">
                            <legend>�߰��Ӽ�</legend>
                            <ul>
                            <% For Each propertyKey In result.propertyBag.keys() %>
                                <li> <%=propertyKey%> : <%=result.propertyBag.get(propertyKey)%></li>
                            <% Next %>
                            </ul>
                        </fieldset>
                        <% For i=0 To Ubound(result.detailList)-1%>
                                <fieldset class="fieldset2">
                                <legend> ���׸� <%=i+1%> </legend>
                                    <ul>
                                        <li> serialNum(�Ϸù�ȣ) : <%=result.detailList(i).serialNum%> </li>
                                        <li> purchaseDT(�ŷ�����) : <%=result.detailList(i).purchaseDT%> </li>
                                        <li> itemName(ǰ���) : <%=result.detailList(i).itemName%> </li>
                                        <li> spec(�԰�) : <%=result.detailList(i).spec%> </li>
                                        <li> qty(����) : <%=result.detailList(i).qty%> </li>
                                        <li> unitCost(�ܰ�) : <%=result.detailList(i).unitCost%> </li>
                                        <li> supplyCost(���ް���) : <%=result.detailList(i).supplyCost%> </li>
                                        <li> tax(����) : <%=result.detailList(i).tax%> </li>
                                        <li> remark(���) : <%=result.detailList(i).remark%> </li>
                                        <li> spare1(����1) : <%=result.detailList(i).spare1%> </li>
                                        <li> spare2(����2) : <%=result.detailList(i).spare2%> </li>
                                        <li> spare3(����3) : <%=result.detailList(i).spare3%> </li>
                                        <li> spare4(����4) : <%=result.detailList(i).spare4%> </li>
                                        <li> spare5(����5) : <%=result.detailList(i).spare5%> </li>
                                        <li> spare6(����6) : <%=result.detailList(i).spare6%> </li>
                                        <li> spare7(����7) : <%=result.detailList(i).spare7%> </li>
                                        <li> spare8(����8) : <%=result.detailList(i).spare8%> </li>
                                        <li> spare9(����9) : <%=result.detailList(i).spare9%> </li>
                                        <li> spare10(����10) : <%=result.detailList(i).spare10%> </li>
                                        <li> spare11(����11) : <%=result.detailList(i).spare11%> </li>
                                        <li> spare12(����12) : <%=result.detailList(i).spare12%> </li>
                                        <li> spare13(����13) : <%=result.detailList(i).spare13%> </li>
                                        <li> spare14(����14) : <%=result.detailList(i).spare14%> </li>
                                        <li> spare15(����15) : <%=result.detailList(i).spare15%> </li>
                                        <li> spare16(����16) : <%=result.detailList(i).spare16%> </li>
                                        <li> spare17(����17) : <%=result.detailList(i).spare17%> </li>
                                        <li> spare18(����18) : <%=result.detailList(i).spare18%> </li>
                                        <li> spare19(����19) : <%=result.detailList(i).spare19%> </li>
                                        <li> spare20(����20) : <%=result.detailList(i).spare20%> </li>
                                    </ul>
                                </fieldset>
                            <%
                                Next
                                Else
                            %>

                            <li>Response.code : <%=code%> </li>
                            <li>Response.message: <%=message%> </li>
                        <%
                            End If
                        %>
                </ul>
            </fieldset>
         </div>
    </body>
</html>