<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���ݿ����� 1���� �������� ��ȸ�մϴ�.
    ' - https://developers.popbill.com/reference/cashbill/asp/api/info#GetDetailInfo
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    userID = "testkorea"

    ' ������ȣ
    mgtKey = "20220720-ASP-001"

    On Error Resume Next

    Set Presponse = m_CashbillService.GetDetailInfo(testCorpNum, mgtKey, userID)

    If Err.Number <> 0 then
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
                <legend>���ݿ����� ������ Ȯ��</legend>
                <ul>
                    <% If code = 0 Then %>
                        <fieldset class="fieldset2">
                            <ul>
                                <li>mgtKey (������ȣ) : <%=Presponse.mgtKey%></li>
                                <li>confirmNum (����û���ι�ȣ) : <%=Presponse.confirmNum%></li>
                                <li>orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : <%=Presponse.orgConfirmNum%></li>
                                <li>orgTradeDate (���� ���ݿ����� �ŷ�����) : <%=Presponse.orgTradeDate%></li>
                                <li>tradeDate (�ŷ�����) : <%=Presponse.tradeDate%></li>
                                <li>tradeDT (�ŷ��Ͻ�) : <%=Presponse.tradeDT%></li>
                                <li>tradeType (��������) : <%=Presponse.tradeType %></li>
                                <li>tradeUsage (�ŷ�����) : <%=Presponse.tradeUsage%></li>
                                <li>tradeOpt (�ŷ�����) : <%=Presponse.tradeOpt %></li>
                                <li>taxationType (��������) : <%=Presponse.taxationType%></li>
                                <li>totalAmount (�ŷ��ݾ�) : <%=Presponse.totalAmount%></li>
                                <li>supplyCost (���ް���) : <%=Presponse.supplyCost%></li>
                                <li>tax (�ΰ���) : <%=Presponse.tax %></li>
                                <li>serviceFee (�����) : <%=Presponse.serviceFee%></li>
                                <li>franchiseCorpNum (������ ����ڹ�ȣ) : <%=Presponse.franchiseCorpNum%></li>
                                <li>franchiseTaxRegID (������ ������� �ĺ���ȣ) : <%=Presponse.franchiseTaxRegID%></li>
                                <li>franchiseCorpName (������ ��ȣ) : <%=Presponse.franchiseCorpName%></li>
                                <li>franchiseCEOName (������ ��ǥ�ڸ�) : <%=Presponse.franchiseCEOName%></li>
                                <li>franchiseAddr (������ �ּ�) : <%=Presponse.franchiseAddr%></li>
                                <li>franchiseTEL (������ ��ȭ��ȣ) : <%=Presponse.franchiseTEL %></li>
                                <li>identityNum (�ĺ���ȣ) : <%=Presponse.identityNum%></li>
                                <li>customerName (�ֹ��ڸ�) : <%=Presponse.customerName%></li>
                                <li>itemName (�ֹ���ǰ��) : <%=Presponse.itemName%></li>
                                <li>orderNumber (�ֹ���ȣ) : <%=Presponse.orderNumber%></li>
                                <li>email (�̸���) : <%=Presponse.email%></li>
                                <li>hp (�޴���) : <%=Presponse.hp%></li>
                                <li>smssendYN (�˸����� ���ۿ���) : <%=Presponse.smssendYN%></li>
                                <li>cancelType (��һ���) : <%=Presponse.cancelType %></li>
                            </ul>
                        </fieldset>
                    <%	Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If%>

                </ul>
            </fieldset>
         </div>
    </body>
</html>