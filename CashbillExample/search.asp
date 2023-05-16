<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˻������� ����Ͽ� ���ݿ����� ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
    ' - https://developers.popbill.com/reference/cashbill/asp/api/info#Search
    '**************************************************************

    '�˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"


    ' ���� ���� ("R" , "T" , "I" �� �� 1)
    ' �� R = ������� , T = �ŷ����� , I = ��������
    DType = "T"

    '��������, yyyyMMdd
    SDate = "20220701"

    '��������, yyyyMMdd
    EDate = "20220720"

    ' �����ڵ� �迭 (2,3��° �ڸ��� ���ϵ�ī��(*) ��� ����)
    ' - ���Է½� ��ü��ȸ
    Dim State(1)
    State(0) = "3**"

    ' �������� �迭 ("N" , "C" �� ����, ���� ���� ����)
    ' - N = �Ϲ� ���ݿ����� , C = ��� ���ݿ�����
    ' - ���Է½� ��ü��ȸ
    Dim TradeType(2)
    TradeType(0) = "N"
    TradeType(1) = "C"

    ' �ŷ����� �迭 ("P" , "C" �� ����, ���� ���� ����)
    ' - P = �ҵ������ , C = ����������
    ' - ���Է½� ��ü��ȸ
    Dim TradeUsage(2)
    TradeUsage(0) = "P"
    TradeUsage(1) = "C"

    ' �ŷ����� �迭 ("N" , "B" , "T" �� ����, ���� ���� ����)
    ' - N = �Ϲ� , B = �������� , T = ���߱���
    ' - ���Է½� ��ü��ȸ
    Dim TradeOpt(3)
    TradeOpt(0) = "N"
    TradeOpt(1) = "B"
    TradeOpt(2) = "T"

    ' �������� �迭 ("T" , "N" �� ����, ���� ���� ����)
    ' - T = ���� , N = �����
    ' - ���Է½� ��ü��ȸ
    Dim TaxationType(2)
    TaxationType(0) = "T"
    TaxationType(1) = "N"


    ' ���Ĺ���, A-��������, D-��������
    Order = "D"

    ' ��������ȣ
    Page = 1

    ' �������� �˻�����, �ִ� 1000
    PerPage = 20

    ' �ĺ���ȣ ��ȸ, ����ó���� ��ü��ȸ
    QString = ""

    ' ������ ������� ��ȣ
    ' - �ټ��� �˻��� �޸�(",")�� ����. ��) 1234,1000
    ' - ���Է½� ��ü��ȸ
    FranchiseTaxRegID = ""

    On Error Resume Next

    Set SearchResult = m_CashbillService.Search(testCorpNum, DType, SDate, EDate, State, TradeType, TradeUsage, TradeOpt, TaxationType, Order, Page, PerPage, QString, FranchiseTaxRegID)

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
                <legend>���ݿ����� �����ȸ</legend>
                    <ul>
                        <li> code (���� �ڵ�) : <%=SearchResult.code%></li>
                        <li> message (���� �޽���) : <%=SearchResult.message%></li>
                        <li> total (�� �˻���� �Ǽ�) : <%=SearchResult.total%></li>
                        <li> pageNum (������ ��ȣ) : <%=SearchResult.pageNum%></li>
                        <li> perPage (�������� �˻�����) : <%=SearchResult.perPage%></li>
                        <li> pageCount (������ ����) : <%=SearchResult.pageCount%></li>
                    </ul>
                    <% If code = 0 Then
                        For i=0 To UBound(SearchResult.list)-1 %>
                        <fieldset class="fieldset2">
                            <legend> ���ݿ����� ��ȸ ��� [<%= i+1 %> / <%= SearchResult.total %>]</legend>
                            <ul>
                                <li>itemKey (���ݿ����� ������Ű) : <%=SearchResult.list(i).itemKey%></li>
                                <li>mgtKey (������ȣ) : <%=SearchResult.list(i).mgtKey%></li>
                                <li>tradeDate (�ŷ�����) : <%=SearchResult.list(i).tradeDate%></li>
                                <li>tradeDT (�ŷ��Ͻ�) : <%=SearchResult.list(i).tradeDT%></li>
                                <li>tradeType (��������) : <%=SearchResult.list(i).tradeType%></li>
                                <li>tradeUsage (�ŷ�����) : <%=SearchResult.list(i).tradeUsage%></li>
                                <li>tradeOpt (�ŷ�����) : <%=SearchResult.list(i).tradeOpt%></li>
                                <li>taxationType (��������) : <%=SearchResult.list(i).taxationType%></li>
                                <li>totalAmount (�ŷ��ݾ�) : <%=SearchResult.list(i).totalAmount%></li>
                                <li>issueDT (�����Ͻ�) : <%=SearchResult.list(i).issueDT%></li>
                                <li>regDT (����Ͻ�) : <%=SearchResult.list(i).regDT%></li>
                                <li>stateMemo (���¸޸�) : <%=SearchResult.list(i).stateMemo%></li>
                                <li>stateCode (�����ڵ�) : <%=SearchResult.list(i).stateCode%></li>
                                <li>stateDT (���º����Ͻ�) : <%=SearchResult.list(i).stateDT%></li>
                                <li>identityNum (�ŷ�ó �ĺ���ȣ) : <%=SearchResult.list(i).identityNum%></li>
                                <li>itemName (��ǰ��) : <%=SearchResult.list(i).itemName%></li>
                                <li>customerName (������) : <%=SearchResult.list(i).customerName%></li>
                                <li>confirmNum (����û ���ι�ȣ) : <%=SearchResult.list(i).confirmNum%></li>
                                <li>orgConfirmNum (���� ���ݿ����� ����û���ι�ȣ) : <%=SearchResult.list(i).orgConfirmNum%></li>
                                <li>orgTradeDate (���� ���ݿ����� �ŷ�����) : <%=SearchResult.list(i).orgTradeDate%></li>
                                <li>ntssendDT (����û �����Ͻ�) : <%=SearchResult.list(i).ntssendDT%></li>
                                <li>ntsresultDT (����û ó����� �����Ͻ�) : <%=SearchResult.list(i).ntsResultDT%></li>
                                <li>ntsresultCode (����û ó����� �����ڵ�) : <%=SearchResult.list(i).ntsResultCode%></li>
                                <li>ntsresultMessage (����û ó����� �޽���) : <%=SearchResult.list(i).ntsResultMessage%></li>
                                <li>printYN (�μ⿩��) : <%=SearchResult.list(i).printYN%></li>
                            </ul>
                        </fieldset>
                    <%	Next
                        Else %>
                        <li> Response.code : <%=code%> </li>
                        <li> Response.message : <%=message%> </li>
                    <% End If%>
                </fieldset>
         </div>
    </body>
</html>