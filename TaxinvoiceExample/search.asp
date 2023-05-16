<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' �˻������� ����Ͽ� ���ݰ�꼭 ����� ��ȸ�մϴ�. (��ȸ�Ⱓ ���� : �ִ� 6����)
    ' - https://developers.popbill.com/reference/taxinvoice/asp/api/info#Search
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ���� 10�ڸ�
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' �������� SELL(����), BUY(����), TRUSTEE(����Ź)
    KeyType = "SELL"

    ' ���� ���� ("R" , "W" , "I" �� �� 1)
    ' �� R = ������� , W = �ۼ����� , I = ��������
    DType = "W"

    ' ��������, yyyyMMdd
    SDate = "20220701"

    ' ��������, yyyyMMdd
    EDate = "20220720"

    ' �����ڵ� �迭 (2,3��° �ڸ��� ���ϵ�ī��(*) ��� ����)
    ' - ���Է½� ��ü��ȸ
    Dim State(2)
    State(0) = "3**"
    State(1) = "6**"


    ' ���� ���� �迭 ("N" , "M" �� ����, ���� ���� ����)
    ' - N = �Ϲ� ���ݰ�꼭 , M = ���� ���ݰ�꼭
    ' - ���Է½� ��ü��ȸ
    Dim TIType(2)
    TIType(0) = "N"
    TIType(1) = "M"

    ' �������� �迭 ("T" , "N" , "Z" �� ����, ���� ���� ����)
    ' - T = ���� , N = �鼼 , Z = ����
    ' - ���Է½� ��ü��ȸ
    Dim TaxType(3)
    TaxType(0) = "T"
    TaxType(1) = "N"
    TaxType(2) = "Z"

    ' �������� �迭 ("N" , "R" , "T" �� ����, ���� ���� ����)
    ' - N = ������ , R = ������ , T = ����Ź����
    ' - ���Է½� ��ü��ȸ
    Dim IssueType(3)
    IssueType(0) = "N"
    IssueType(1) = "R"
    IssueType(2) = "T"

    ' ������� �迭 ("P" , "H" �� ����, ���� ���� ����)
    ' - P = �˺�, H = Ȩ�ý� �Ǵ� �ܺ�ASP
    ' - ���Է½� ��ü��ȸ
    Dim RegType(2)
    RegType(0) = "P"
    RegType(1) = "H"

    ' ���޹޴��� ��������� �迭 ("N" , "0" , "1" , "2" , "3" , "4" �� ����, ���� ���� ����)
    ' - N = ��Ȯ�� , 0 = �̵�� , 1 = ��� , 2 = ��� , 3 = �޾� , 4 = Ȯ�ν���
    ' - ���Է½� ��ü��ȸ
    Dim CloseDownState(5)
    CloseDownState(0) = "N"
    CloseDownState(1) = "0"
    CloseDownState(2) = "1"
    CloseDownState(3) = "2"
    CloseDownState(4) = "3"

    ' �������� ���� (null , true , false �� �� 1)
    ' - null = ��ü��ȸ , true = �������� , false = �������
    LateOnly = null

    ' ���Ĺ���, A-��������, D-��������
    Order = "D"

    ' ������ ��ȣ
    Page = 1

    ' �������� �˻�����, �ִ� 1000
    PerPage = 5

    ' ��������ȣ�� ��ü ("S" , "B" , "T" �� �� 1)
    ' �� S = ������ , B = ���޹޴��� , T = ��Ź��
    ' - ���Է½� ��ü��ȸ
    TaxRegIDType = "S"

    ' ��������ȣ ���� (null , "0" , "1" �� �� 1)
    ' - null = ��ü , 0 = ����, 1 = ����
    TaxRegIDYN = ""

    ' ��������ȣ
    ' �ټ������ �޸�(",")�� �����Ͽ� ���� ex ) "0001,0002"
    ' - ���Է½� ��ü��ȸ
    TaxRegID = ""

    ' �ŷ�ó ��ȣ / ����ڹ�ȣ (�����) / �ֹε�Ϲ�ȣ (����) / "9999999999999" (�ܱ���) �� �˻��ϰ��� �ϴ� ���� �Է�
    ' - ����ڹ�ȣ / �ֹε�Ϲ�ȣ�� ������('-')�� ������ ���ڸ� �Է�
    ' - ���Է½� ��ü��ȸ
    QString = ""

    ' ���ݰ�꼭�� ������ȣ / ����û ���ι�ȣ �� �˻��ϰ��� �ϴ� ���� �Է�
    ' - ���Է½� ��ü��ȸ
    MgtKey = ""

    ' �������� ���� (null , "0" , "1" �� �� 1)
    ' - null = ��ü��ȸ , 0 = �Ϲݹ��� , 1 = ��������
    ' - �Ϲݹ��� : ���ݰ�꼭 �ۼ� �� API�� �ƴ� �˺� ����Ʈ�� ���� ����� ����
    ' - �������� : ���ݰ�꼭 �ۼ� �� API�� ���� ����� ����
    InterOPYN = ""

    On Error Resume Next

    Set result = m_TaxinvoiceService.Search(testCorpNum, KeyType, DType, SDate, EDate, State, TIType, TaxType, _
                        IssueType, RegType, CloseDownState, LateOnly, Order, Page, PerPage, TaxRegIDType, TaxRegIDYN, _
                        TaxRegID, QString, MgtKey, InterOPYN, UsreID)

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
                <%
                    If code = 0 Then
                %>
                        <legend>���ݰ�꼭 �����ȸ</legend>
                        <ul>
                            <li> code (�����ڵ�) : <%=result.code%></li>
                            <li> message (����޽���) : <%=result.message%></li>
                            <li> total (�� �˻���� �Ǽ�) : <%=result.total%></li>
                            <li> pageNum (������ ��ȣ) : <%=result.pageNum%></li>
                            <li> perPage (�������� ��ϰ���) : <%=result.perPage%></li>
                            <li> pageCount (������ ����) : <%=result.pageCount%></li>
                        </ul>
                        <%
                            For i=0 To UBound(result.list) -1
                        %>
                            <fieldset class="fieldset2">
                                <legend>  ���ݰ�꼭 ����/������� [ <%=i+1%> / <%=UBound(result.list)%> ]</legend>
                                    <ul>
                                        <li> itemKey (���ݰ�꼭 ������Ű) :  <%=result.list(i).itemKey%> </li>
                                        <li> taxType (��������) :  <%=result.list(i).taxType%> </li>
                                        <li> writeDate (�ۼ�����) :  <%=result.list(i).writeDate%> </li>
                                        <li> regDT (�ӽ����� ����) :  <%=result.list(i).regDT%> </li>
                                        <li> issueType (��������) :  <%=result.list(i).issueType %> </li>
                                        <li> supplyCostTotal (���ް��� �հ�) :  <%=result.list(i).supplyCostTotal%> </li>
                                        <li> taxTotal (���� �հ�) :  <%=result.list(i).taxTotal%> </li>
                                        <li> purposeType (����/û��) :  <%=result.list(i).purposeType%> </li>
                                        <li> issueDT (�����Ͻ�) :  <%=result.list(i).issueDT%> </li>
                                        <li> lateIssueYN (�������� ����) :  <%=result.list(i).lateIssueYN%> </li>
                                        <li> preIssueDT (���࿹���Ͻ�) :  <%=result.list(i).preIssueDT%> </li>
                                        <li> openYN (���� ����) :  <%=result.list(i).openYN%> </li>
                                        <li> openDT (�����Ͻ�) :  <%=result.list(i).openDT%> </li>
                                        <li> stateMemo (���¸޸�) :  <%=result.list(i).stateMemo%> </li>
                                        <li> stateCode (�����ڵ�) :  <%=result.list(i).stateCode%> </li>
                                        <li> stateDT (���� �����Ͻ�) :  <%=result.list(i).stateDT%> </li>
                                        <li> ntsconfirmNum (����û ���ι�ȣ) :  <%=result.list(i).ntsconfirmNum %> </li>
                                        <li> ntsresult (����û ���۰��) :  <%=result.list(i).ntsresult%> </li>
                                        <li> ntssendDT (����û �����Ͻ�) :  <%=result.list(i).ntssendDT%> </li>
                                        <li> ntsresultDT  (����û ��� �����Ͻ�) :  <%=result.list(i).ntsresultDT%> </li>
                                        <li> ntssendErrCode (���۽��� �����ڵ�) :  <%=result.list(i).ntssendErrCode%> </li>
                                        <li> modifyCode (���������ڵ�) : <%=result.list(i).modifyCode%></li>
                                        <li> interOPYN (������������) :  <%=result.list(i).interOPYN%> </li>
                                        <li> invoicerCorpName (������ ��ȣ) :  <%=result.list(i).invoicerCorpName%> </li>
                                        <li> invoicerCorpNum (������ ����ڹ�ȣ) :  <%=result.list(i).invoicerCorpNum%> </li>
                                        <li> invoicerMgtKey (������ ������ȣ) :  <%=result.list(i).invoicerMgtKey%> </li>
                                        <li> invoicerPrintYN (������ �μ⿩��) :  <%=result.list(i).invoicerPrintYN%> </li>
                                        <li> invoiceeCorpName (���޹޴��� ��ȣ) :  <%=result.list(i).invoiceeCorpName%> </li>
                                        <li> invoiceeCorpNum (���޹޴��� ����ڹ�ȣ) :  <%=result.list(i).invoiceeCorpNum%> </li>
                                        <li> invoiceeMgtKey (���޹޴��� ������ȣ) :  <%=result.list(i).invoiceeMgtKey%> </li>
                                        <li> invoiceePrintYN (���޹޴��� �μ⿩��) :  <%=result.list(i).invoiceePrintYN%> </li>
                                        <li> closeDownState (���޹޴��� ���������) :  <%=result.list(i).closeDownState%> </li>
                                        <li> closeDownStateDate (���޹޴��� ���������) :  <%=result.list(i).closeDownStateDate%> </li>
                                    </ul>
                                </fieldset>
                <%
                        Next
                    Else
                %>
                </fieldset>
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
