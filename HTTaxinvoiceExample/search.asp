<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���� ���� Ȯ��(GetJobState API) �Լ��� ���� ���� ������ Ȯ�ε� �۾����̵� Ȱ���Ͽ� ������ ���ڼ��ݰ�꼭 ����/���� ������ ��ȸ�մϴ�.
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#Search
    '**************************************************************

    ' �˺�ȸ�� ����ڹ�ȣ, "-" ����
    testCorpNum = "1234567890"

    ' �˺�ȸ�� ���̵�
    UserID = "testkorea"

    ' ���� ��û(requestJob) �� ��ȯ���� �۾����̵�(jobID)
    JobID = "019102415000000014"

    ' �������� �迭 ("N" �� "M" �� ����, ���� ���� ����)
    ' �� N = �Ϲ� , M = ����
    ' - ���Է� �� ��ü��ȸ
    Dim TIType(2)
    TIType(0) = "N"
    TIType(1) = "M"

    ' �������� �迭 ("T" , "N" , "Z" �� ����, ���� ���� ����)
    ' �� T = ����, N = �鼼, Z = ����
    ' - ���Է� �� ��ü��ȸ
    Dim TaxType(3)
    TaxType(0) = "T"
    TaxType(1) = "N"
    TaxType(2) = "Z"

    ' ������� �迭 ("R" , "C", "N" �� ����, ���� ���� ����)
    ' �� R = ����, C = û��, N = ����
    ' - ���Է� �� ��ü��ȸ
    Dim PurposeType(3)
    PurposeType(0) = "R"
    PurposeType(1) = "C"
    PurposeType(2) = "N"

    ' ��������ȣ ���� (null , "0" , "1" �� �� 1)
    ' - null = ��ü , 0 = ����, 1 = ����
    TaxRegIDYN = ""

    ' ��������ȣ�� ��ü ("S" , "B" , "T" �� �� 1)
    ' �� S = ������ , B = ���޹޴��� , T = ��Ź��
    ' - ���Է½� ��ü��ȸ
    TaxRegIDType = "S"

    ' ��������ȣ
    ' �ټ������ �޸�(",")�� �����Ͽ� ���� ex ) "0001,0002"
    ' - ���Է½� ��ü��ȸ
    TaxRegID = ""

    ' ������ ��ȣ
    Page  = 1

    ' �������� ��ϰ���
    PerPage = 10

    ' ���Ĺ���, D-��������, A-��������
    Order = "D"

    ' �ŷ�ó ��ȣ / ����ڹ�ȣ (�����) / �ֹε�Ϲ�ȣ (����) / "9999999999999" (�ܱ���) �� �˻��ϰ��� �ϴ� ���� �Է�
    ' - ����ڹ�ȣ / �ֹε�Ϲ�ȣ�� ������('-')�� ������ ���ڸ� �Է�
    ' - ���Է½� ��ü��ȸ
    SearchString = ""

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.Search(testCorpNum, JobID, TIType, TaxType, PurposeType, _
                                TaxRegIDYN, TaxRegIDType, TaxRegID, Page, PerPage, Order, UserID, SearchString)

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
                <legend>���� ��� ��ȸ</legend>
                <%
                    If code = 0 Then
                %>
                    <ul>
                        <li> code (�����ڵ�) : <%=result.code%> </li>
                        <li> message  (����޽���) : <%=result.message%> </li>
                        <li> total (�� �˻���� �Ǽ�) : <%=result.total%> </li>
                        <li> perPage (�������� �˻�����) : <%=result.perPage%> </li>
                        <li> pageNum (������ ��ȣ) : <%=result.pageNum%> </li>
                        <li> pageCount (������ ����) : <%=result.pageCount%> </li>
                    </ul>

                <%
                    For i=0 To UBound(result.list) -1
                %>
                    <fieldset class="fieldset2">
                        <legend>���ݰ�꼭 ���� [ <%=i+1%> / <%= UBound(result.list) %> ] </legend>
                            <ul>

                                <li> ntsconfirmNum (����û���ι�ȣ) : <%= result.list(i).ntsconfirmNum %></li>
                                <li> writeDate (�ۼ�����) : <%= result.list(i).writeDate %></li>
                                <li> issueDate (��������) : <%= result.list(i).issueDate %></li>
                                <li> sendDate (��������) : <%= result.list(i).sendDate %></li>
                                <li> taxType (��������) : <%= result.list(i).taxType %></li>
                                <li> invoiceType (����/����) : <%= result.list(i).invoiceType %></li>
                                <li> purposeType (����/û��) : <%= result.list(i).purposeType %></li>
                                <li> supplyCostTotal (���ް��� �հ�) : <%= result.list(i).supplyCostTotal %></li>
                                <li> taxTotal (���� �հ�) : <%= result.list(i).taxTotal %></li>
                                <li> totalAmount (�հ�ݾ�) : <%= result.list(i).totalAmount %></li>
                                <li> remark1 (���) : <%= result.list(i).remark1 %></li>
                                <li> purchaseDate (�ŷ�����) : <%= result.list(i).purchaseDate %></li>
                                <li> itemName (ǰ��) : <%= result.list(i).itemName %></li>
                                <li> spec (�԰�) : <%= result.list(i).spec %></li>
                                <li> qty (����) : <%= result.list(i).qty %></li>
                                <li> unitCost (�ܰ�) : <%= result.list(i).unitCost %></li>
                                <li> supplyCost (���ް���) : <%= result.list(i).supplyCost %></li>
                                <li> tax (����) : <%= result.list(i).tax %></li>
                                <li> remark (���) : <%= result.list(i).remark %></li>
                                <li> modifyYN (���� ���ڼ��ݰ�꼭 ���� ) : <%= result.list(i).modifyYN %></li>
                                <li> orgNTSConfirmNum (���� ���ڼ��ݰ�꼭 ����û���ι�ȣ) : <%= result.list(i).orgNTSConfirmNum %></li>
                                <br/>
                                <p><b>������ ����</b></p>
                                <li> invoicerCorpNum (����ڹ�ȣ) : <%= result.list(i).invoicerCorpNum %></li>
                                <li> invoicerTaxRegID (��������ȣ) : <%= result.list(i).invoicerTaxRegID %></li>
                                <li> invoicerCorpName (��ȣ) : <%= result.list(i).invoicerCorpName %></li>
                                <li> invoicerCEOName (��ǥ�� ����) : <%= result.list(i).invoicerCEOName %></li>
                                <li> invoicerEmail (����� �̸���) : <%= result.list(i).invoicerEmail %></li>
                                <br/>
                                <p><b>���޹޴��� ����</b></p>
                                <li> invoiceeCorpNum (����ڹ�ȣ) : <%= result.list(i).invoiceeCorpNum %></li>
                                <li> invoiceeType (���޹޴��� ����) : <%= result.list(i).invoiceeType %></li>
                                <li> invoiceeTaxRegID (��������ȣ) : <%= result.list(i).invoiceeTaxRegID %></li>
                                <li> invoiceeCorpName (��ȣ) : <%= result.list(i).invoiceeCorpName %></li>
                                <li> invoiceeCEOName (��ǥ�� ����) : <%= result.list(i).invoiceeCEOName %></li>
                                <li> invoiceeEmail1 (����� �̸���) : <%= result.list(i).invoiceeEmail1 %></li>
                                <br/>
                                <p><b>��Ź�� ����</b></p>
                                <li> trusteeCorpNum (����ڹ�ȣ) : <%= result.list(i).trusteeCorpNum %></li>
                                <li> trusteeTaxRegID (��������ȣ) : <%= result.list(i).trusteeTaxRegID %></li>
                                <li> trusteeCorpName (��ȣ) : <%= result.list(i).trusteeCorpName %></li>
                                <li> trusteeCEOName (��ǥ�� ����) : <%= result.list(i).trusteeCEOName %></li>
                                <li> trusteeEmail (����� �̸���) : <%= result.list(i).trusteeEmail %></li>

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

