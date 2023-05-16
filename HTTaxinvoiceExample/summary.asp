<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
        <title>�˺� SDK ASP Example.</title>
    </head>
<!--#include file="common.asp"-->
<%
    '**************************************************************
    ' ���� ���� Ȯ��(GetJobState API) �Լ��� ���� ���� ������ Ȯ�ε� �۾����̵� Ȱ���Ͽ� ������ ���ڼ��ݰ�꼭 ����/���� ������ ��� ������ ��ȸ�մϴ�.
    ' - ��� ���� : ���ڼ��ݰ�꼭 ���� �Ǽ�, ���ް��� �հ�, ���� �հ�, �հ� �ݾ�
    ' - https://developers.popbill.com/reference/httaxinvoice/asp/api/search#Summary
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

    ' �ŷ�ó ��ȣ / ����ڹ�ȣ (�����) / �ֹε�Ϲ�ȣ (����) / "9999999999999" (�ܱ���) �� �˻��ϰ��� �ϴ� ���� �Է�
    ' - ����ڹ�ȣ / �ֹε�Ϲ�ȣ�� ������('-')�� ������ ���ڸ� �Է�
    ' - ���Է½� ��ü��ȸ
    SearchString = ""

    On Error Resume Next

    Set result = m_HTTaxinvoiceService.Summary(testCorpNum, JobID, TIType, TaxType,  _
                            PurposeType, TaxRegIDYN, TaxRegIDType, TaxRegID, UserID, SearchString)

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
                        <li> count (���� ��� �Ǽ�) : <%=result.count%> </li>
                        <li> supplyCostTotal (���ް��� �հ�) : <%=result.supplyCostTotal%> </li>
                        <li> taxTotal (���� �հ�) : <%=result.taxTotal%> </li>
                        <li> amountTotal (�հ� �ݾ�) : <%=result.amountTotal%> </li>
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