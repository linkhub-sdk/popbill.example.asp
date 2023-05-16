<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
        <link rel="stylesheet" type="text/css" href="../Example.css" media="screen" />
        <title>�������ȸ API SDK ASP Example.</title>
    </head>
    <!--#include file="common.asp"-->
    <%
        '**************************************************************
        ' ����ڹ�ȣ 1�ǿ� ���� ����������� Ȯ���մϴ�.
        ' - https://developers.popbill.com/reference/closedown/asp/api/check#CheckCorpNum
        '**************************************************************

        ' �˺�ȸ�� ����ڹ�ȣ
        UserCorpNum = "1234567890"

        ' ��ȸ�� ����ڹ�ȣ
        CorpNum = request.QueryString("CorpNum")

        If CorpNum <> "" Then

            On Error Resume Next

            Set result = m_ClosedownService.checkCorpNum(UserCorpNum, CorpNum)

            If Err.Number <> 0 Then
                code = Err.Number
                message = Err.Description
                Err.Clears
            End If

            On Error GoTo 0
        End if
    %>
    <body>
        <div id="content">
            <p class="heading1">Response</p>
            <br/>
            <fieldset class="fieldset1">
                <legend>�������ȸ - �ܰ�</legend>
                    <div class ="fieldset4">
                    <form method= "GET" id="corpnum_form" action="checkCorpNum.asp">
                        <%
                            If IsEmpty(result) then
                        %>
                                <input class= "txtCorpNum left" type="text" placeholder="����ڹ�ȣ ����" id="CorpNum" name="CorpNum"  tabindex=1/>
                        <%
                            Else
                        %>
                                <input class= "txtCorpNum left" type="text" placeholder="����ڹ�ȣ ����" id="CorpNum" name="CorpNum"  value="<%=result.corpNum%>" tabindex=1/>
                        <%
                            End if
                        %>

                        <p class="find_btn find_btn01 hand" onclick="search()" tabindex=2>��ȸ</p>
                    </form>
                    </div>
            </fieldset>
            <%
                If Not IsEmpty(result) Then
            %>
                <fieldset class="fieldset2">
                    <legend>�������ȸ - �ܰ�</legend>
                    <br/>
                    <p class="info">> state (���������) : null-�˼�����, 0-��ϵ��� ���� ����ڹ�ȣ, 1-�����, 2-���, 3-�޾�</p>
                    <p class="info">> taxType (��� ����) : null-�˼�����, 10-�Ϲݰ�����, 20-�鼼������, 30-���̰�����, 31-���̰�����(���ݰ�꼭 �߱޻����), 40-�񿵸�����, �������</p>
                    <ul>
                        <li>����ڹ�ȣ (corpNum) : <%= result.corpNum%></li>
                        <li>��������� (state) : <%= result.state%></li>
                        <li>��������� (taxType) : <%= result.taxType%></li>
                        <li>��������� (stateDate) : <%= result.stateDate%></li>
                        <li>�������� ��ȯ���� (typeDate) : <%= result.typeDate%></li>
                        <li>����û Ȯ������ (checkDate) : <%= result.checkDate%></li>
                    </ul>
                </fieldset>
            <%
                End If
                If Not IsEmpty(code) then
            %>
                <fieldset class="fieldset2">
                    <legend>�������ȸ - �ܰ�</legend>
                    <ul>
                        <li>Response.code : <%= code %> </li>
                        <li>Response.message : <%= message %></li>
                    </ul>
                </fieldset>
            <%
                End If
            %>
         </div>

        <script type ="text/javascript">
             window.onload=function(){
                 document.getElementById('CorpNum').focus();
             }

             function search(){
                document.getElementById('corpnum_form').submit();
             }
         </script>
    </body>
</html>