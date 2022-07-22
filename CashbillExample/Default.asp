<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    <title>팝빌 SDK ASP Example.</title>
</head>
<body>
<div id="content">
    <p class="heading1">팝빌 현금영수증 SDK ASP Example.</p>
    <br/>
    <fieldset class="fieldset1">
        <legend>현금영수증 발행</legend>
        <ul>
            <li><a href="checkMgtKeyInUse.asp">checkMgtKeyInUse</a> -문서번호 확인</li>
            <li><a href="registIssue.asp">registIssue</a> -즉시발행</li>
            <li><a href="bulkSubmit.asp">bulkSubmit</a> -  초대량 발행 접수</li>
            <li><a href="getBulkResult.asp">getBulkResult</a> -  초대량 접수 결과 확인</li>
            <li><a href="cancelIssue.asp">cancelIssue</a> -발행취소</li>
            <li><a href="delete.asp">delete</a> -삭제</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>취소현금영수증 발행</legend>
        <ul>
            <li><a href="revokeRegistIssue.asp">revokeRegistIssue</a> -즉시발행</li>
            <li><a href="revokeRegistIssue_part.asp">revokeRegistIssue_part</a> -부분) 즉시발행</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>현금영수증 정보확인</legend>
        <ul>
            <li><a href="getInfo.asp">getInfo</a> -상태확인</li>
            <li><a href="getInfos.asp">getInfos</a> -상태 대량 확인</li>
            <li><a href="getDetailInfo.asp">getDetailInfo</a> -상세정보 확인</li>
            <li><a href="search.asp">search</a> -목록 조회</li>
            <li><a href="getLogs.asp">getLogs</a> -상태 변경이력 확인</li>
            <li><a href="getURL.asp">getURL</a> -현금영수증 문서함 관련 URL</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>현금영수증 보기/인쇄</legend>
        <ul>
            <li><a href="getPopUpURL.asp">getPopUpURL</a> -현금영수증 보기 URL</li>
            <li><a href="getViewURL.asp">getViewURL</a> -현금영수증 보기 URL (메뉴/버튼 제외)</li>
            <li><a href="getPrintURL.asp">getPrintURL</a> -현금영수증 인쇄 URL</li>
            <li><a href="getMassPrintURL.asp">getMassPrintURL</a> -현금영수증 대량 인쇄 URL</li>
            <li><a href="getMailURL.asp">getMailURL</a> -현금영수증 메일링크 URL</li>
            <li><a href="getPDFURL.asp">getPDFURL</a> -현금영수증 PDF 다운로드 URL</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>부가기능</legend>
        <ul>
            <li><a href="getAccessURL.asp">getAccessURL</a> -팝빌 로그인 URL</li>
            <li><a href="sendEmail.asp">sendEmail</a> -메일 전송</li>
            <li><a href="sendSMS.asp">sendSMS</a> -문자 전송</li>
            <li><a href="sendFAX.asp">sendFAX</a> -팩스 전송</li>
            <li><a href="assignMgtKey.asp">assignMgtKey</a> -현금영수증 문서번호 할당</li>
            <li><a href="listEmailConfig.asp">listEmailConfig</a> -현금영수증 알림메일 전송목록 조회</li>
            <li><a href="updateEmailConfig.asp">updateEmailConfig</a> -현금영수증 알림메일 전송설정 수정</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>포인트관리</legend>
        <ul>
            <li><a href="getBalance.asp">getBalance</a> -연동회원 잔여포인트 확인</li>
            <li><a href="getChargeURL.asp">getChargeURL</a> -연동회원 포인트충전 URL</li>
            <li><a href="getPaymentURL.asp">GetPaymentURL</a> - 연동회원 포인트 결재내역 URL</li>
            <li><a href="getUseHistoryURL.asp">GetUseHistoryURL</a> - 연동회원 포인트 사용내역 URL</li>
            <li><a href="getPartnerBalance.asp">getPartnerBalance</a> -파트너 잔여포인트 확인</li>
            <li><a href="getPartnerURL.asp">getPartnerURL</a> -파트너 포인트충전 URL</li>
            <li><a href="getUnitCost.asp">getUnitCost</a> -발행 단가 확인</li>
            <li><a href="getChargeInfo.asp">getChargeInfo</a> -과금정보 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>회원정보</legend>
        <ul>
            <li><a href="checkIsMember.asp">checkIsMember</a> -연동회원 가입여부 확인</li>
            <li><a href="checkID.asp">checkID</a> -아이디 중복 확인</li>
            <li><a href="joinMember.asp">joinMember</a> -연동회원 신규가입</li>
            <li><a href="getCorpInfo.asp">getCorpInfo</a> -회사정보 확인</li>
            <li><a href="updateCorpInfo.asp">updateCorpInfo</a> -회사정보 수정</li>
            <li><a href="registContact.asp">registContact</a> -담당자 등록</li>
            <li><a href="getContactInfo.asp">getContactInfo</a> - 담당자 정보 확인</li>
            <li><a href="listContact.asp">listContact</a> -담당자 목록 확인</li>
            <li><a href="updateContact.asp">updateContact</a> -담당자 정보 수정</li>
        </ul>
    </fieldset>
</div>
</body>
</html>