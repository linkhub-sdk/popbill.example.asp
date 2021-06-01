<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    <title>팝빌 SDK ASP Example.</title>
</head>
<body>
<div id="content">
    <p class="heading1">팝빌 계좌조회 API SDK Example.</p>
    <br/>
    <fieldset class="fieldset1">
        <legend>계좌 관리</legend>
        <ul>
            <li><a href="registBankAccount.asp">registBankAccount</a> - 계좌 등록</li>
            <li><a href="updateBankAccount.asp">updateBankAccount</a> - 계좌 수정</li>
            <li><a href="getBankAccountInfo.asp">getBankAccountInfo</a> - 계좌정보 확인</li>
            <li><a href="listBankAccount.asp">listBankAccount</a> - 계좌 목록 확인</li>
            <li><a href="getBankAccountMgtURL.asp">getBankAccountMgtURL</a> - 계좌 관리 팝업 URL</li>
            <li><a href="closeBankAccount.asp">closeBankAccount</a> - 계좌 정액제 해지신청</li>
            <li><a href="revokeCloseBankAccount.asp">revokeCloseBankAccount</a> - 계좌 정액제 해지신청 취소</li>
            <li><a href="deleteBankAccount.asp">deleteBankAccount</a> - 종량제 계좌 삭제</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>계좌 거래내역 수집</legend>
        <ul>
            <li><a href="requestJob.asp">requestJob</a> - 수집 요청</li>
            <li><a href="getJobState.asp">getJobState</a> - 수집 상태 확인</li>
            <li><a href="listActiveJob.asp">listActiveJob</a> - 수집 상태 목록 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>계좌 거래내역 관리</legend>
        <ul>
            <li><a href="search.asp">search</a> - 수집 결과 조회</li>
            <li><a href="summary.asp">summary</a> - 수집 결과 요약정보 조회</li>
            <li><a href="saveMemo.asp">saveMemo</a> - 거래내역 메모 저장</li>

        </ul>
    </fieldset>

    <fieldset class="fieldset1">
        <legend>포인트 관리 / 정액제 신청</legend>
        <ul>
            <li><a href="getFlatRatePopUpURL.asp">getFlatRatePopUpURL</a> - 정액제 서비스 신청 URL</li>
            <li><a href="getFlatRateState.asp">getFlatRateState</a> - 정액제 서비스 상태 확인</li>
            <li><a href="getBalance.asp">getBalance</a> - 연동회원 잔여포인트 확인</li>
            <li><a href="getChargeURL.asp">getChargeURL</a> - 연동회원 포인트충전 URL</li>
            <li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
            <li><a href="getPartnerURL.asp">getPartnerURL</a> - 파트너 포인트충전 URL</li>
            <li><a href="getChargeInfo.asp">getChargeInfo</a> - 과금정보 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>회원정보</legend>
        <ul>
            <li><a href="checkIsMember.asp">checkIsMember</a> - 연동회원 가입여부 확인</li>
            <li><a href="checkID.asp">checkID</a> - 아이디 중복 확인</li>
            <li><a href="joinMember.asp">joinMember</a> - 연동회원 신규가입</li>
            <li><a href="getAccessURL.asp">getAccessURL</a> - 팝빌 로그인 URL</li>
            <li><a href="getCorpInfo.asp">getCorpInfo</a> - 회사정보 확인</li>
            <li><a href="updateCorpInfo.asp">updateCorpInfo</a> - 회사정보 수정</li>
            <li><a href="registContact.asp">registContact</a> - 담당자 등록</li>
            <li><a href="getContactInfo.asp">getContactInfo</a> - 담당자 정보 확인</li>
            <li><a href="listContact.asp">listContact</a> - 담당자 목록 확인</li>
            <li><a href="updateContact.asp">updateContact</a> - 담당자 정보 수정</li>
        </ul>
    </fieldset>
</div>
</body>
</html>