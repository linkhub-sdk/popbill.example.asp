<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    <title>팝빌 SDK ASP Example.</title>
</head>
<body>
<div id="content">
    <p class="heading1">팝빌 문자 API SDK ASP Example.</p>
    <br/>
    <fieldset class="fieldset1">
        <legend>발신번호 사전등록</legend>
        <ul>
            <li><a href="checkSenderNumber.asp">checkSenderNumber</a> - 발신번호 등록여부 확인</li>
            <li><a href="getSenderNumberMgtURL.asp">getSenderNumberMgtURL</a> - 발신번호 관리 팝업 URL</li>
            <li><a href="getSenderNumberList.asp">getSenderNumberList</a> - 발신번호 목록 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>문자 전송</legend>
        <ul>
            <li><a href="sendSMS.asp">sendSMS</a> - 단문 전송</li>
            <li><a href="sendSMS_Multi.asp">sendSMS</a> - 단문 전송 [대량]</li>
            <li><a href="sendLMS.asp">sendLMS</a> - 장문 전송</li>
            <li><a href="sendLMS_Multi.asp">sendLMS</a> - 장문 전송 [대량]</li>
            <li><a href="sendMMS.asp">sendMMS</a> - 포토 전송</li>
            <li><a href="sendMMS_Multi.asp">sendMMS</a> - 포토 전송 [대량]</li>
            <li><a href="sendXMS.asp">sendXMS</a> - 단문/장문 자동인식 전송</li>
            <li><a href="sendXMS_Multi.asp">sendXMS</a> - 단문/장문 자동인식 전송 [대량]</li>
            <li><a href="cancelReserve.asp">cancelReserve</a> - 예약전송 취소</li>
            <li><a href="cancelReserveRN.asp">cancelReserveRN</a> - 예약전송 취소 (요청번호 할당)</li>
            <li><a href="cancelReservebyRCV.asp">cancelReservebyRCV</a> - 예약전송 취소 (접수번호, 수신번호)</li>
            <li><a href="cancelReserveRNbyRCV.asp">cancelReserveRNbyRCV</a> - 예약전송 취소 (요청번호, 수신번호)</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>정보확인</legend>
        <ul>
            <li><a href="getMessages.asp">getMessages</a> - 전송내역 확인</li>
            <li><a href="getMessagesRN.asp">getMessagesRN</a> - 전송내역 확인 (요청번호 할당)</li>
            <li><a href="search.asp">search</a> - 전송내역 목록 조회</li>
            <li><a href="getSentListURL.asp">getSentListURL</a> - 문자 전송내역 팝업 URL</li>
            <li><a href="getAutoDenyList.asp">getAutoDenyList</a> - 080 수신거부 목록 확인</li>
            <li><a href="CheckAutoDenyNumber.asp">CheckAutoDenyNumber</a> - 080 번호 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>포인트 관리</legend>
        <ul>
            <li><a href="getBalance.asp">getBalance</a> - 연동회원 잔여포인트 확인</li>
            <li><a href="getChargeURL.asp">getChargeURL</a> - 연동회원 포인트충전 URL</li>
            <li><a href="getPaymentURL.asp">GetPaymentURL</a> - 연동회원 포인트 결재내역 URL</li>
            <li><a href="getUseHistoryURL.asp">GetUseHistoryURL</a> - 연동회원 포인트 사용내역 URL</li>
            <li><a href="getPartnerBalance.asp">getPartnerBalance</a> - 파트너 잔여포인트 확인</li>
            <li><a href="getPartnerURL.asp">getPartnerURL</a> - 파트너 포인트충전 URL</li>
            <li><a href="getChargeInfo.asp">getChargeInfo</a> - 과금정보 확인</li>
            <li><a href="getUnitCost.asp">getUnitCost</a> - 전송 단가 확인</li>
            <li><a href="paymentRequest.asp">paymentRequest</a> - 연동회원 무통장 입금신청</li>
            <li><a href="getSettleResult.asp">getSettleResult</a> - 연동회원 무통장 입금신청 확인</li>
            <li><a href="getPaymentHistory.asp">getPaymentHistory</a> - 연동회원 포인트 결제내역 확인</li>
            <li><a href="getUseHistory.asp">getUseHistory</a> - 연동회원 포인트 사용내역 확인</li>
            <li><a href="refund.asp">refund</a> - 연동회원 포인트 환불신청</li>
            <li><a href="getRefundHistory.asp">getRefundHistory</a> - 연동회원 포인트 환불내역 확인</li>
            <li><a href="getRefundInfo.asp">getRefundInfo</a> - 환불 신청 상태 조회</li>
			<li><a href="getRefundableBalance.asp">getRefundableBalance</a> - 환불 가능 포인트 조회</li>
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
            <li><a href="quitMember.asp">quitMember</a> - 팝빌회원 탈퇴</li>
        </ul>
    </fieldset>
</div>
</body>
</html>