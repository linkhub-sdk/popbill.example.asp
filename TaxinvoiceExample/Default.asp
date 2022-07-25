<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=euc-kr" />
    <link rel="stylesheet" type="text/css" href="/Example.css" media="screen" />
    <title>팝빌 SDK ASP Example.</title>
</head>
<body>
<div id="content">
    <p class="heading1">팝빌 세금계산서 SDK ASP Example.</p>
    <br/>
    <fieldset class="fieldset1">
        <legend>정방행/역발행/위수탁발행</legend>
        <ul>
            <li><a href="checkMgtKeyInUse.asp">CheckMgtKeyInUse</a> - 문서번호 확인</li>
            <li><a href="registIssue.asp">RegistIssue</a> - 즉시 발행</li>
            <li><a href="bulkSubmit.asp">bulkSubmit</a> -  초대량 발행 접수</li>
            <li><a href="getBulkResult.asp">getBulkResult</a> -  초대량 접수 결과 확인</li>
            <li><a href="register.asp">Register</a> - 임시저장</li>
            <li><a href="update.asp">Update</a> - 수정</li>
            <li><a href="issue.asp">Issue</a> - 발행</li>
            <li><a href="cancelIssue.asp">CancelIssue</a> - 발행취소</li>
            <li><a href="delete.asp">Delete</a> - 삭제</li>
            <li><a href="registRequest.asp">RegistRequest</a> - [역발행] 즉시 요청</li>
            <li><a href="request.asp">Request</a> - 역발행요청</li>
            <li><a href="cancelRequest.asp">CancelRequest</a> - 역발행요청 취소</li>
            <li><a href="refuse.asp">Refuse</a> - 역발행요청 거부</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>국세청 즉시 전송</legend>
        <ul>
            <li><a href="sendToNTS.asp">SendToNTS</a> - 국세청 즉시전송</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>세금계산서 정보확인</legend>
        <ul>
            <li><a href="getInfo.asp">GetInfo</a> - 상태 확인</li>
            <li><a href="getInfos.asp">GetInfos</a> - 상태 대량 확인</li>
            <li><a href="getDetailInfo.asp">GetDetailInfo</a> - 상세정보 확인</li>
            <li><a href="search.asp">Search</a> - 목록 조회</li>
            <li><a href="getLogs.asp">GetLogs</a> - 상태 변경이력 확인</li>
            <li><a href="getURL.asp">GetURL</a> - 세금계산서 문서함 관련 URL</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>세금계산서 보기/인쇄</legend>
        <ul>
            <li><a href="getPopUpURL.asp">GetPopUpURL</a> - 세금계산서 보기 URL</li>
            <li><a href="getViewURL.asp">GetViewURL</a> - 세금계산서 보기 URL - 메뉴/버튼 제외</li>
            <li><a href="getPrintURL.asp">GetPrintURL</a> - 세금계산서 인쇄 [공급자/공급받는자] URL</li>
            <li><a href="getOldPrintURL.asp">getOldPrintURL</a> - 세금계산서 (구)인쇄 [공급자/공급받는자] URL</li>
            <li><a href="getEPrintURL.asp">GetEPrintURL</a> - 세금계산서 인쇄 [공급받는자용] URL</li>
            <li><a href="getMassPrintURL.asp">GetMassPrintURL</a> - 세금계산서 대량 인쇄 URL</li>
            <li><a href="getMailURL.asp">GetMailURL</a> - 세금계산서 메일링크 URL</li>
            <li><a href="getPDFURL.asp">GetPDFURL</a> - 세금계산서 PDF 다운로드 URL</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>부가기능</legend>
        <ul>
            <li><a href="getAccessURL.asp">GetAccessURL</a> - 팝빌 로그인 URL</li>
            <li><a href="getSealURL.asp"> GetSealURL</a> - 인감 및 첨부문서 등록 URL</li>
            <li><a href="attachFile.asp">AttachFile</a> - 첨부파일 추가</li>
            <li><a href="deleteFile.asp">DeleteFile</a> - 첨부파일 삭제</li>
            <li><a href="getFiles.asp">GetFiles</a> - 첨부파일 목록 확인</li>
            <li><a href="sendEmail.asp">SendEmail</a> - 메일 전송</li>
            <li><a href="sendSMS.asp">SendSMS</a> - 문자 전송</li>
            <li><a href="sendFAX.asp">SendFAX</a> - 팩스 전송</li>
            <li><a href="attachStatement.asp">AttachStatement</a> - 전자명세서 첨부</li>
            <li><a href="detachStatement.asp">DetachStatement</a> - 전자명세서 첨부해제</li>
            <li><a href="getEmailPublicKeys.asp">GetEmailPublicKeys</a> - 유통사업자 메일 목록 확인</li>
            <li><a href="assignMgtKey.asp">AssignMgtKey</a> - 문서번호 할당</li>
            <li><a href="listEmailConfig.asp">ListEmailConfig</a> - 세금계산서 알림메일 전송목록 조회</li>
            <li><a href="updateEmailConfig.asp">UpdateEmailConfig</a> - 세금계산서 알림메일 전송설정 수정</li>
            <li><a href="getSendToNTSConfig.asp">getSendToNTSConfig</a> - 국세청 전송 설정 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>공인인증서 관리</legend>
        <ul>
            <li><a href="getTaxCertURL.asp">GetTaxCertURL</a> - 공인인증서 등록 URL</li>
            <li><a href="getCertificateExpireDate.asp">GetCertificateExpireDate</a> - 공인인증서 만료일 확인</li>
            <li><a href="checkCertValidation.asp">CheckCertValidation</a> - 공인인증서 유효성 확인</li>
            <li><a href="getTaxCertInfo.asp">GetTaxCertInfo</a> - 공인인증서 정보 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>포인트 관리</legend>
        <ul>
            <li><a href="getBalance.asp">GetBalance</a> - 연동회원 잔여포인트 확인</li>
            <li><a href="getChargeURL.asp">GetChargeURL</a> - 연동회원 포인트충전 URL</li>
            <li><a href="getPaymentURL.asp">GetPaymentURL</a> - 연동회원 포인트 결재내역 URL</li>
            <li><a href="getUseHistoryURL.asp">GetUseHistoryURL</a> - 연동회원 포인트 사용내역 URL</li>
            <li><a href="getPartnerBalance.asp">GetPartnerBalance</a> - 파트너 잔여포인트 확인</li>
            <li><a href="getPartnerURL.asp">GetPartnerURL</a> - 파트너 포인트충전 URL</li>
            <li><a href="getUnitCost.asp">GetUnitCost</a> - 발행 단가 확인</li>
            <li><a href="getChargeInfo.asp">GetChargeInfo</a> - 과금정보 확인</li>
        </ul>
    </fieldset>
    <fieldset class="fieldset1">
        <legend>회원정보</legend>
        <ul>
            <li><a href="checkIsMember.asp">CheckIsMember</a> - 연동회원 가입여부 확인</li>
            <li><a href="checkID.asp">CheckID</a> - 아이디 중복 확인</li>
            <li><a href="joinMember.asp">JoinMember</a> - 연동회원 신규가입</li>
            <li><a href="getCorpInfo.asp">GetCorpInfo</a> - 회사정보 확인</li>
            <li><a href="updateCorpInfo.asp">UpdateCorpInfo</a> - 회사정보 수정</li>
            <li><a href="registContact.asp">RegistContact</a> - 담당자 등록</li>
            <li><a href="getContactInfo.asp">getContactInfo</a> - 담당자 정보 확인</li>
            <li><a href="listContact.asp">ListContact</a> - 담당자 목록 확인</li>
            <li><a href="updateContact.asp">UpdateContact</a> - 담당자 정보 수정</li>
        </ul>
    </fieldset>
</div>
</body>
</html>