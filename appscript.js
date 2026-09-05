/**
 * 채널 정산 인보이스 자동화 - Apps Script
 * 
 * 구글시트 메뉴에서 채널/기간 선택 후 Cloud Run API를 호출하여
 * 인보이스 생성 및 Gmail 드래프트를 자동 생성합니다.
 * 
 * 스크립트 속성 필요:
 * - CLOUD_RUN_URL: Cloud Run 서비스 URL
 * - API_KEY: API 인증키
 * - SERVICE_ACCOUNT_EMAIL: GCP 서비스 계정 이메일
 * - PRIVATE_KEY: 서비스 계정 private key
 */

/**
 * 스크립트 속성에서 설정값 로드
 */
function getConfig(key) {
  return PropertiesService.getScriptProperties().getProperty(key);
}

/**
 * 스프레드시트 열릴 때 메뉴 생성
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 인보이스 자동화')
    .addItem('🧾 인보이스 생성', 'showInvoiceDialog')
    .addSeparator()
    .addItem('📋 채널 목록 확인', 'showChannelList')
    .addToUi();
}

/**
 * 서비스 계정으로 Cloud Run ID 토큰 발급
 */
function getIdToken() {
  const privateKey = getConfig('PRIVATE_KEY').replace(/\\n/g, '\n');
  const serviceAccountEmail = getConfig('SERVICE_ACCOUNT_EMAIL');
  const targetUrl = getConfig('CLOUD_RUN_URL');
  const now = Math.floor(Date.now() / 1000);

  // JWT Header
  const header = Utilities.base64EncodeWebSafe(JSON.stringify({alg: 'RS256', typ: 'JWT'}));

  // JWT Claim
  const claim = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: serviceAccountEmail,
    sub: serviceAccountEmail,
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600,
    target_audience: targetUrl,
  }));

  // 서명 (header.claim 전체를 서명)
  const signatureInput = header + '.' + claim;
  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(signatureInput, privateKey)
  );

  const jwt = signatureInput + '.' + signature;

  // Google OAuth2 토큰 교환
  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt,
    },
    muteHttpExceptions: true,
  });

  const data = JSON.parse(response.getContentText());
  if (data.error) {
    Logger.log('ID Token 발급 실패: ' + data.error + ' - ' + data.error_description);
    return null;
  }
  return data.id_token;
}

/**
 * Cloud Run API 호출 (GET)
 */
function callCloudRunGet(endpoint) {
  const url = getConfig('CLOUD_RUN_URL') + endpoint;
  const token = getIdToken();
  const apiKey = getConfig('API_KEY');

  const headers = { 'Authorization': 'Bearer ' + token };
  if (apiKey) headers['x-api-key'] = apiKey;

  return UrlFetchApp.fetch(url, {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true,
  });
}

/**
 * Cloud Run API 호출 (POST + JSON body)
 */
function callCloudRunPost(endpoint, body) {
  const url = getConfig('CLOUD_RUN_URL') + endpoint;
  const token = getIdToken();
  const apiKey = getConfig('API_KEY');

  const headers = {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json',
  };
  if (apiKey) headers['x-api-key'] = apiKey;

  return UrlFetchApp.fetch(url, {
    method: 'post',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
}

/**
 * 인보이스 생성 다이얼로그
 */
function showInvoiceDialog() {
  const html = HtmlService.createHtmlOutput(getInvoiceDialogHtml())
    .setWidth(520)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '📄 채널 정산 인보이스 생성');
}

/**
 * 인보이스 생성 API 호출 (단일)
 */
function createInvoice(channelIdx, year, month, period) {
  const payload = {
    channel_idx: parseInt(channelIdx),
    year: parseInt(year),
    month: parseInt(month),
    period: parseInt(period) || 0
  };

  try {
    const response = callCloudRunPost('/invoice', payload);
    const code = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    if (code === 200 && result.success) {
      return {
        success: true,
        message: result.message,
        total_count: result.total_count,
        total_amount: result.total_amount
      };
    } else {
      return {
        success: false,
        message: result.detail || result.message || '서버 오류 (' + code + ')'
      };
    }
  } catch (e) {
    return {
      success: false,
      message: '서버 연결 실패: ' + e.message
    };
  }
}

/**
 * 인보이스 일괄 생성 API 호출 (배치)
 * @param {Array} items - [{channel_idx, period}, ...]
 */
function createInvoiceBatch(year, month, items) {
  const payload = {
    year: parseInt(year),
    month: parseInt(month),
    items: items.map(function(it) {
      return { channel_idx: parseInt(it.channel_idx), period: parseInt(it.period) || 0 };
    })
  };

  try {
    const response = callCloudRunPost('/invoice/batch', payload);
    const code = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    if (code === 200) {
      return {
        success: true,
        total: result.total,
        succeeded: result.succeeded,
        failed: result.failed,
        results: result.results
      };
    } else {
      return {
        success: false,
        message: result.detail || result.message || '서버 오류 (' + code + ')'
      };
    }
  } catch (e) {
    return { success: false, message: '서버 연결 실패: ' + e.message };
  }
}

/**
 * 채널 목록 조회
 */
function showChannelList() {
  const ui = SpreadsheetApp.getUi();

  try {
    const response = callCloudRunGet('/channels');
    const code = response.getResponseCode();

    if (code !== 200) {
      ui.alert('오류', '채널 목록 조회 실패 (HTTP ' + code + ')', ui.ButtonSet.OK);
      return;
    }

    const data = JSON.parse(response.getContentText());
    const periodLabel = { monthly: '월1회', biweekly: '월2회' };
    const dateLabel = { checkin: '체크인', checkout: '체크아웃' };
    const channels = data.channels.map(function(c) {
      return c.channel_idx + ': ' + c.name +
        ' (' + (periodLabel[c.period] || c.period) +
        ', ' + (dateLabel[c.date_type] || c.date_type) + ')';
    }).join('\n');
    ui.alert('정산 대상 채널', channels.substring(0, 2000), ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('오류', '채널 목록 조회 실패: ' + e.message, ui.ButtonSet.OK);
  }
}

/**
 * 채널 목록 API에서 가져오기 (다이얼로그용)
 *
 * 채널 목록은 거의 변하지 않는 정적 데이터이므로 상수로 내장하여
 * 다이얼로그가 즉시 뜨도록 한다 (Cloud Run 콜드스타트/토큰발급 왕복 제거).
 * 서버 config(channel_config.py)와 동기화 필요 시 syncChannelsFromServer() 사용.
 */
var SETTLEMENT_CHANNELS = [
  { channel_idx: 85,  name: 'DIDA(API)',      template: 'pkfare', date_type: 'checkout', period: 'biweekly' },
  { channel_idx: 108, name: 'Traveloka(API)', template: 'pkfare', date_type: 'checkout', period: 'monthly' },
  { channel_idx: 103, name: 'CN TRAVEL',      template: 'pkfare', date_type: 'checkout', period: 'monthly' },
  { channel_idx: 113, name: 'MG',             template: 'pkfare', date_type: 'checkin',  period: 'biweekly' },
  { channel_idx: 95,  name: 'PKFare(API)',    template: 'pkfare', date_type: 'checkin',  period: 'biweekly' },
  { channel_idx: 98,  name: '누아(API)',       template: 'nuua',   date_type: 'checkout', period: 'biweekly' },
  { channel_idx: 87,  name: '다보_B2B',        template: 'nuua',   date_type: 'checkout', period: 'monthly' },
  { channel_idx: 114, name: '엠스퀘어(l2u)',    template: 'kkday',  date_type: 'checkout', period: 'monthly' },
  { channel_idx: 5,   name: '온다',            template: 'kkday',  date_type: 'checkout', period: 'monthly' },
  { channel_idx: 102, name: 'KKDAY',          template: 'kkday',  date_type: 'checkin',  period: 'monthly' },
  { channel_idx: 51,  name: '트립토파즈',       template: 'kkday',  date_type: 'checkin',  period: 'monthly' }
];

function getChannelListForDialog() {
  return SETTLEMENT_CHANNELS;
}

/**
 * (선택) 서버에서 채널 목록을 다시 받아 확인용으로 반환.
 * 상수와 서버가 다를 경우 점검에 사용.
 */
function syncChannelsFromServer() {
  try {
    const response = callCloudRunGet('/channels');
    if (response.getResponseCode() === 200) {
      return JSON.parse(response.getContentText()).channels;
    }
    return [];
  } catch (e) {
    return [];
  }
}

/**
 * 인보이스 다이얼로그 HTML
 */
function getInvoiceDialogHtml() {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1;

  return `
    <html>
    <head>
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        html, body { height: auto; }
        body {
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
          padding: 24px;
          padding-bottom: 40px;
          color: #1f2937;
          background: #f8fafc;
          line-height: 1.5;
          overflow-y: auto;
        }
        .header {
          display: flex;
          align-items: center;
          gap: 10px;
          margin-bottom: 24px;
          padding-bottom: 16px;
          border-bottom: 1px solid #e5e7eb;
        }
        .header-icon { font-size: 24px; }
        .header h2 {
          font-size: 18px;
          font-weight: 700;
          color: #111827;
        }
        .header p {
          font-size: 12px;
          color: #6b7280;
          margin-top: 2px;
        }

        .form-group { margin-bottom: 18px; position: relative; }
        .form-label {
          display: block;
          font-size: 12px;
          font-weight: 600;
          color: #374151;
          margin-bottom: 6px;
          text-transform: uppercase;
          letter-spacing: 0.5px;
        }
        .form-input, .form-select {
          width: 100%;
          padding: 10px 12px;
          border: 1.5px solid #d1d5db;
          border-radius: 8px;
          font-size: 14px;
          color: #1f2937;
          background: white;
          transition: border-color 0.2s, box-shadow 0.2s;
          outline: none;
        }
        .form-input:focus, .form-select:focus {
          border-color: #3b82f6;
          box-shadow: 0 0 0 3px rgba(59,130,246,0.1);
        }
        .form-input::placeholder { color: #9ca3af; }

        /* 채널 검색 */
        .channel-wrapper { position: relative; }
        .channel-input-wrap {
          position: relative;
        }
        .channel-input-wrap .search-icon {
          position: absolute;
          left: 12px;
          top: 50%;
          transform: translateY(-50%);
          color: #9ca3af;
          font-size: 14px;
          pointer-events: none;
        }
        .channel-input-wrap input {
          padding-left: 34px;
        }
        .channel-dropdown {
          position: absolute;
          top: 100%;
          left: 0;
          right: 0;
          margin-top: 4px;
          background: white;
          border: 1.5px solid #e5e7eb;
          border-radius: 10px;
          max-height: 220px;
          overflow-y: auto;
          display: none;
          z-index: 100;
          box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }
        .channel-dropdown.show { display: block; }
        .channel-dropdown::-webkit-scrollbar { width: 6px; }
        .channel-dropdown::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 3px; }
        .channel-item {
          padding: 9px 14px;
          cursor: pointer;
          font-size: 13px;
          display: flex;
          align-items: center;
          gap: 8px;
          transition: background 0.15s;
          border-bottom: 1px solid #f3f4f6;
        }
        .channel-item:last-child { border-bottom: none; }
        .channel-item:hover { background: #eff6ff; }
        .channel-item.selected { background: #dbeafe; font-weight: 600; }
        .channel-item .code {
          display: inline-block;
          background: #f3f4f6;
          color: #6b7280;
          font-size: 11px;
          padding: 2px 6px;
          border-radius: 4px;
          font-weight: 500;
          min-width: 28px;
          text-align: center;
        }
        .channel-item.selected .code { background: #bfdbfe; color: #1e40af; }
        .channel-badge {
          margin-top: 6px;
          font-size: 12px;
          color: #3b82f6;
          font-weight: 500;
          display: flex;
          align-items: center;
          gap: 4px;
        }
        .channel-badge.empty { color: #9ca3af; }

        /* 채널 정보 pill */
        .channel-info {
          display: flex;
          gap: 8px;
          margin-bottom: 4px;
          flex-wrap: wrap;
        }
        .info-pill {
          font-size: 12px;
          font-weight: 600;
          padding: 5px 12px;
          border-radius: 999px;
          background: #eff6ff;
          color: #1e40af;
          border: 1px solid #bfdbfe;
        }
        .no-results {
          padding: 16px;
          text-align: center;
          color: #9ca3af;
          font-size: 13px;
        }

        /* 그리드 */
        .row { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }
        .row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }

        /* 체크박스 채널 아이템 */
        .channel-item input[type="checkbox"] { margin: 0; cursor: pointer; }
        .channel-item .ch-name { flex: 1; }
        .channel-item .ch-meta { font-size: 11px; color: #9ca3af; }

        /* 카운트 배지 */
        .count-badge {
          display: inline-block;
          background: #3b82f6; color: white;
          font-size: 11px; font-weight: 700;
          min-width: 18px; text-align: center;
          padding: 1px 6px; border-radius: 999px;
          margin-left: 4px;
        }

        /* 선택된 채널 칩 목록 */
        .selected-list {
          display: flex; flex-direction: column; gap: 8px;
          max-height: 150px; overflow-y: auto;
        }
        .chip {
          display: flex; align-items: center; gap: 8px;
          background: white; border: 1.5px solid #e5e7eb;
          border-radius: 10px; padding: 8px 12px;
        }
        .chip-name { font-weight: 600; font-size: 13px; color: #1f2937; }
        .chip-meta { font-size: 11px; color: #6b7280; flex: 1; }
        .chip-period {
          padding: 4px 8px; border: 1px solid #d1d5db;
          border-radius: 6px; font-size: 12px; background: #f9fafb;
        }
        .chip-x {
          cursor: pointer; color: #9ca3af; font-size: 13px;
          padding: 2px 4px; border-radius: 4px;
        }
        .chip-x:hover { color: #ef4444; background: #fef2f2; }

        /* 액션 버튼 행 */
        .action-row {
          display: flex; gap: 10px; margin-top: 20px;
          align-items: stretch; width: 100%;
        }

        /* 라벨 우측 링크 버튼 (전체 선택/해제) */
        .label-row {
          display: flex; align-items: center; justify-content: space-between;
        }
        .link-btn {
          background: none; border: none; cursor: pointer;
          color: #3b82f6; font-size: 12px; font-weight: 600;
          padding: 2px 6px; border-radius: 6px;
        }
        .link-btn:hover { background: #eff6ff; text-decoration: underline; }
        .action-row .btn {
          margin-top: 0;
          white-space: nowrap;
          width: auto;
        }
        .action-row .btn-secondary {
          flex: 0 0 110px;
          padding: 13px 12px;
          background: #f3f4f6; color: #374151;
        }
        .action-row .btn-secondary:hover { background: #e5e7eb; }
        .action-row .btn-primary { flex: 1 1 auto; min-width: 0; }

        /* 결과 리스트 */
        .results { margin-top: 12px; display: flex; flex-direction: column; gap: 6px; }
        .result-item {
          display: flex; align-items: center; gap: 8px;
          padding: 8px 12px; border-radius: 8px; font-size: 13px;
        }
        .result-item.ok { background: #ecfdf5; }
        .result-item.fail { background: #fef2f2; }
        .r-name { font-weight: 600; min-width: 100px; }
        .r-detail { color: #059669; font-size: 12px; }
        .r-detail.err { color: #dc2626; }
        .status.partial {
          display: block;
          background: #fffbeb; color: #92400e; border: 1px solid #fde68a;
        }

        /* 버튼 */
        .btn {
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 8px;
          width: 100%;
          padding: 13px 20px;
          border: none;
          border-radius: 10px;
          font-size: 15px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.2s;
          margin-top: 20px;
        }
        .btn-primary {
          background: linear-gradient(135deg, #3b82f6, #2563eb);
          color: white;
          box-shadow: 0 4px 12px rgba(37,99,235,0.3);
        }
        .btn-primary:hover {
          background: linear-gradient(135deg, #2563eb, #1d4ed8);
          box-shadow: 0 6px 16px rgba(37,99,235,0.4);
          transform: translateY(-1px);
        }
        .btn-primary:disabled {
          background: #d1d5db;
          box-shadow: none;
          transform: none;
          cursor: not-allowed;
        }

        /* 상태 메시지 */
        .status {
          margin-top: 16px;
          padding: 14px 16px;
          border-radius: 10px;
          font-size: 13px;
          display: none;
          animation: slideIn 0.3s ease;
        }
        @keyframes slideIn { from { opacity: 0; transform: translateY(-8px); } to { opacity: 1; transform: translateY(0); } }
        .status.loading {
          display: flex;
          align-items: center;
          gap: 10px;
          background: #eff6ff;
          color: #1e40af;
          border: 1px solid #bfdbfe;
        }
        .status.success {
          display: block;
          background: #ecfdf5;
          color: #065f46;
          border: 1px solid #a7f3d0;
        }
        .status.error {
          display: block;
          background: #fef2f2;
          color: #991b1b;
          border: 1px solid #fecaca;
        }
        .spinner {
          width: 16px; height: 16px;
          border: 2px solid #bfdbfe;
          border-top: 2px solid #2563eb;
          border-radius: 50%;
          animation: spin 0.8s linear infinite;
        }
        @keyframes spin { to { transform: rotate(360deg); } }
        .success-detail {
          margin-top: 8px;
          display: flex;
          gap: 16px;
          font-size: 14px;
          font-weight: 600;
        }
      </style>
    </head>
    <body>
      <div class="header">
        <span class="header-icon">📄</span>
        <div>
          <h2>채널 정산 인보이스</h2>
          <p>여러 채널을 선택해 한번에 생성할 수 있습니다</p>
        </div>
      </div>

      <div class="row2">
        <div class="form-group">
          <label class="form-label">연도</label>
          <input type="number" class="form-input" id="year" value="${currentYear}" min="2024" max="2030">
        </div>
        <div class="form-group">
          <label class="form-label">월</label>
          <select class="form-select" id="month">
            ${Array.from({length: 12}, (_, i) => 
              '<option value="' + (i+1) + '"' + ((i+1) === currentMonth ? ' selected' : '') + '>' + (i+1) + '월</option>'
            ).join('')}
          </select>
        </div>
      </div>

      <div class="form-group channel-wrapper">
        <label class="form-label label-row">
          <span>채널 선택 (다중 가능)</span>
          <button type="button" class="link-btn" id="selectAllBtn" onclick="selectAll()">전체 선택</button>
        </label>
        <div class="channel-input-wrap">
          <span class="search-icon">🔍</span>
          <input type="text" class="form-input" id="channelSearch" 
                 placeholder="채널명 또는 코드로 검색 후 선택..." 
                 oninput="filterChannels()" onfocus="openDropdown()" autocomplete="off">
        </div>
        <div class="channel-dropdown" id="channelDropdown"></div>
        <div class="channel-badge empty" id="channelBadge">⬆ 채널을 검색해 선택하세요</div>
      </div>

      <!-- 선택된 채널 목록 -->
      <div class="form-group" id="selectedGroup" style="display:none;">
        <label class="form-label">선택된 채널 <span id="selectedCount" class="count-badge">0</span></label>
        <div class="selected-list" id="selectedList"></div>
      </div>

      <div class="action-row">
        <button class="btn btn-primary" id="submitBtn" onclick="submitBatch()">
          🧾 인보이스 생성
        </button>
      </div>

      <div class="status" id="status"></div>
      <div class="results" id="results"></div>

      <script>
        // 채널 목록을 서버 왕복 없이 즉시 주입 (지연 제거)
        var channels = ${JSON.stringify(SETTLEMENT_CHANNELS)};
        var selected = {};   // { channel_idx: {name, period, dateType, template, chosenPeriod} }

        var PERIOD_LABEL = { monthly: '월1회', biweekly: '월2회' };
        var DATE_LABEL = { checkin: '체크인', checkout: '체크아웃' };
        var TEMPLATE_LABEL = { pkfare: '영문', nuua: '누아/다보', kkday: 'KKDAY' };

        // 로드 즉시 안내 표시
        (function() {
          var badge = document.getElementById('channelBadge');
          if (badge) {
            badge.innerHTML = '<span class="empty">⬆ 채널을 검색해 선택하세요 (' + channels.length + '개)</span>';
          }
        })();

        function normalize(str) {
          return String(str).toLowerCase().replace(/[\\s()\\-_]/g, '');
        }

        function findChannel(idx) {
          for (var i = 0; i < channels.length; i++) {
            if (channels[i].channel_idx === idx) return channels[i];
          }
          return null;
        }

        function filterChannels() {
          var q = document.getElementById('channelSearch').value;
          document.getElementById('channelDropdown').classList.add('show');
          if (!q) { renderChannels(channels); return; }

          var qn = normalize(q);
          var filtered = channels.filter(function(c) {
            var target = normalize(c.channel_idx + ' ' + c.name);
            if (target.indexOf(qn) !== -1) return true;
            var qi = 0;
            for (var i = 0; i < target.length && qi < qn.length; i++) {
              if (target[i] === qn[qi]) qi++;
            }
            return qi === qn.length;
          });
          renderChannels(filtered);
        }

        function renderChannels(list) {
          var el = document.getElementById('channelDropdown');
          if (list.length === 0) {
            el.innerHTML = '<div class="no-results">검색 결과 없음</div>';
            return;
          }
          el.innerHTML = list.map(function(c) {
            var checked = selected[c.channel_idx] ? ' checked' : '';
            var cls = selected[c.channel_idx] ? ' selected' : '';
            return '<label class="channel-item' + cls + '">' +
              '<input type="checkbox"' + checked + ' onchange="toggleChannel(' + c.channel_idx + ')">' +
              '<span class="code">' + c.channel_idx + '</span>' +
              '<span class="ch-name">' + c.name + '</span>' +
              '<span class="ch-meta">' + (PERIOD_LABEL[c.period]||'') + ' · ' + (DATE_LABEL[c.date_type]||'') + '</span>' +
              '</label>';
          }).join('');
        }

        function toggleChannel(idx) {
          if (selected[idx]) {
            delete selected[idx];
          } else {
            var c = findChannel(idx);
            if (c) {
              selected[idx] = {
                name: c.name, period: c.period,
                dateType: c.date_type, template: c.template,
                chosenPeriod: (c.period === 'biweekly') ? '1' : '0'
              };
            }
          }
          renderSelected();
          filterChannels(); // 드롭다운 체크 상태 갱신
        }

        function selectAll() {
          var allSelected = channels.every(function(c){ return selected[c.channel_idx]; });
          if (allSelected) {
            selected = {};
          } else {
            channels.forEach(function(c) {
              if (!selected[c.channel_idx]) {
                selected[c.channel_idx] = {
                  name: c.name, period: c.period,
                  dateType: c.date_type, template: c.template,
                  chosenPeriod: (c.period === 'biweekly') ? '1' : '0'
                };
              }
            });
          }
          renderSelected();
          filterChannels();
        }

        function setPeriod(idx, val) {
          if (selected[idx]) selected[idx].chosenPeriod = val;
        }

        function renderSelected() {
          var keys = Object.keys(selected);
          document.getElementById('selectedCount').textContent = keys.length;
          document.getElementById('selectAllBtn').textContent =
            (keys.length === channels.length && channels.length > 0) ? '전체 해제' : '전체 선택';

          var group = document.getElementById('selectedGroup');
          if (keys.length === 0) { group.style.display = 'none'; return; }
          group.style.display = 'block';

          document.getElementById('selectedList').innerHTML = keys.map(function(idx) {
            var s = selected[idx];
            var meta = '<span class="chip-meta">' + (PERIOD_LABEL[s.period]||'') + ' · ' + (DATE_LABEL[s.dateType]||'') + ' · ' + (TEMPLATE_LABEL[s.template]||'') + '</span>';
            // 월2회면 차수 선택
            var periodSel = '';
            if (s.period === 'biweekly') {
              periodSel = '<select class="chip-period" onchange="setPeriod(' + idx + ', this.value)">' +
                '<option value="1"' + (s.chosenPeriod==='1'?' selected':'') + '>1차(1~15)</option>' +
                '<option value="2"' + (s.chosenPeriod==='2'?' selected':'') + '>2차(16~말)</option>' +
                '</select>';
            }
            return '<div class="chip">' +
              '<span class="chip-name">' + s.name + '</span>' + meta + periodSel +
              '<span class="chip-x" onclick="toggleChannel(' + idx + ')">✕</span>' +
              '</div>';
          }).join('');
        }

        function openDropdown() {
          document.getElementById('channelDropdown').classList.add('show');
          filterChannels();
        }

        document.addEventListener('click', function(e) {
          if (!e.target.closest('.channel-wrapper')) {
            document.getElementById('channelDropdown').classList.remove('show');
          }
        });

        function submitBatch() {
          var keys = Object.keys(selected);
          if (keys.length === 0) {
            showStatus('error', '❌ 채널을 하나 이상 선택해주세요');
            return;
          }
          var year = document.getElementById('year').value;
          var month = document.getElementById('month').value;
          var items = keys.map(function(idx) {
            return { channel_idx: parseInt(idx), period: parseInt(selected[idx].chosenPeriod) || 0 };
          });

          document.getElementById('submitBtn').disabled = true;
          document.getElementById('results').innerHTML = '';
          showStatus('loading', '<div class="spinner"></div>' + keys.length + '개 채널 처리 중... (채널별 DB조회→엑셀→드래프트)');

          google.script.run
            .withSuccessHandler(onBatchSuccess)
            .withFailureHandler(onFailure)
            .createInvoiceBatch(year, month, items);
        }

        function showStatus(type, html) {
          var el = document.getElementById('status');
          el.className = 'status ' + type;
          el.innerHTML = html;
        }

        function onBatchSuccess(res) {
          document.getElementById('submitBtn').disabled = false;
          if (!res.success) {
            showStatus('error', '❌ ' + (res.message || '배치 처리 실패'));
            return;
          }
          var cls = res.failed === 0 ? 'success' : (res.succeeded === 0 ? 'error' : 'partial');
          showStatus(cls, '완료: 총 ' + res.total + '건 · ✅ 성공 ' + res.succeeded + ' · ❌ 실패 ' + res.failed);

          // 채널별 결과 리스트
          document.getElementById('results').innerHTML = res.results.map(function(r) {
            var icon = r.success ? '✅' : '❌';
            var detail = r.success
              ? ('<span class="r-detail">' + r.total_count + '건 · ' + Number(r.total_amount).toLocaleString() + ' KRW</span>')
              : ('<span class="r-detail err">' + r.message + '</span>');
            return '<div class="result-item ' + (r.success?'ok':'fail') + '">' +
              '<span class="r-icon">' + icon + '</span>' +
              '<span class="r-name">' + r.channel_name + '</span>' + detail +
              '</div>';
          }).join('');
        }

        function onFailure(error) {
          document.getElementById('submitBtn').disabled = false;
          showStatus('error', '❌ 오류: ' + error.message);
        }
      </script>
    </body>
    </html>
  `;
}
