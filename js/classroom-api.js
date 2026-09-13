/**
 * 授業用 Classroom API クライアント（Google Apps Script 経由）
 */
(function (global) {
  const SESSION_KEY = 'agy_classroom_session';
  const TEACHER_TOKEN_KEY = 'agy_teacher_token';

  function getConfig() {
    return global.LABYRINTH_CONFIG || {};
  }

  function getGasUrl() {
    return (getConfig().gasUrl || '').trim();
  }

  function isConfigured() {
    return Boolean(getGasUrl());
  }

  const GET_ACTIONS = new Set(['studentLogin', 'saveProgress']);
  const FETCH_TIMEOUT_MS = 30000;
  let saveQueue = Promise.resolve();

  function validateApiData(action, data) {
    if (action === 'studentLogin') {
      if (!data.progress || !data.progress.key) {
        const err = new Error('invalid_student_login_response');
        err.code = 'gas_stale_deploy';
        throw err;
      }
    }
    if (action === 'saveProgress') {
      if (!data.progress || !data.progress.key) {
        const err = new Error('invalid_save_response');
        err.code = 'gas_stale_deploy';
        throw err;
      }
    }
  }

  function buildGetUrl(baseUrl, payload) {
    const params = new URLSearchParams();
    Object.keys(payload).forEach((key) => {
      const val = payload[key];
      if (val !== undefined && val !== null && val !== '') {
        params.set(key, String(val));
      }
    });
    const sep = baseUrl.includes('?') ? '&' : '?';
    return baseUrl + sep + params.toString();
  }

  function parseApiResponse(text) {
    const trimmed = String(text || '').trim();
    if (!trimmed) {
      const err = new Error('empty_response');
      err.code = 'empty_response';
      throw err;
    }
    if (trimmed.charAt(0) === '<') {
      const err = new Error('html_response');
      err.code = 'html_response';
      throw err;
    }
    try {
      return JSON.parse(trimmed);
    } catch (e) {
      const err = new Error('invalid_json');
      err.code = 'invalid_json';
      throw err;
    }
  }

  async function fetchGas(method, url, bodyText) {
    const isOrgGasUrl = /\/a\/macros\//.test(url);
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), FETCH_TIMEOUT_MS);
    const options = { method, redirect: 'follow', signal: controller.signal };
    if (method === 'POST') {
      options.headers = { 'Content-Type': 'text/plain' };
      options.body = bodyText;
    }
    let res;
    try {
      res = await fetch(url, options);
    } catch (err) {
      if (err && err.name === 'AbortError') {
        const timeoutErr = new Error('timeout');
        timeoutErr.code = 'timeout';
        throw timeoutErr;
      }
      throw err;
    } finally {
      clearTimeout(timer);
    }
    const text = await res.text();
    if (!res.ok) {
      const err = new Error('http_' + res.status);
      err.code = (res.status === 401 || res.status === 403)
        ? (isOrgGasUrl ? 'gas_org_only_cors' : 'gas_auth_required')
        : undefined;
      err.detail = text.slice(0, 120);
      throw err;
    }
    const data = parseApiResponse(text);
    if (!data || data.ok === false) {
      const err = new Error((data && data.error) || 'api_error');
      err.code = data && data.error;
      throw err;
    }
    return data;
  }

  async function callApi(payload) {
    const baseUrl = getGasUrl();
    if (!baseUrl) {
      throw new Error('gas_url_missing');
    }
    const isOrgGasUrl = /\/a\/macros\//.test(baseUrl);
    const action = payload.action;

    async function request(method, url, bodyText) {
      const data = await fetchGas(method, url, bodyText);
      validateApiData(action, data);
      return data;
    }

    try {
      // 生徒ログイン・進捗保存は GET 優先（失敗時 POST にフォールバック）
      if (GET_ACTIONS.has(action)) {
        try {
          return await request('GET', buildGetUrl(baseUrl, payload));
        } catch (getErr) {
          try {
            return await request('POST', baseUrl, JSON.stringify(payload));
          } catch (postErr) {
            throw getErr.code ? getErr : postErr;
          }
        }
      }
      return await request('POST', baseUrl, JSON.stringify(payload));
    } catch (err) {
      if (err && err.name === 'TypeError' && isOrgGasUrl) {
        const hint = new Error('gas_org_only_cors');
        hint.code = 'gas_org_only_cors';
        throw hint;
      }
      throw err;
    }
  }

  function csvToIndexMap(csv) {
    const map = {};
    String(csv || '')
      .split(',')
      .map((s) => s.trim())
      .filter(Boolean)
      .forEach((n) => {
        const stageNum = Number(n);
        if (!Number.isFinite(stageNum) || stageNum < 1) return;
        map[stageNum - 1] = true;
      });
    return map;
  }

  function indexMapToCsv(map) {
    return Object.keys(map || {})
      .filter((k) => map[k])
      .map((k) => Number(k) + 1)
      .filter((n) => Number.isFinite(n) && n >= 1)
      .sort((a, b) => a - b)
      .join(',');
  }

  function progressFromSheet(row) {
    return {
      standard: {
        cleared: csvToIndexMap(row.cleared),
        perfect: csvToIndexMap(row.perfect)
      },
      master: {
        cleared: csvToIndexMap(row.masterCleared),
        perfect: csvToIndexMap(row.masterPerfect)
      },
      medals: String(row.medals || ''),
      updatedAt: row.updatedAt || '',
      key: row.key || ''
    };
  }

  function sheetPayloadFromGameProgress(gameProgress, medalIds) {
    return {
      cleared: indexMapToCsv(gameProgress.standard && gameProgress.standard.cleared),
      perfect: indexMapToCsv(gameProgress.standard && gameProgress.standard.perfect),
      masterCleared: indexMapToCsv(gameProgress.master && gameProgress.master.cleared),
      masterPerfect: indexMapToCsv(gameProgress.master && gameProgress.master.perfect),
      medals: (medalIds || []).join(',')
    };
  }

  function getStudentSession() {
    try {
      const raw = sessionStorage.getItem(SESSION_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }

  function setStudentSession(session) {
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
  }

  function clearStudentSession() {
    sessionStorage.removeItem(SESSION_KEY);
  }

  function getTeacherToken() {
    try {
      return sessionStorage.getItem(TEACHER_TOKEN_KEY) || '';
    } catch (e) {
      return '';
    }
  }

  function setTeacherToken(token) {
    sessionStorage.setItem(TEACHER_TOKEN_KEY, token || '');
  }

  function clearTeacherToken() {
    sessionStorage.removeItem(TEACHER_TOKEN_KEY);
  }

  async function studentLogin(grade, className, studentNo) {
    const data = await callApi({
      action: 'studentLogin',
      grade: Number(grade),
      className: Number(className),
      classNo: Number(className),
      studentNo: Number(studentNo)
    });
    const session = {
      grade: Number(grade),
      className: Number(className),
      studentNo: Number(studentNo),
      key: data.progress.key
    };
    setStudentSession(session);
    return {
      session,
      progress: progressFromSheet(data.progress),
      created: data.created
    };
  }

  async function saveProgress(gameProgress, medalIds) {
    const session = getStudentSession();
    if (!session) throw new Error('no_session');
    const payload = sheetPayloadFromGameProgress(gameProgress, medalIds);
    const run = async () => {
      const data = await callApi({
        action: 'saveProgress',
        grade: session.grade,
        className: session.className,
        classNo: session.className,
        studentNo: session.studentNo,
        ...payload
      });
      return progressFromSheet(data.progress);
    };
    saveQueue = saveQueue.then(run, run);
    return saveQueue;
  }

  async function teacherLogin(password) {
    const data = await callApi({
      action: 'teacherLogin',
      password: String(password || '')
    });
    setTeacherToken(data.token);
    return data;
  }

  async function getClassProgress(grade, className) {
    const token = getTeacherToken();
    if (!token) throw new Error('unauthorized');
    const data = await callApi({
      action: 'getClassProgress',
      token,
      grade: Number(grade),
      className: Number(className)
    });
    return {
      grade: data.grade,
      className: data.className,
      students: (data.students || []).map((row) => ({
        ...row,
        parsed: progressFromSheet(row)
      }))
    };
  }

  global.ClassroomAPI = {
    isConfigured,
    getGasUrl,
    studentLogin,
    saveProgress,
    teacherLogin,
    getClassProgress,
    getStudentSession,
    setStudentSession,
    clearStudentSession,
    getTeacherToken,
    clearTeacherToken,
    progressFromSheet,
    sheetPayloadFromGameProgress,
    indexMapToCsv,
    csvToIndexMap
  };
})(window);
