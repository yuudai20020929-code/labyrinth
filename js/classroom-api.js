/**
 * 授業用 Classroom API クライアント（Google Apps Script 経由）
 */
(function (global) {
  const SESSION_KEY = 'agy_classroom_session';
  const TEACHER_TOKEN_KEY = 'agy_teacher_token';
  const LOGIN_CACHE_KEY = 'agy_login_cache_v1';

  const GET_ACTIONS = new Set(['studentLogin', 'saveProgress']);
  const FETCH_TIMEOUT_MS = 30000;
  const LOGIN_CACHE_TTL_MS = 5 * 60 * 1000;
  const SAVE_DEBOUNCE_MS = 4000;

  let saveQueue = Promise.resolve();
  let saveDebounceTimer = null;
  let pendingSaveArgs = null;
  let lastSavedHash = '';

  function getConfig() {
    return global.LABYRINTH_CONFIG || {};
  }

  function getGasUrl() {
    return (getConfig().gasUrl || '').trim();
  }

  function isConfigured() {
    return Boolean(getGasUrl());
  }

  function makeSessionKey(grade, className, studentNo) {
    return 'G' + grade + '-C' + className + '-N' + studentNo;
  }

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

  function buildLoginResult(data, grade, className, studentNo) {
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

  function getLoginCache(grade, className, studentNo) {
    try {
      const raw = sessionStorage.getItem(LOGIN_CACHE_KEY);
      if (!raw) return null;
      const cached = JSON.parse(raw);
      if (
        Number(cached.grade) !== Number(grade) ||
        Number(cached.className) !== Number(className) ||
        Number(cached.studentNo) !== Number(studentNo)
      ) {
        return null;
      }
      if (Date.now() - cached.at > LOGIN_CACHE_TTL_MS) return null;
      return cached.result;
    } catch (e) {
      return null;
    }
  }

  function setLoginCache(grade, className, studentNo, result) {
    try {
      sessionStorage.setItem(LOGIN_CACHE_KEY, JSON.stringify({
        grade: Number(grade),
        className: Number(className),
        studentNo: Number(studentNo),
        at: Date.now(),
        result
      }));
    } catch (e) {}
  }

  function clearLoginCache() {
    try {
      sessionStorage.removeItem(LOGIN_CACHE_KEY);
    } catch (e) {}
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
    clearLoginCache();
    lastSavedHash = '';
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

  async function studentLogin(grade, className, studentNo, options) {
    const force = options && options.force;
    if (!force) {
      const cached = getLoginCache(grade, className, studentNo);
      if (cached) return cached;
    }

    const data = await callApi({
      action: 'studentLogin',
      grade: Number(grade),
      className: Number(className),
      classNo: Number(className),
      studentNo: Number(studentNo)
    });
    const result = buildLoginResult(data, grade, className, studentNo);
    setLoginCache(grade, className, studentNo, result);
    return result;
  }

  function hashSavePayload(payload, reset) {
    return JSON.stringify({
      reset: Boolean(reset),
      cleared: payload.cleared,
      perfect: payload.perfect,
      masterCleared: payload.masterCleared,
      masterPerfect: payload.masterPerfect,
      medals: payload.medals
    });
  }

  async function runSaveProgress(gameProgress, medalIds, options) {
    const session = getStudentSession();
    if (!session) throw new Error('no_session');
    const reset = options && options.reset;
    const payload = sheetPayloadFromGameProgress(gameProgress, medalIds);
    const hash = hashSavePayload(payload, reset);
    if (!reset && hash === lastSavedHash) {
      return null;
    }

    const data = await callApi({
      action: 'saveProgress',
      grade: session.grade,
      className: session.className,
      classNo: session.className,
      studentNo: session.studentNo,
      ...payload,
      ...(reset ? { reset: 'true' } : {})
    });
    lastSavedHash = hash;
    return progressFromSheet(data.progress);
  }

  function saveProgress(gameProgress, medalIds, options) {
    const immediate = options && (options.reset || options.immediate);
    pendingSaveArgs = { gameProgress, medalIds, options: options || {} };

    if (immediate) {
      if (saveDebounceTimer) {
        clearTimeout(saveDebounceTimer);
        saveDebounceTimer = null;
      }
      saveQueue = saveQueue.then(() => runSaveProgress(
        pendingSaveArgs.gameProgress,
        pendingSaveArgs.medalIds,
        pendingSaveArgs.options
      ), () => runSaveProgress(
        pendingSaveArgs.gameProgress,
        pendingSaveArgs.medalIds,
        pendingSaveArgs.options
      ));
      return saveQueue;
    }

    return new Promise((resolve, reject) => {
      if (saveDebounceTimer) clearTimeout(saveDebounceTimer);
      saveDebounceTimer = setTimeout(() => {
        saveDebounceTimer = null;
        saveQueue = saveQueue.then(() => {
          if (!pendingSaveArgs) return null;
          return runSaveProgress(
            pendingSaveArgs.gameProgress,
            pendingSaveArgs.medalIds,
            pendingSaveArgs.options
          );
        }, () => {
          if (!pendingSaveArgs) return null;
          return runSaveProgress(
            pendingSaveArgs.gameProgress,
            pendingSaveArgs.medalIds,
            pendingSaveArgs.options
          );
        });
        saveQueue.then(resolve).catch(reject);
      }, SAVE_DEBOUNCE_MS);
    });
  }

  function flushSaveProgress() {
    if (!pendingSaveArgs) return Promise.resolve(null);
    if (saveDebounceTimer) {
      clearTimeout(saveDebounceTimer);
      saveDebounceTimer = null;
    }
    const args = pendingSaveArgs;
    saveQueue = saveQueue.then(
      () => runSaveProgress(args.gameProgress, args.medalIds, Object.assign({}, args.options, { immediate: true })),
      () => runSaveProgress(args.gameProgress, args.medalIds, Object.assign({}, args.options, { immediate: true }))
    );
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
    makeSessionKey,
    studentLogin,
    saveProgress,
    flushSaveProgress,
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
