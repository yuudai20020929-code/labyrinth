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

  async function callApi(payload) {
    const url = getGasUrl();
    if (!url) {
      throw new Error('gas_url_missing');
    }
    const isOrgGasUrl = /\/a\/macros\//.test(url);
    try {
      // text/plain にして CORS プリフライトを避ける（GAS 定番パターン）
      const res = await fetch(url, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      if (!res.ok) {
        const err = new Error('http_' + res.status);
        err.code = (res.status === 401 || res.status === 403)
          ? (isOrgGasUrl ? 'gas_org_only_cors' : 'gas_auth_required')
          : undefined;
        throw err;
      }
      const data = await res.json();
      if (!data || data.ok === false) {
        const err = new Error((data && data.error) || 'api_error');
        err.code = data && data.error;
        throw err;
      }
      return data;
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
    const data = await callApi({
      action: 'saveProgress',
      grade: session.grade,
      className: session.className,
      studentNo: session.studentNo,
      ...payload
    });
    return progressFromSheet(data.progress);
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
