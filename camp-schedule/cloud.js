(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.ScheduleCloud = api;
})(typeof window === 'object' ? window : globalThis, function () {
  'use strict';

  function validDate(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const d = new Date(value + 'T12:00:00Z');
    return !isNaN(d) && d.toISOString().slice(0, 10) === value && value >= '2000-01-01' && value <= '2100-12-31';
  }
  function normalizeCamp(value) {
    if (!value || !Number.isInteger(value.term) || value.term < 1 || value.term > 10000 || !validDate(value.openDate)) throw new Error('营期格式不正确');
    const skips = value.skips || [];
    if (!Array.isArray(skips) || skips.some(n => !Number.isInteger(n) || n < 1 || n > 12)) throw new Error('休息周格式不正确');
    return { term: value.term, openDate: value.openDate, skips: [...new Set(skips)].sort((a, b) => a - b) };
  }
  function normalizeList(values) {
    if (!Array.isArray(values) || values.length > 200) throw new Error('营期格式不正确');
    const result = values.map(normalizeCamp);
    if (new Set(result.map(c => c.term)).size !== result.length) throw new Error('期数重复');
    return result;
  }
  function parseShare(hash) {
    try {
      const text = decodeURIComponent(String(hash || '').replace(/^#/, ''));
      if (!text) return [];
      return normalizeList(text.split(',').map(part => {
        const bits = part.split(':');
        if (bits.length < 2 || bits.length > 3 || !/^\d+$/.test(bits[0])) throw new Error('链接格式不正确');
        return { term: Number(bits[0]), openDate: bits[1], skips: bits[2] ? bits[2].split('.').map(Number) : [] };
      }));
    } catch (_) { return []; }
  }
  function readLegacy(storage, hash, campKey, doneKey) {
    let local = [], done = {}, warning = '';
    try { local = normalizeList(JSON.parse(storage.getItem(campKey) || '[]')); }
    catch (_) { warning = '本机旧排期无法直接读取，原始数据仍保留在浏览器中。'; }
    try {
      const value = JSON.parse(storage.getItem(doneKey) || '{}');
      if (value && !Array.isArray(value) && typeof value === 'object') done = value;
    } catch (_) { warning = '本机旧勾选记录无法直接读取，原始数据仍保留在浏览器中。'; }
    return { local, shared: parseShare(hash), done, warning };
  }
  function errorOf(error) {
    if (error && error.scheduleError) return error;
    const raw = String(error && error.message || error || '');
    const code = String(error && error.code || '');
    let kind = 'network', message = '暂时无法连接云端。没有确认保存成功，请刷新后核对。';
    if (/schedule_conflict/.test(raw)) { kind = 'conflict'; message = '这条记录已被其他人修改或删除。已重新读取团队排期，请核对后再操作。'; }
    else if (code === '42501' || /schedule_members_only/.test(raw)) { kind = 'member'; message = '这个账号还不是运营团队成员，请联系现有成员添加权限。'; }
    else if (['PGRST202', '42P01', '42883'].includes(code)) { kind = 'setup'; message = '云端排期服务尚未就绪，请联系管理员。'; }
    else if (code === '23505') { kind = 'duplicate'; message = '团队排期中已经有这个期数，请刷新后查看。'; }
    else if (/^22|^235/.test(code)) { kind = 'invalid'; message = '排期内容格式不正确，请检查期数、日期和休息周。'; }
    else if (/Invalid login credentials/i.test(raw)) { kind = 'auth'; message = '邮箱或密码不正确。'; }
    else if (/Email not confirmed/i.test(raw)) { kind = 'auth'; message = '请先完成邮箱验证。'; }
    else if (code === 'PGRST301' || /JWT expired|refresh token|session missing/i.test(raw)) { kind = 'auth'; message = '登录已过期，请重新登录。'; }
    const result = new Error(message);
    result.code = kind; result.scheduleError = true;
    return result;
  }

  class Store {
    constructor(client, program, onChange) {
      if (!['ai', 'crawler'].includes(program)) throw new Error('Unknown schedule program');
      this.client = client; this.program = program; this.onChange = onChange || (() => {});
      this.epoch = 0; this.readSequence = 0;
      this.state = { phase: 'loading', user: null, camps: [], tasks: {}, busy: false, online: false, realtime: false, lastSync: null, error: '' };
    }
    emit() { this.onChange(this.state); }
    async rpc(name, params = {}) {
      try {
        const result = await this.client.rpc(name, { p_program: this.program, ...params });
        if (result.error) throw result.error;
        return result.data;
      } catch (error) { throw errorOf(error); }
    }
    async start() {
      this.auth = this.client.auth.onAuthStateChange((_event, session) => {
        setTimeout(() => { if (!this.stopped) this.setSession(session); }, 0);
      });
      const epoch = this.epoch;
      try {
        const { data, error } = await this.client.auth.getSession();
        if (error) throw error;
        if (epoch === this.epoch) await this.setSession(data.session);
      } catch (error) {
        if (epoch === this.epoch) { this.state.phase = 'login'; this.state.error = errorOf(error).message; this.emit(); }
      }
    }
    disconnect() {
      clearInterval(this.timer); clearTimeout(this.debounce);
      if (this.channel) { this.client.removeChannel(this.channel); this.channel = null; }
    }
    async setSession(session, force = false) {
      const user = session && session.user;
      if (!force && user && this.state.user && user.id === this.state.user.id) return;
      this.epoch++; this.readSequence++; this.disconnect();
      this.state = { phase: user ? 'loading' : 'login', user: user || null, camps: [], tasks: {}, busy: false, online: false, realtime: false, lastSync: null, error: '' };
      this.emit();
      if (!user) return;
      const epoch = this.epoch;
      await this.refresh();
      if (epoch !== this.epoch) return;
      this.timer = setInterval(() => this.refresh(), 15000);
      if (this.timer.unref) this.timer.unref();
      this.subscribe();
    }
    subscribe() {
      const epoch = this.epoch;
      const refresh = () => {
        clearTimeout(this.debounce);
        this.debounce = setTimeout(() => { if (epoch === this.epoch) this.refresh(); }, 150);
      };
      this.channel = this.client.channel('schedule-' + this.program + '-' + epoch)
        .on('postgres_changes', { event: '*', schema: 'public', table: 'schedule_camps', filter: 'program=eq.' + this.program }, refresh)
        .on('postgres_changes', { event: '*', schema: 'public', table: 'schedule_tasks', filter: 'program=eq.' + this.program }, refresh)
        .subscribe(status => {
          if (epoch !== this.epoch) return;
          this.state.realtime = status === 'SUBSCRIBED'; this.emit();
          if (status === 'SUBSCRIBED') refresh();
        });
    }
    async refresh() {
      if (!this.state.user || this.state.busy) return false;
      const epoch = this.epoch, sequence = ++this.readSequence;
      try {
        const snapshot = await this.rpc('schedule_snapshot');
        if (epoch !== this.epoch || sequence !== this.readSequence) return false;
        const camps = snapshot.camps.map(c => ({ ...normalizeCamp({ ...c, openDate: c.open_date }), id: c.id, version: c.version, updatedAt: c.updated_at }));
        const tasks = {};
        for (const task of snapshot.tasks) tasks[task.camp_id + '|' + task.task_key] = task;
        Object.assign(this.state, { camps, tasks, phase: 'ready', online: true, lastSync: new Date(), error: '' });
        this.emit(); return true;
      } catch (error) {
        if (epoch !== this.epoch || sequence !== this.readSequence) return false;
        const e = errorOf(error);
        if (['member', 'auth', 'setup'].includes(e.code)) {
          this.state.camps = []; this.state.tasks = {}; this.state.phase = e.code === 'auth' ? 'login' : 'blocked';
        } else if (this.state.phase !== 'ready') this.state.phase = 'blocked';
        this.state.error = e.message; this.state.online = false; this.emit(); return false;
      }
    }
    async mutate(name, params) {
      if (this.state.phase !== 'ready' || !this.state.online || this.state.busy) throw errorOf(new Error('not ready'));
      const epoch = this.epoch;
      this.state.busy = true; this.readSequence++; this.emit();
      try {
        const result = await this.rpc(name, params);
        if (epoch !== this.epoch) throw errorOf({ message: 'session missing' });
        this.state.busy = false;
        await this.refresh();
        return result;
      } catch (error) {
        if (epoch === this.epoch) {
          this.state.busy = false;
          await this.refresh();
        }
        throw error;
      }
    }
    save(camp, original) {
      const value = normalizeCamp(camp);
      return this.mutate('save_schedule_camp', {
        p_id: original ? original.id : null, p_version: original ? original.version : null,
        p_term: value.term, p_open_date: value.openDate, p_skips: value.skips,
      });
    }
    remove(camp) { return this.mutate('delete_schedule_camp', { p_id: camp.id, p_version: camp.version }); }
    setTask(id, completed) {
      const divider = id.indexOf('|');
      if (divider < 0) throw new Error('Invalid task ID');
      const task = this.state.tasks[id];
      return this.mutate('set_schedule_task', { p_camp_id: id.slice(0, divider), p_task_key: id.slice(divider + 1), p_version: task ? task.version : 0, p_completed: completed });
    }
    importCamps(entries) { return this.mutate('import_schedule_camps', { p_camps: entries }); }
    async signIn(email, password) {
      const { data, error } = await this.client.auth.signInWithPassword({ email: email.trim(), password });
      if (error) throw errorOf(error);
      await this.setSession(data.session, true);
    }
    async signOut() {
      const { error } = await this.client.auth.signOut({ scope: 'local' });
      await this.setSession(null);
      if (error) throw errorOf(error);
    }
    stop() { this.stopped = true; this.epoch++; this.disconnect(); if (this.auth) this.auth.data.subscription.unsubscribe(); }
  }
  return { Store, normalizeCamp, parseShare, readLegacy, errorOf };
});
