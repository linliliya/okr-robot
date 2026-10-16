// 复制整个文件到 Supabase Edge Functions，函数名 sales-ai。
// 每次优先 OpenAI，服务故障时尝试 DeepSeek；登录和销售成员检查始终在服务器执行。
// Secrets：OPENAI_API_KEY、DEEPSEEK_API_KEY；可只配置一家。
// 可选：OPENAI_MODEL（默认 gpt-6-luna）、DEEPSEEK_MODEL（默认 deepseek-flash），
// OPENAI_VISION_MODEL、DEEPSEEK_VISION_MODEL（未填时使用各自的文本模型）。
// deepseek-flash 当前对应 V4.1-Flash，支持文字与图片；模型信息核对于 2026-10-16。
// 兼容旧 LLM_API_KEY，但只按官方 LLM_BASE_URL 识别归属，不向自定义地址发送密钥。

import { createClient } from "npm:@supabase/supabase-js@2";

type Provider = "openai" | "deepseek";
type Reason = "network" | "timeout" | "auth" | "quota" | "rate_limit" | "region" | "server" | "invalid_response" | "request" | "not_configured";
type ProviderStatus = {
  state: "available" | "unavailable" | "not_configured" | "unknown";
  reason: Reason | null;
  message: string;
  checked_at: string | null;
};
type AIStatus = Record<Provider, ProviderStatus>;
type ProviderConfig = { provider: Provider; key: string; base: string; model: string; visionModel: string; timeout: number };
type RequestBody = { action?: string; messages?: unknown[]; json?: boolean; vision?: boolean; stream?: boolean };
type UpstreamResult = { ok: boolean; status: number; data: any; validJSON: boolean };
type Failure = { ok: false; reason: Reason; message: string; retry: boolean; httpStatus: number; available: boolean };
type Generation = { ok: true; content: string; model: string; usage: unknown };

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type, x-region",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
const MAX_BODY = 40 * 1024 * 1024;
const MAX_STREAM_BYTES = 2 * 1024 * 1024;
const MAX_OUTPUT_CHARS = 1024 * 1024;
const LABELS: Record<Provider, string> = { openai: "OpenAI", deepseek: "DeepSeek" };
const env = (name: string) => (Deno.env.get(name) || "").trim();
const now = () => new Date().toISOString();

function unknownStatus(): ProviderStatus {
  return { state: "unknown", reason: null, message: "尚未检测", checked_at: null };
}

function reply(status: number, body: unknown) {
  return new Response(JSON.stringify(body), {
    status, headers: { ...CORS, "Content-Type": "application/json; charset=utf-8" },
  });
}

function fail(status: number, error: string, ai_status: AIStatus) {
  return reply(status, { error, ai_status, provider: null });
}

function legacyProvider(): Provider | null {
  const base = env("LLM_BASE_URL");
  if (!base) return "deepseek";
  try {
    const url = new URL(base);
    if (url.username || url.password || url.search || url.hash || !/^\/(?:v1\/?)?$/.test(url.pathname)) return null;
    if (url.origin === "https://api.openai.com") return "openai";
    if (url.origin === "https://api.deepseek.com") return "deepseek";
  } catch { /* 不识别的旧地址不参与密钥迁移 */ }
  return null;
}

function configuration(): ProviderConfig[] {
  const legacy = legacyProvider();
  return (["openai", "deepseek"] as Provider[]).map(provider => {
    const prefix = provider.toUpperCase();
    const explicitKey = env(prefix + "_API_KEY");
    const useLegacy = !explicitKey && legacy === provider && !!env("LLM_API_KEY");
    const model = env(prefix + "_MODEL") || (useLegacy ? env("LLM_MODEL") : "") || (provider === "openai" ? "gpt-6-luna" : "deepseek-flash");
    return {
      provider,
      key: explicitKey || (useLegacy ? env("LLM_API_KEY") : ""),
      base: provider === "openai" ? "https://api.openai.com/v1" : "https://api.deepseek.com",
      model,
      visionModel: env(prefix + "_VISION_MODEL") || (useLegacy ? env("LLM_VISION_MODEL") : "") || model,
      timeout: provider === "openai" ? 45_000 : 90_000,
    };
  });
}

function initialStatus(configs: ProviderConfig[]): AIStatus {
  const statuses: AIStatus = { openai: unknownStatus(), deepseek: unknownStatus() };
  for (const config of configs) {
    if (!config.key) statuses[config.provider] = {
      state: "not_configured", reason: "not_configured",
      message: `未配置 ${LABELS[config.provider]} API key`, checked_at: null,
    };
  }
  return statuses;
}

// 超时包括接收响应头和读取响应体；主服务最多 45 秒，留出备用服务的 90 秒。
async function fetchUpstream(url: string, init: RequestInit, timeout: number): Promise<UpstreamResult> {
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let timedOut = false;
  try {
    return await Promise.race([
      (async () => {
        const response = await fetch(url, { ...init, signal: controller.signal });
        const text = await response.text();
        let data: any = null, validJSON = false;
        try { data = JSON.parse(text); validJSON = true; } catch { /* 在调用处判断无效响应 */ }
        return { ok: response.ok, status: response.status, data, validJSON };
      })(),
      new Promise<never>((_resolve, reject) => {
        timer = setTimeout(() => {
          timedOut = true;
          controller.abort();
          reject(new Error("timeout"));
        }, timeout);
      }),
    ]);
  } catch {
    throw new Error(timedOut ? "timeout" : "network");
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

function failure(provider: Provider, reason: Reason, retry = true, available = false): Failure {
  const name = LABELS[provider];
  const messages: Record<Reason, string> = {
    network: `${name} 连接失败`,
    timeout: `${name} 响应超时`,
    auth: `${name} API key 无效或没有访问权限`,
    quota: `${name} 余额或调用额度不足`,
    rate_limit: `${name} 请求过于频繁，请稍后重试`,
    region: `${name} 当前服务地区不可用`,
    server: `${name} 服务暂时异常`,
    invalid_response: `${name} 返回了空内容或无法识别的结果`,
    request: `${name} 无法处理此次请求，请检查输入内容和模型设置`,
    not_configured: `未配置 ${name} API key`,
  };
  return { ok: false, reason, message: messages[reason], retry, available,
    httpStatus: reason === "timeout" ? 504 : reason === "request" ? 400 : 502 };
}

function upstreamFailure(provider: Provider, result: UpstreamResult): Failure {
  // 原始上游文字只用于分类，不向网页返回，避免暴露密钥或服务内部信息。
  const err = result.data?.error;
  const detail = [err?.code, err?.type, err?.message].filter(x => typeof x === "string").join(" ").toLowerCase();
  if (/content_filter|content_policy|policy_violation|moderation|safety|refusal/.test(detail)) {
    return failure(provider, "request", false, true);
  }
  if (result.status === 400 || result.status === 404 || result.status === 422) return failure(provider, "request", false, true);
  if (/unsupported_country|country.*not.*support|region|territor|geographic/.test(detail)) return failure(provider, "region");
  if (result.status === 402 || /insufficient_quota|insufficient.*(balance|credit|fund)|billing_hard_limit|exceeded.*current quota|余额不足/.test(detail)) return failure(provider, "quota");
  if (result.status === 401 || result.status === 403) return failure(provider, "auth");
  if (result.status === 429) return failure(provider, "rate_limit");
  if (result.status === 408 || result.status === 504) return failure(provider, "timeout");
  if (result.status >= 500) return failure(provider, "server");
  return failure(provider, "request", false, true);
}

function failureStatus(result: Failure): ProviderStatus {
  return { state: result.available ? "available" : "unavailable", reason: result.reason,
    message: result.message, checked_at: now() };
}

async function probe(config: ProviderConfig): Promise<ProviderStatus> {
  if (!config.key) return initialStatus([config])[config.provider];
  try {
    const result = await fetchUpstream(config.base + "/models", {
      method: "GET", headers: { Authorization: "Bearer " + config.key },
    }, 10_000);
    if (!result.ok) {
      const status = failureStatus(upstreamFailure(config.provider, result));
      status.state = "unavailable";
      return status;
    }
    if (!result.validJSON || !Array.isArray(result.data?.data)) return failureStatus(failure(config.provider, "invalid_response"));
    return { state: "available", reason: null, message: "连接验证通过", checked_at: now() };
  } catch (error) {
    return failureStatus(failure(config.provider, error instanceof Error && error.message === "timeout" ? "timeout" : "network"));
  }
}

function generationPayload(config: ProviderConfig, body: RequestBody, stream: boolean) {
  const model = body.vision ? config.visionModel : config.model;
  const payload: Record<string, unknown> = { model, messages: body.messages, stream };
  if (stream) payload.stream_options = { include_usage: true };
  if (body.json) payload.response_format = { type: "json_object" };
  // Luna 默认为 medium；沿用原 mini 的非思考模式，避免升级后增加等待时间。
  // 仅对已核对支持 none 的型号设置，其他自定义模型仍使用其原有默认行为。
  if (config.provider === "openai" && model === "gpt-6-luna") payload.reasoning_effort = "none";
  return { model, payload };
}

async function generate(config: ProviderConfig, body: RequestBody): Promise<Failure | Generation> {
  const { model, payload } = generationPayload(config, body, false);
  try {
    const result = await fetchUpstream(config.base + "/chat/completions", {
      method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + config.key },
      body: JSON.stringify(payload),
    }, config.timeout);
    if (!result.ok) return upstreamFailure(config.provider, result);
    const choice = result.data?.choices?.[0];
    if (choice?.message?.refusal || choice?.finish_reason === "content_filter") return failure(config.provider, "request", false, true);
    const content = choice?.message?.content;
    if (!result.validJSON || typeof content !== "string" || !content.trim()) return failure(config.provider, "invalid_response");
    if (body.json) {
      try { JSON.parse(content); } catch { return failure(config.provider, "invalid_response"); }
    }
    return { ok: true, content, model, usage: result.data?.usage ?? null };
  } catch (error) {
    return failure(config.provider, error instanceof Error && error.message === "timeout" ? "timeout" : "network");
  }
}

// 两家使用 Chat Completions SSE；必须同时收到正常 finish_reason 和 [DONE] 才算完成。
async function generateStream(config: ProviderConfig, body: RequestBody, clientSignal: AbortSignal,
  onDelta: (text: string) => void): Promise<Failure | Generation> {
  const { model, payload } = generationPayload(config, body, true);
  const upstream = new AbortController();
  let reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  let timer: ReturnType<typeof setTimeout> | undefined;
  let timedOut = false;
  let rejectStop: (error: Error) => void = () => {};
  const stopped = new Promise<never>((_resolve, reject) => { rejectStop = reject; });
  const cancel = () => { upstream.abort(); rejectStop(new Error("cancelled")); };
  clientSignal.addEventListener("abort", cancel, { once: true });
  const invalid = () => failure(config.provider, "invalid_response");
  try {
    if (clientSignal.aborted) throw new Error("cancelled");
    timer = setTimeout(() => {
      timedOut = true; upstream.abort(); rejectStop(new Error("timeout"));
    }, config.timeout);
    return await Promise.race([stopped, (async (): Promise<Failure | Generation> => {
      const response = await fetch(config.base + "/chat/completions", {
        method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + config.key },
        body: JSON.stringify(payload), signal: upstream.signal,
      });
      if (upstream.signal.aborted) {
        void response.body?.cancel().catch(() => {});
        throw new Error(clientSignal.aborted ? "cancelled" : "timeout");
      }
      if (!response.body) return response.ok ? invalid() : upstreamFailure(config.provider,
        { ok: false, status: response.status, data: null, validJSON: false });
      reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8", { fatal: true });
      let bytes = 0;
      const readText = async () => {
        if (upstream.signal.aborted) throw new Error(clientSignal.aborted ? "cancelled" : "timeout");
        const chunk = await reader!.read();
        if (upstream.signal.aborted) throw new Error(clientSignal.aborted ? "cancelled" : "timeout");
        bytes += chunk.value?.byteLength || 0;
        if (bytes > MAX_STREAM_BYTES) throw invalid();
        try { return { done: chunk.done, text: chunk.done ? decoder.decode() : decoder.decode(chunk.value, { stream: true }) }; }
        catch { throw invalid(); }
      };
      if (!response.ok) {
        let text = "";
        for (;;) { const chunk = await readText(); text += chunk.text; if (chunk.done) break; }
        let data: any = null;
        try { data = JSON.parse(text); } catch { /* HTTP 状态仍可用于分类 */ }
        return upstreamFailure(config.provider, { ok: false, status: response.status, data, validJSON: data !== null });
      }

      let buffer = "", dataLines: string[] = [], content = "", finished = false, done = false;
      let usage: unknown = null;
      const dispatch = () => {
        if (!dataLines.length) return;
        const data = dataLines.join("\n"); dataLines = [];
        if (data.trim() === "[DONE]") { done = true; return; }
        let event: any;
        try { event = JSON.parse(data); } catch { throw invalid(); }
        if (event?.error) {
          const detail = [event.error.code, event.error.type].filter(x => typeof x === "string").join(" ").toLowerCase();
          const suppliedStatus = Number(event.error.status ?? event.error.status_code);
          const status = Number.isInteger(suppliedStatus) && suppliedStatus >= 400 && suppliedStatus <= 599 ? suppliedStatus
            : /invalid_api_key|authentication/.test(detail) ? 401 : /rate_limit/.test(detail) ? 429
            : /invalid_request|bad_request|context_length_exceeded|invalid_parameter/.test(detail) ? 400 : 500;
          throw upstreamFailure(config.provider, { ok: false, status, data: event, validJSON: true });
        }
        if (!Array.isArray(event?.choices)) throw invalid();
        if (event.usage != null) usage = event.usage;
        // OpenAI include_usage 可额外发送 choices=[]；DeepSeek 的 usage 可跟在结束块里。
        if (!event.choices.length) return;
        const choice = event.choices.find((c: any) => c && (c.index === 0 || c.index == null));
        if (!choice) throw invalid();
        if (choice.delta?.refusal || choice.finish_reason === "content_filter") throw failure(config.provider, "request", false, true);
        if (choice.finish_reason != null && choice.finish_reason !== "stop") throw invalid();
        const delta = choice.delta?.content;
        if (delta != null && typeof delta !== "string") throw invalid();
        if (delta) {
          if (finished || content.length + delta.length > MAX_OUTPUT_CHARS) throw invalid();
          content += delta; onDelta(delta);
        }
        if (choice.finish_reason === "stop") finished = true;
      };
      const parseLines = (eof: boolean) => {
        for (;;) {
          const end = buffer.search(/[\r\n]/);
          if (end < 0 || (buffer[end] === "\r" && end === buffer.length - 1 && !eof)) return;
          const line = buffer.slice(0, end);
          buffer = buffer.slice(end + (buffer[end] === "\r" && buffer[end + 1] === "\n" ? 2 : 1));
          if (!line) dispatch();
          else if (line.startsWith("data:")) dataLines.push(line.slice(5).replace(/^ /, ""));
          // 注释心跳、event/id/retry 等字段不会作为模型内容发送。
          if (done) return;
        }
      };
      while (!done) {
        const chunk = await readText(); buffer += chunk.text; parseLines(chunk.done);
        if (chunk.done) break;
      }
      if (!done || !finished || !content.trim()) return invalid();
      if (body.json) { try { JSON.parse(content); } catch { return invalid(); } }
      return { ok: true, content, model, usage };
    })()]);
  } catch (error) {
    if (clientSignal.aborted) throw new Error("cancelled");
    if (timedOut) return failure(config.provider, "timeout");
    if (error && typeof error === "object" && "ok" in error && error.ok === false) return error as Failure;
    return failure(config.provider, "network");
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    clientSignal.removeEventListener("abort", cancel);
    upstream.abort();
    if (reader) {
      // 不等待服务端取消握手，防止故障连接拖延备用服务；取消也会终止未完成的 read。
      void reader.cancel().catch(() => {});
      reader.releaseLock();
    }
  }
}

function streamReply(req: Request, configs: ProviderConfig[], body: RequestBody, ai_status: AIStatus) {
  const client = new AbortController(), encoder = new TextEncoder();
  let closed = false;
  const disconnected = () => client.abort();
  req.signal.addEventListener("abort", disconnected, { once: true });
  if (req.signal.aborted) disconnected();
  const stream = new ReadableStream<Uint8Array>({
    start(controller) {
      const emit = (event: string, data: unknown) => {
        if (closed || client.signal.aborted) throw new Error("cancelled");
        controller.enqueue(encoder.encode(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`));
      };
      const run = async () => {
        try {
          let lastFailure: Failure | null = null;
          const available = configs.filter(config => config.key);
          for (let i = 0; i < available.length; i++) {
            if (client.signal.aborted) return;
            const config = available[i];
            emit("status", { provider: config.provider, model: body.vision ? config.visionModel : config.model,
              fallback: config.provider === "deepseek", ai_status });
            let sentContent = false;
            const result = await generateStream(config, body, client.signal, text => { emit("delta", { text }); sentContent = true; });
            if (client.signal.aborted) return;
            if (result.ok) {
              ai_status[config.provider] = { state: "available", reason: null, message: "本次请求成功", checked_at: now() };
              emit("done", { content: result.content, model: result.model, usage: result.usage,
                provider: config.provider, fallback: config.provider === "deepseek", ai_status });
              return;
            }
            lastFailure = result; ai_status[config.provider] = failureStatus(result);
            if (!result.retry) { emit("error", { error: result.message, ai_status, provider: null }); return; }
            if (sentContent && i + 1 < available.length) emit("reset", { message: "主服务生成中断，正在使用备用服务重新生成", ai_status });
          }
          emit("error", { error: lastFailure ? configs.map(config => ai_status[config.provider].message).join("；")
            : "尚未配置 AI 服务，请添加 OPENAI_API_KEY 或 DEEPSEEK_API_KEY", ai_status, provider: null });
        } catch {
          if (!closed && !client.signal.aborted) emit("error", { error: "生成连接中断，请重试", ai_status, provider: null });
        } finally {
          req.signal.removeEventListener("abort", disconnected);
          if (!closed) { closed = true; controller.close(); }
        }
      };
      void run();
    },
    cancel() { closed = true; client.abort(); req.signal.removeEventListener("abort", disconnected); },
  });
  return new Response(stream, { headers: { ...CORS, "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-cache, no-transform", "X-Accel-Buffering": "no" } });
}

Deno.serve(async (req) => {
  const unchecked: AIStatus = { openai: unknownStatus(), deepseek: unknownStatus() };
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  if (req.method !== "POST") return fail(405, "只支持 POST", unchecked);
  const auth = req.headers.get("Authorization") || "";
  const apikey = req.headers.get("apikey") || env("SUPABASE_ANON_KEY");
  if (!auth.startsWith("Bearer ") || !auth.slice(7).trim()) return fail(401, "请先登录", unchecked);
  try {
    const sb = createClient(env("SUPABASE_URL"), apikey, {
      global: { headers: { Authorization: auth } }, auth: { persistSession: false },
    });
    const { data, error } = await sb.auth.getUser(auth.slice(7));
    if (error || !data?.user) return fail(401, "登录已过期，请重新登录", unchecked);
    const { data: isMember, error: memberError } = await sb.rpc("is_sales_member");
    if (memberError) return fail(500, "成员权限检查失败，请确认数据库已初始化后重试", unchecked);
    if (!isMember) return fail(403, "你的账号还没开通话术复盘系统", unchecked);
  } catch {
    return fail(503, "登录和成员权限检查暂时失败，请稍后重试", unchecked);
  }

  const configs = configuration();
  const ai_status = initialStatus(configs);
  if (Number(req.headers.get("content-length") || 0) > MAX_BODY) return fail(413, "内容太大，请减少文字或图片后重试", ai_status);
  let body: RequestBody;
  try {
    body = await req.json();
    if (!body || typeof body !== "object" || Array.isArray(body)) throw new Error("invalid");
  } catch { return fail(400, "请求格式不对", ai_status); }

  if (body.action === "status") {
    await Promise.all(configs.map(async config => { ai_status[config.provider] = await probe(config); }));
    return reply(200, { ai_status });
  }
  if (!Array.isArray(body.messages) || !body.messages.length) return fail(400, "缺少 messages", ai_status);
  if (body.stream === true) return streamReply(req, configs, body, ai_status);

  let lastFailure: Failure | null = null;
  for (const config of configs) {
    if (!config.key) continue;
    const result = await generate(config, body);
    if (result.ok) {
      ai_status[config.provider] = { state: "available", reason: null, message: "本次请求成功", checked_at: now() };
      return reply(200, { content: result.content, model: result.model, usage: result.usage,
        provider: config.provider, fallback: config.provider === "deepseek", ai_status });
    }
    lastFailure = result;
    ai_status[config.provider] = failureStatus(result);
    if (!result.retry) return fail(result.httpStatus, result.message, ai_status);
  }
  if (!lastFailure) return fail(503, "尚未配置 AI 服务，请添加 OPENAI_API_KEY 或 DEEPSEEK_API_KEY", ai_status);
  const message = configs.map(config => ai_status[config.provider].message).join("；");
  return fail(lastFailure.httpStatus, message, ai_status);
});
