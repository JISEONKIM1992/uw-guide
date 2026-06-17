import http from "node:http";

const PORT = Number(process.env.PORT || 8787);
const API_KEY = String(process.env.BIZROUTER_API_KEY || "").trim().replace(/^['"]|['"]$/g, "");
const BASE_URL = String(process.env.BIZROUTER_BASE_URL || "https://api.openai.com/v1").replace(/\/+$/, "");
const MODEL = String(process.env.BIZROUTER_MODEL || "gpt-5.4-mini");

function isAscii(value) {
  return /^[\x20-\x7E]*$/.test(value);
}

function send(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  });
  res.end(body);
}

function readJson(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (chunk) => {
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    });
    req.on("end", () => {
      if (!chunks.length) {
        resolve({});
        return;
      }
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString("utf8")));
      } catch (error) {
        reject(error);
      }
    });
    req.on("error", reject);
  });
}

function buildSystemPrompt() {
  return [
    "너는 보험사 신계약지원 Agent의 실답변 모듈이다.",
    "사용자는 저장된 상품의 가이드라인과 RD 반영 상태를 비교해 질문한다.",
    "반드시 한국어로 답하고, 제공된 구조화 데이터만 근거로 답한다.",
    "불확실하면 불확실하다고 말하고, 추측은 하지 않는다.",
    "출력은 간결하되 실무자가 바로 볼 수 있도록 변경 포인트와 판단 근거를 분리해 설명한다.",
    "질문이 변경 여부 확인이면 현재 상태, 차이점, 검토 필요 여부를 우선 답한다.",
    "한도 판단은 product.comparisonSummary와 product.limits의 generalAsIs/generalToBe, simpleAsIs/simpleToBe만 사용한다.",
    "product.reference.individualSum은 참고값이며, 사용자가 별도로 묻지 않는 한 최종 한도로 사용하지 않는다.",
    "product.comparisonSummary가 있으면 그것을 가장 우선 근거로 읽고, raw limits는 보조 근거로만 사용한다.",
    "상품명과 보험코드를 혼동하지 말고, 질문 대상 상품 1개만 답한다.",
    "응답 마지막에는 필요 시 후속 확인 포인트를 1~3개 정도 제안한다.",
  ].join("\n");
}

function buildUserPayload(body) {
  const product = body?.product || {};
  const products = Array.isArray(body?.products) ? body.products.slice(0, 20) : [];
  const uploads = Array.isArray(body?.uploads) ? body.uploads.slice(0, 10) : [];
  return JSON.stringify({
    question: String(body?.question || "").trim(),
    answerFocus: body?.answerFocus || "",
    ignoreFields: Array.isArray(body?.ignoreFields) ? body.ignoreFields : [],
    product,
    products,
    uploads,
  }, null, 2);
}

const server = http.createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    send(res, 204, {});
    return;
  }

  const url = new URL(req.url || "/", "http://127.0.0.1");
  if (req.method === "GET" && url.pathname === "/health") {
    send(res, 200, { ok: true, model: MODEL, baseUrl: BASE_URL });
    return;
  }
  if (req.method === "GET" && url.pathname === "/") {
    res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("support-agent-openai-proxy is running");
    return;
  }
  if (req.method !== "POST" || url.pathname !== "/api/openai-answer") {
    send(res, 404, { ok: false, error: "Not found" });
    return;
  }
  if (!API_KEY) {
    send(res, 500, { ok: false, error: "BIZROUTER_API_KEY가 설정되지 않았습니다." });
    return;
  }
  if (!isAscii(API_KEY)) {
    send(res, 500, {
      ok: false,
      error: "BIZROUTER_API_KEY에 ASCII가 아닌 문자가 포함되어 있습니다. 실제 OpenAI API 키만 넣어 주세요.",
    });
    return;
  }
  if (!isAscii(BASE_URL) || !isAscii(MODEL)) {
    send(res, 500, {
      ok: false,
      error: "BIZROUTER_BASE_URL 또는 BIZROUTER_MODEL 값에 ASCII가 아닌 문자가 포함되어 있습니다.",
    });
    return;
  }

  let body;
  try {
    body = await readJson(req);
  } catch (error) {
    send(res, 400, { ok: false, error: `JSON 파싱 실패: ${error.message}` });
    return;
  }

  const question = String(body?.question || "").trim();
  if (!question) {
    send(res, 400, { ok: false, error: "질문이 비어 있습니다." });
    return;
  }

  const payload = {
    model: MODEL,
    messages: [
      { role: "system", content: buildSystemPrompt() },
      { role: "user", content: buildUserPayload(body) },
    ],
    temperature: 0.2,
  };

  try {
    const upstream = await fetch(`${BASE_URL}/chat/completions`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });
    const rawText = await upstream.text();
    let data = null;
    try {
      data = JSON.parse(rawText);
    } catch {
      data = null;
    }

    if (!upstream.ok) {
      const message = data?.error?.message || rawText || `OpenAI 요청 실패 (${upstream.status})`;
      send(res, upstream.status, { ok: false, error: message });
      return;
    }

    const answer = data?.choices?.[0]?.message?.content || "";
    send(res, 200, {
      ok: true,
      answer,
      model: MODEL,
      usage: data?.usage || null,
    });
  } catch (error) {
    send(res, 500, { ok: false, error: error.message });
  }
});

server.listen(PORT, "127.0.0.1", () => {
  console.log(`support-agent-openai-proxy listening on http://127.0.0.1:${PORT}`);
  console.log(`model=${MODEL}`);
  console.log(`baseUrl=${BASE_URL}`);
  console.log(API_KEY ? "API key loaded" : "API key missing");
});
