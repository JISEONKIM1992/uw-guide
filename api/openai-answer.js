function isAscii(value) {
  return /^[\x20-\x7E]*$/.test(value);
}

function buildSystemPrompt() {
  return [
    "너는 보험사 신계약지원 Agent의 실답변 모듈이다.",
    "사용자는 저장된 상품의 가이드라인과 RD 반영 상태를 비교해 질문한다.",
    "반드시 한국어로 답하고, 제공된 구조화 데이터만 근거로 답한다.",
    "불확실하면 불확실하다고 말하고, 추측은 하지 않는다.",
    "출력은 특약 기준으로 정리하고, 답변 시작 부분에 반드시 '특약: [특약코드] [특약명]'을 표기한다.",
    "답변 본문은 해당 특약 1개만 기준으로 설명하며, 다른 특약과 혼동하지 않는다.",
    "단, 질문이 '정리해줘', '변경되는 특약', '한도 변경 특약'처럼 목록 요청이면 product.changedItems 또는 product.relatedItems 중 limit 변경이 있는 항목을 모두 모아 특약별로 나열한다.",
    "목록 요청의 경우 각 항목마다 특약코드+특약명, 일반/건강 단일건 변화, 간편 단일건 변화를 함께 보여준다.",
    "출력은 간결하되 실무자가 바로 볼 수 있도록 변경 포인트와 판단 근거를 분리해 설명한다.",
    "질문이 변경 여부 확인이면 현재 상태, 차이점, 검토 필요 여부를 우선 답한다.",
    "한도 판단은 product.comparisonSummary와 product.limits의 generalAsIs/generalToBe, simpleAsIs/simpleToBe만 사용한다.",
    "product.reference.individualSum은 참고값이며, 사용자가 별도로 묻지 않는 한 최종 한도로 사용하지 않는다.",
    "product.comparisonSummary가 있으면 그것을 가장 우선 근거로 읽고, raw limits는 보조 근거로만 사용한다.",
    "상품명과 보험코드를 혼동하지 말고, 질문 대상 상품 1개만 답한다.",
    "가능하면 답변 형식을 '특약 정보 / 판단 / 근거 / 후속 확인 포인트' 순서로 정리한다.",
    "응답 마지막에는 필요 시 후속 확인 포인트를 1~3개 정도 제안한다.",
  ].join("\n");
}

function buildUserPayload(body) {
  const product = body?.product || {};
  const products = Array.isArray(body?.products) ? body.products.slice(0, 20) : [];
  const uploads = Array.isArray(body?.uploads) ? body.uploads.slice(0, 10) : [];
  return JSON.stringify(
    {
      question: String(body?.question || "").trim(),
      questionMode: body?.questionMode || "",
      answerFocus: body?.answerFocus || "",
      ignoreFields: Array.isArray(body?.ignoreFields) ? body.ignoreFields : [],
      selectedProduct: body?.selectedProduct || null,
      product,
      relatedItems: Array.isArray(body?.relatedItems) ? body.relatedItems.slice(0, 50) : [],
      changedItems: Array.isArray(body?.changedItems) ? body.changedItems.slice(0, 50) : [],
      products,
      uploads,
    },
    null,
    2,
  );
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(204).end();
  }

  const API_KEY = String(process.env.BIZROUTER_API_KEY || "").trim().replace(/^['"]|['"]$/g, "");
  const BASE_URL = String(process.env.BIZROUTER_BASE_URL || "https://api.openai.com/v1").replace(/\/+$/, "");
  const MODEL = String(process.env.BIZROUTER_MODEL || "gpt-5.4-mini");

  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      message: "support-agent-openai-api is running",
      model: MODEL,
      hasApiKey: !!API_KEY,
    });
  }

  if (req.method !== "POST") {
    return res.status(404).json({ ok: false, error: "Not found" });
  }

  if (!API_KEY) {
    return res.status(500).json({ ok: false, error: "BIZROUTER_API_KEY가 설정되지 않았습니다." });
  }
  if (!isAscii(API_KEY)) {
    return res.status(500).json({
      ok: false,
      error: "BIZROUTER_API_KEY에 ASCII가 아닌 문자가 포함되어 있습니다. 실제 OpenAI API 키만 넣어 주세요.",
    });
  }
  if (!isAscii(BASE_URL) || !isAscii(MODEL)) {
    return res.status(500).json({
      ok: false,
      error: "BIZROUTER_BASE_URL 또는 BIZROUTER_MODEL 값에 ASCII가 아닌 문자가 포함되어 있습니다.",
    });
  }

  const question = String(req.body?.question || "").trim();
  if (!question) {
    return res.status(400).json({ ok: false, error: "질문이 비어 있습니다." });
  }

  const payload = {
    model: MODEL,
    messages: [
      { role: "system", content: buildSystemPrompt() },
      { role: "user", content: buildUserPayload(req.body) },
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
      return res.status(upstream.status).json({ ok: false, error: message });
    }

    const answer = data?.choices?.[0]?.message?.content || "";
    return res.status(200).json({
      ok: true,
      answer,
      model: MODEL,
      usage: data?.usage || null,
    });
  } catch (error) {
    return res.status(500).json({ ok: false, error: error.message });
  }
}
