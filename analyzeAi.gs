function analyzeHouseholdBudget(tyoubo, sisyutu) {
  if (typeof tyoubo !== "string" || !tyoubo.trim() ||
      typeof sisyutu !== "string" || !sisyutu.trim()) {
    throw new Error("AI分析に必要な月次収支または支出内訳がありません。");
  }

  const properties = PropertiesService.getScriptProperties();
  const model = (properties.getProperty("OPENAI_MODEL") || "").trim();
  const fallbackModel = (properties.getProperty("OPENAI_FALLBACK_MODEL") || "").trim();
  if (!model) throw new Error("スクリプトプロパティ OPENAI_MODEL を設定してください。");
  if (!API_OPENAI) throw new Error("スクリプトプロパティ OPENAI_API_KEY を設定してください。");

  const prompt = [
    "以下は対象月の収支とカテゴリ別支出の集計です。",
    "集計に記載された金額と項目に基づいて分析してください。",
    "前月との比較、個々の買い物、支払い手段など、集計から分からないことは推測しないでください。",
    "",
    "【月次収支】",
    tyoubo.trim(),
    "",
    "【支出内訳】",
    sisyutu.trim(),
    "",
    "LINE通知用に、日本語のプレーンテキストで次の2行だけを返してください。",
    "良かった点：収支や支出配分から評価できる点を1つ（50文字程度）",
    "改善アドバイス：見直し候補のカテゴリと次月に試せる具体策を1つ（50文字程度）"
  ].join("\n");

  let activeModel = model;
  let result = requestBudgetAnalysis_(activeModel, prompt);
  const apiError = result.body?.error;
  const modelUnavailable = result.status === 404 && (
    apiError?.code === "model_not_found" ||
    apiError?.code === "model_deprecated" ||
    /the model .* has been deprecated/i.test(apiError?.message || "")
  );

  if (modelUnavailable && fallbackModel && fallbackModel !== model) {
    console.warn(`OpenAIモデル ${model} を利用できないため、${fallbackModel} で再試行します。`);
    activeModel = fallbackModel;
    result = requestBudgetAnalysis_(activeModel, prompt);
  }

  if (result.status < 200 || result.status >= 300) {
    const code = result.body?.error?.code;
    const reason = code ? `、エラーコード ${code}` : "";
    throw new Error(`OpenAI APIの呼び出しに失敗しました（HTTP ${result.status}、モデル ${activeModel}${reason}）。`);
  }

  const comment = result.body?.choices?.[0]?.message?.content;
  if (typeof comment !== "string" || !comment.trim()) {
    throw new Error(`OpenAI APIから分析コメントを取得できませんでした（モデル ${activeModel}）。`);
  }
  return comment.trim();
}

function requestBudgetAnalysis_(model, prompt) {
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + API_OPENAI },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      model,
      messages: [{ role: "user", content: prompt }],
      max_completion_tokens: 500
    })
  });

  const status = response.getResponseCode();
  let body = null;
  try {
    body = JSON.parse(response.getContentText());
  } catch (error) {
    // JSON以外の応答でも呼び出し元にHTTPステータスを返す。
  }
  return { status, body };
}
