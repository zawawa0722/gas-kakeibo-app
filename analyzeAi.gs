function analyzeHouseholdBudget(tyoubo, sisyutu) {
  // AIへのプロンプトを作成
  // オブジェクトをスペース区切りで配列として取得
  let tyouboArray = tyoubo.match(/[^:\s]+:\s*[^:\s]+/g);
  let sisyutuArray = sisyutu.match(/[^:\s]+:\s*[^:\s]+/g);

  // プロンプトの作成
  let prompt = `
  以下は家計簿のデータです。収支の傾向を分析し、節約のアドバイスを一言提供してください。

  【帳簿データ】
  項目,金額
  ${tyouboArray}
  【支出データ】
  カテゴリー,金額
  ${sisyutuArray}
  `;

  let payload = {
    model: "gpt-4",
    messages: [{ role: "user", content: prompt }],
    temperature: 0.7,
    max_tokens: 500,
  };

  let options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + API_OPENAI,
    },
    payload: JSON.stringify(payload),
  };

  let response = UrlFetchApp.fetch(
    "https://api.openai.com/v1/chat/completions",
    options,
  );
  let json = JSON.parse(response.getContentText());
  let aiComment = json.choices[0].message.content.trim();

  return aiComment;
}
