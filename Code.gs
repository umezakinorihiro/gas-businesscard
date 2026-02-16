const SHEET_NAME = "名刺";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("名刺")
    .addItem("読み込み（サイドバー）", "showSidebar")
    .addToUi();

  initSheet_();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("名刺読み込み");
  SpreadsheetApp.getUi().showSidebar(html);
}

function initSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const header = [
    "登録日時", "氏名", "フリガナ", "会社", "部署", "役職",
    "Email", "電話", "携帯", "FAX", "URL", "郵便番号", "住所",
    "メモ", "名刺全文(raw_text)", "画像URL"
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
}

/**
 * Sidebarから呼ばれる：dataURL(base64)の画像をGeminiに投げて抽出 → シートに追記
 */
function saveCardAndExtract(dataUrl, filename, mimeType) {
  initSheet_();

  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("GEMINI_API_KEY");
  if (!apiKey) throw new Error("スクリプトプロパティに GEMINI_API_KEY を設定してください。");

  const base64 = (dataUrl || "").split(",")[1];
  if (!base64) throw new Error("画像データの取得に失敗しました。");

  // 任意：Driveに画像保存
  let imageUrl = "";
  const folderId = props.getProperty("CARD_FOLDER_ID");
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(base64),
        mimeType || "image/jpeg",
        filename || `card_${Date.now()}.jpg`
      );
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
    } catch (e) {
      // フォルダIDが無い/権限無い等でも抽出自体は続行
      imageUrl = "";
    }
  }

  const extracted = callGeminiForBusinessCard_(apiKey, base64, mimeType || "image/jpeg");

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  sh.appendRow([
    new Date(),
    extracted.name || "",
    extracted.name_kana || "",
    extracted.company || "",
    extracted.department || "",
    extracted.title || "",
    (extracted.email || "").toLowerCase(),
    extracted.phone || "",
    extracted.mobile || "",
    extracted.fax || "",
    (extracted.url || "").toLowerCase(),
    extracted.postal_code || "",
    extracted.address || "",
    extracted.notes || "",
    extracted.raw_text || "",
    imageUrl
  ]);

  return {
    ok: true,
    rowIndex: sh.getLastRow(),
    imageUrl,
    extracted
  };
}

function callGeminiForBusinessCard_(apiKey, base64, mimeType) {
  // Gemini API generateContent エンドポイント :contentReference[oaicite:2]{index=2}
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash-lite:generateContent?key=${apiKey}`;

  // JSON Mode（response_mime_type / response_schema）を使う :contentReference[oaicite:3]{index=3}
  const responseSchema = {
    type: "OBJECT",
    properties: {
      name: { type: "STRING" },
      name_kana: { type: "STRING" },
      company: { type: "STRING" },
      department: { type: "STRING" },
      title: { type: "STRING" },
      email: { type: "STRING" },
      phone: { type: "STRING" },
      mobile: { type: "STRING" },
      fax: { type: "STRING" },
      url: { type: "STRING" },
      postal_code: { type: "STRING" },
      address: { type: "STRING" },
      notes: { type: "STRING" },
      raw_text: { type: "STRING" }
    }
  };

  const payload = {
    contents: [{
      parts: [
        {
          text:
`あなたは名刺の情報抽出アシスタントです。
画像の名刺から情報を読み取り、日本語のJSONで返してください。
見つからない項目は空文字 "" にしてください。
住所は都道府県から。電話番号は可能ならハイフン形式に。
raw_text には名刺の全文（改行を含む）を入れてください。`
        },
        // 画像は inline_data で送れる :contentReference[oaicite:4]{index=4}
        { inline_data: { mime_type: mimeType, data: base64 } }
      ]
    }],
    generationConfig: {
      response_mime_type: "application/json",
      response_schema: responseSchema,
      temperature: 0.2
    }
  };

  const res = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Gemini API error ${code}: ${body}`);
  }

  const obj = JSON.parse(body);
  const text = (((obj.candidates || [])[0] || {}).content || {}).parts
    ?.map(p => p.text || "")
    .join("") || "";

  // JSON Modeでも「text」にJSONが入るケースがあるのでパース
  try {
    return JSON.parse(text);
  } catch (e) {
    // 念のための保険
    return { raw_text: text };
  }
}