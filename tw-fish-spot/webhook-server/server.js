require('dotenv').config();
const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());

// 配置
const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const WEBSITE_URL = 'https://cfw2214.github.io/tw-fish-spot/';
const SPOTS_JSON_URL = 'https://cfw2214.github.io/tw-fish-spot/spots.json';

let SPOTS = [];

// ==================
// 初始化：載入釣點數據
// ==================
async function initializeSpots() {
  try {
    console.log('🔄 正在載入釣點數據...');
    const response = await axios.get(SPOTS_JSON_URL);
    SPOTS = response.data;
    console.log(`✅ 已載入 ${SPOTS.length} 個釣點`);
  } catch (error) {
    console.error('❌ 載入釣點數據失敗:', error.message);
    SPOTS = [];
  }
}

// ==================
// 搜尋釣點函數
// ==================
function searchSpots(query) {
  if (!query || query.trim() === '') {
    return [];
  }

  const normalizedQuery = query.trim();
  
  // 精確匹配名稱
  const exactMatch = SPOTS.find(spot => spot.name === normalizedQuery);
  if (exactMatch) {
    return [exactMatch];
  }

  // 模糊搜尋（包含關鍵字）
  return SPOTS.filter(spot => 
    spot.name.includes(normalizedQuery) || 
    spot.area.includes(normalizedQuery)
  ).slice(0, 5); // 最多回傳 5 個結果
}

// ==================
// 生成查詢連結
// ==================
function generateQueryUrl(spotName, date, time = '08:00') {
  const params = new URLSearchParams({
    spot: spotName,
    date: date || new Date().toISOString().split('T')[0],
    time: time
  });
  return `${WEBSITE_URL}?${params.toString()}`;
}

// ==================
// 取得台灣目前日期
// ==================
function getTWDateString() {
  const now = new Date();
  const tw = new Date(now.toLocaleString('en-US', { timeZone: 'Asia/Taipei' }));
  const y = tw.getFullYear();
  const m = String(tw.getMonth() + 1).padStart(2, '0');
  const d = String(tw.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

// ==================
// 回覆用戶函數
// ==================
async function replyToUser(replyToken, messages) {
  try {
    await axios.post(
      'https://api.messaging.line.biz/v2/bot/message/reply',
      {
        replyToken: replyToken,
        messages: Array.isArray(messages) ? messages : [messages]
      },
      {
        headers: {
          'Authorization': `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
          'Content-Type': 'application/json'
        }
      }
    );
    console.log('✅ 訊息已回覆');
  } catch (error) {
    console.error('❌ 回覆失敗:', error.response?.data || error.message);
  }
}

// ==================
// Webhook 路由
// ==================
app.post('/webhook', async (req, res) => {
  try {
    const events = req.body.events || [];
    console.log(`📨 收到 ${events.length} 個事件`);

    for (const event of events) {
      // 只處理文字訊息
      if (event.type !== 'message' || event.message.type !== 'text') {
        continue;
      }

      const userMessage = event.message.text.trim();
      const replyToken = event.replyToken;

      console.log(`👤 用戶訊息: "${userMessage}"`);

      // 搜尋釣點
      const results = searchSpots(userMessage);

      if (results.length === 0) {
        // 找不到釣點
        const notFoundMessage = {
          type: 'text',
          text: `找不到 "${userMessage}" 這個釣點 😅\n\n可試試：\n• 野柳港\n• 基隆外木山\n• 東北角區\n• 淡水三芝\n• 金山萬里\n\n或輸入釣點名稱查詢 🎣`
        };
        await replyToUser(replyToken, notFoundMessage);
      } else if (results.length === 1) {
        // 找到一個釣點 → 顯示 Button Template
        const spot = results[0];
        const today = getTWDateString();
        const queryUrl = generateQueryUrl(spot.name, today, '08:00');

        const spotMessage = {
          type: 'template',
          altText: `查詢 ${spot.name}`,
          template: {
            type: 'buttons',
            title: spot.name,
            text: `📍 區域：${spot.area}\n🎣 類型：${spot.type}${spot.allowed ? '' : '\n⚠️ 禁釣區'}`,
            actions: [
              {
                type: 'uri',
                label: '🌤️ 查詢天氣詳情',
                uri: queryUrl
              },
              {
                type: 'message',
                label: '📋 推薦其他釣點',
                text: '推薦'
              }
            ]
          }
        };

        await replyToUser(replyToken, spotMessage);
      } else {
        // 找到多個釣點 → 用 Carousel 或文字列表
        const spotList = results
          .map((s, idx) => `${idx + 1}. ${s.name} (${s.area})`)
          .join('\n');

        const multipleMessage = {
          type: 'text',
          text: `找到 ${results.length} 個釣點：\n\n${spotList}\n\n請輸入完整釣點名稱查詢詳情 🎣`
        };

        await replyToUser(replyToken, multipleMessage);
      }
    }

    res.status(200).send('OK');
  } catch (error) {
    console.error('❌ Webhook 錯誤:', error);
    res.status(500).send('Internal Server Error');
  }
});

// ==================
// 健康檢查路由
// ==================
app.get('/health', (req, res) => {
  res.status(200).json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    spotsLoaded: SPOTS.length
  });
});

// ==================
// 啟動伺服器
// ==================
const PORT = process.env.PORT || 3000;

// 在啟動前先載入釣點
initializeSpots().then(() => {
  app.listen(PORT, () => {
    console.log(`\n🎣 釣點 Webhook 伺服器已啟動`);
    console.log(`📍 監聽埠位: ${PORT}`);
    console.log(`🌐 健康檢查: http://localhost:${PORT}/health`);
    console.log(`📨 Webhook 路由: http://localhost:${PORT}/webhook\n`);
  });
});

module.exports = app;
