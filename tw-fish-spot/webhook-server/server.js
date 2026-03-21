const express = require('express');
const axios = require('axios');
require('dotenv').config();

const app = express();
app.use(express.json());

const LINE_CHANNEL_ACCESS_TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN;
const WEBSITE_URL = 'https://cfw2214.github.io/tw-fish-spot';

let spotsData = [];

async function loadSpotsData() {
  try {
    const response = await axios.get(`${WEBSITE_URL}/spots.json`);
    spotsData = response.data;
    console.log(`✅ 已載入 ${spotsData.length} 個釣點`);
  } catch (error) {
    console.error('❌ 無法載入釣點資料:', error.message);
    spotsData = [];
  }
}

loadSpotsData();
setInterval(loadSpotsData, 60 * 60 * 1000);

function searchSpots(keyword) {
  if (!keyword) return [];
  const lowerKeyword = keyword.toLowerCase();
  return spotsData.filter(spot =>
    spot.name.includes(keyword) ||
    spot.area.includes(keyword) ||
    spot.type.includes(keyword) ||
    spot.note.includes(keyword)
  ).slice(0, 5);
}

function generateFlexMessage(spot) {
  const weatherUrl = `${WEBSITE_URL}?lat=${spot.lat}&lng=${spot.lng}&spot=${encodeURIComponent(spot.name)}`;

  return {
    type: 'bubble',
    body: {
      type: 'box',
      layout: 'vertical',
      contents: [
        {
          type: 'text',
          text: spot.name,
          weight: 'bold',
          size: 'lg',
          color: '#1DB446'
        },
        {
          type: 'box',
          layout: 'baseline',
          margin: 'md',
          contents: [
            {
              type: 'text',
              text: '📍 區域',
              color: '#aaaaaa',
              size: 'sm',
              flex: 1
            },
            {
              type: 'text',
              text: spot.area,
              wrap: true,
              color: '#666666',
              size: 'sm',
              flex: 5
            }
          ]
        },
        {
          type: 'box',
          layout: 'baseline',
          margin: 'md',
          contents: [
            {
              type: 'text',
              text: '🎣 類型',
              color: '#aaaaaa',
              size: 'sm',
              flex: 1
            },
            {
              type: 'text',
              text: spot.type,
              wrap: true,
              color: '#666666',
              size: 'sm',
              flex: 5
            }
          ]
        },
        ...(spot.note ? [{
          type: 'box',
          layout: 'baseline',
          margin: 'md',
          contents: [
            {
              type: 'text',
              text: '📝 備註',
              color: '#aaaaaa',
              size: 'sm',
              flex: 1
            },
            {
              type: 'text',
              text: spot.note,
              wrap: true,
              color: '#666666',
              size: 'sm',
              flex: 5
            }
          ]
        }] : []),
        {
          type: 'box',
          layout: 'vertical',
          margin: 'lg',
          spacing: 'sm',
          contents: [
            {
              type: 'button',
              style: 'link',
              height: 'sm',
              action: {
                type: 'uri',
                label: '🌤️ 查詢天氣詳情',
                uri: weatherUrl
              }
            },
            {
              type: 'button',
              style: 'link',
              height: 'sm',
              action: {
                type: 'uri',
                label: '📋 推薦其他釣點',
                uri: WEBSITE_URL
              }
            }
          ]
        }
      ]
    }
  };
}

app.post('/webhook', (req, res) => {
  const body = req.body;

  if (!body.events || !Array.isArray(body.events)) {
    return res.status(400).json({ error: 'Invalid request' });
  }

  body.events.forEach(async (event) => {
    if (event.type !== 'message' || event.message.type !== 'text') {
      return;
    }

    const userMessage = event.message.text;
    const replyToken = event.replyToken;

    try {
      const results = searchSpots(userMessage);

      let replyMessage;

      if (results.length === 0) {
        replyMessage = {
          type: 'text',
          text: `找不到「${userMessage}」相關的釣點。\n\n💡 試試看：\n• 輸入釣點名稱（如：野柳港）\n• 輸入區域名稱（如：金山萬里）\n• 輸入釣點類型（如：野場）\n\n📍 造訪官網查看完整釣點列表：${WEBSITE_URL}`
        };
      } else if (results.length === 1) {
        replyMessage = {
          type: 'flex',
          altText: results[0].name,
          contents: {
            type: 'carousel',
            contents: [generateFlexMessage(results[0])]
          }
        };
      } else {
        replyMessage = {
          type: 'flex',
          altText: `找到 ${results.length} 個釣點`,
          contents: {
            type: 'carousel',
            contents: results.map(spot => generateFlexMessage(spot))
          }
        };
      }

      await axios.post(
        'https://api.line.biz/v2/bot/message/reply',
        {
          replyToken: replyToken,
          messages: [replyMessage]
        },
        {
          headers: {
            Authorization: `Bearer ${LINE_CHANNEL_ACCESS_TOKEN}`,
            'Content-Type': 'application/json'
          }
        }
      );

      console.log(`✅ 已回覆: ${userMessage}`);
    } catch (error) {
      console.error('❌ 回覆失敗:', error.response?.data || error.message);
    }
  });

  res.status(200).json({ ok: true });
});

app.get('/health', (req, res) => {
  res.status(200).json({
    status: 'OK',
    spotsLoaded: spotsData.length,
    timestamp: new Date().toISOString()
  });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`🚀 Webhook 伺服器運行於 http://localhost:${PORT}`);
  console.log(`📡 Webhook 端點: http://localhost:${PORT}/webhook`);
  console.log(`💚 健康檢查: http://localhost:${PORT}/health`);
});
