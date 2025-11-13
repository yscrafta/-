import { Hono } from 'hono'
import { cors } from 'hono/cors'
import ExcelJS from 'exceljs'

const app = new Hono()

// CORS設定
app.use('/api/*', cors())

// CSV解析関数
function parseCSV(csvText: string): any[] {
  const lines = csvText.split('\n')
  if (lines.length < 2) return []
  
  const headers = lines[0].split(',')
  const data = []
  
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue
    
    const values = []
    let current = ''
    let inQuotes = false
    
    for (let j = 0; j < lines[i].length; j++) {
      const char = lines[i][j]
      if (char === '"') {
        inQuotes = !inQuotes
      } else if (char === ',' && !inQuotes) {
        values.push(current.trim())
        current = ''
      } else {
        current += char
      }
    }
    values.push(current.trim())
    
    const row: any = {}
    headers.forEach((header, index) => {
      row[header.trim()] = values[index] || ''
    })
    data.push(row)
  }
  
  return data
}

// 販路マッピング
const channelMapping: Record<string, string> = {
  'Airbnb': 'Airbnb',
  'Booking.com': 'Booking',
  '一休.com': '一休',
  'konjakuso': '自社サイト',
  '楽天トラベル': '楽天',
  'じゃらん': 'じゃらん'
}

// 言語判定
function getLanguage(nationality: string): string {
  if (nationality === 'Japan') return '日本語'
  if (nationality === 'United States of America') return '英語'
  if (['China', 'Taiwan', 'Hong Kong'].includes(nationality)) return '中国語'
  if (['Switzerland', 'Germany', 'Austria'].includes(nationality)) return 'ドイツ語'
  if (nationality === 'France') return 'フランス語'
  if (nationality === 'Spain') return 'スペイン語'
  if (nationality === 'Korea') return '韓国語'
  if (['Singapore', 'Malaysia'].includes(nationality)) return '英語'
  return ''
}

// 国名日本語変換
function getCountryJP(nationality: string): string {
  const countryMap: Record<string, string> = {
    'Japan': '日本',
    'United States of America': 'アメリカ',
    'Switzerland': 'スイス',
    'Germany': 'ドイツ',
    'China': '中国',
    'Taiwan': '台湾',
    'Hong Kong': '香港',
    'Korea': '韓国',
    'France': 'フランス',
    'Spain': 'スペイン',
    'Singapore': 'シンガポール',
    'Malaysia': 'マレーシア',
    'United Kingdom': 'イギリス',
    'Australia': 'オーストラリア'
  }
  return countryMap[nationality] || nationality
}

// API: CSV処理とExcel生成
app.post('/api/process', async (c) => {
  try {
    const formData = await c.req.formData()
    const file = formData.get('csv') as File
    const year = formData.get('year') as string
    const month = formData.get('month') as string
    
    if (!file) {
      return c.json({ error: 'CSVファイルがありません' }, 400)
    }
    
    // CSVを読み込む
    const csvText = await file.text()
    const bookings = parseCSV(csvText)
    
    // 指定月のデータをフィルタリング
    const filteredBookings = bookings.filter(row => {
      if (row['状態'] === 'システムキャンセル') return false
      
      const checkin = row['チェックイン']
      if (!checkin) return false
      
      try {
        const date = new Date(checkin)
        return date.getFullYear() === parseInt(year) && 
               date.getMonth() + 1 === parseInt(month)
      } catch {
        return false
      }
    })
    
    // Excelワークブックを作成
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet(`${year}年${month}月`)
    
    // ヘッダー行を追加（3行目にタイトル）
    worksheet.mergeCells('B3:C3')
    worksheet.getCell('B3').value = '売上計算書'
    worksheet.getCell('D3').value = new Date(parseInt(year), parseInt(month) - 1, 1)
    worksheet.getCell('G3').value = '今昔荘　弁天町　大阪ベイ'
    
    // 5行目にサブヘッダー
    const subHeaders = [
      '', '起算日', '決算日', '全日数', '', '', '', 'RevPER', '予約日数', '稼働率', 
      '平均客数', '月次売上\n（税込み）', 'ADR', '客単価', 'OTAサイト\n手数料', 
      'ADR\nOTA手数料\n差し引き後', 'OTAサイト\n手数料比率', '上代ADR\n（円/日）', 
      '清掃外注/リネン費', '清掃外注/リネン費'
    ]
    worksheet.getRow(5).values = subHeaders
    
    // 8行目にデータヘッダー
    const dataHeaders = [
      '', '言語', '国籍', '販路', 'ゲスト名', '予約日', 'C/I', 'C/O', '滞在日数', 
      '予約間隔', '人数', '支払金額', 'ADR\n（円/日）', '客単価\n（円/日人）', 
      'OTAサイト\n手数料', 'ADR\nOTA手数料\n差し引き後', '注釈', '上代ADR\n（円/日）', 
      '清掃外注/リネン費', '付加価値利益'
    ]
    worksheet.getRow(8).values = dataHeaders
    
    // データを追加（9行目から）
    let currentRow = 9
    filteredBookings.forEach(booking => {
      const site = booking['予約サイト']
      const channel = channelMapping[site] || site
      const language = getLanguage(booking['国籍'])
      const country = getCountryJP(booking['国籍'])
      
      const checkinDate = booking['チェックイン'] ? new Date(booking['チェックイン']) : null
      const checkoutDate = booking['チェックアウト'] ? new Date(booking['チェックアウト']) : null
      const bookingDate = booking['予約日'] ? new Date(booking['予約日']) : null
      
      const sales = parseFloat(booking['販売']) || 0
      const otaFee = parseFloat(booking['OTA サービス料']) || 0
      const guestCount = parseInt(booking['ゲスト数']) || 0
      
      const guestName = booking['ゲスト名']
      const channelId = booking['チャンネル予約ID']
      const guestWithId = `${guestName} (${channelId})`
      
      worksheet.getRow(currentRow).values = [
        '',
        language,
        country,
        channel,
        guestWithId,
        bookingDate,
        checkinDate,
        checkoutDate,
        { formula: `H${currentRow}-G${currentRow}` }, // 滞在日数
        { formula: `G${currentRow}-F${currentRow}` }, // 予約間隔
        guestCount,
        sales,
        { formula: `IF(I${currentRow}=0,"",L${currentRow}/I${currentRow})` }, // ADR
        { formula: `IF(I${currentRow}=0,"",M${currentRow}/K${currentRow})` }, // 客単価
        otaFee,
        { formula: `IF(I${currentRow}=0,"",(L${currentRow}-O${currentRow})/I${currentRow})` } // ADR OTA手数料差し引き後
      ]
      
      currentRow++
    })
    
    // Excelファイルを生成
    const buffer = await workbook.xlsx.writeBuffer()
    
    return new Response(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="売上計算書_${year}年${month}月.xlsx"`
      }
    })
    
  } catch (error: any) {
    console.error('Error:', error)
    return c.json({ error: error.message }, 500)
  }
})

// メインページ
app.get('/', (c) => {
  return c.html(`
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>売上管理表作成ツール</title>
        <script src="https://cdn.tailwindcss.com"></script>
        <link href="https://cdn.jsdelivr.net/npm/@fortawesome/fontawesome-free@6.4.0/css/all.min.css" rel="stylesheet">
    </head>
    <body class="bg-gradient-to-br from-blue-50 to-indigo-100 min-h-screen p-8">
        <div class="max-w-4xl mx-auto">
            <!-- ヘッダー -->
            <div class="bg-white rounded-lg shadow-lg p-8 mb-6">
                <h1 class="text-4xl font-bold text-gray-800 mb-2">
                    <i class="fas fa-file-excel text-green-600 mr-3"></i>
                    売上管理表作成ツール
                </h1>
                <p class="text-gray-600">CSVファイルから売上管理表（Excel）を自動生成します</p>
            </div>

            <!-- メインコンテンツ -->
            <div class="bg-white rounded-lg shadow-lg p-8">
                <form id="uploadForm" class="space-y-6">
                    <!-- ファイルアップロード -->
                    <div>
                        <label class="block text-sm font-semibold text-gray-700 mb-2">
                            <i class="fas fa-upload mr-2"></i>CSVファイル
                        </label>
                        <div class="relative">
                            <input type="file" id="csvFile" name="csv" accept=".csv" required
                                   class="block w-full text-sm text-gray-900 border border-gray-300 rounded-lg cursor-pointer bg-gray-50 focus:outline-none focus:border-indigo-500 p-2.5">
                        </div>
                        <p class="mt-1 text-sm text-gray-500">AirHostからエクスポートしたCSVファイルを選択してください</p>
                    </div>

                    <!-- 年月選択 -->
                    <div class="grid grid-cols-2 gap-4">
                        <div>
                            <label class="block text-sm font-semibold text-gray-700 mb-2">
                                <i class="fas fa-calendar mr-2"></i>年
                            </label>
                            <select id="year" name="year" required
                                    class="block w-full px-4 py-2.5 text-gray-900 border border-gray-300 rounded-lg bg-gray-50 focus:ring-indigo-500 focus:border-indigo-500">
                                <option value="2024">2024年</option>
                                <option value="2025" selected>2025年</option>
                                <option value="2026">2026年</option>
                            </select>
                        </div>
                        <div>
                            <label class="block text-sm font-semibold text-gray-700 mb-2">
                                <i class="fas fa-calendar-alt mr-2"></i>月
                            </label>
                            <select id="month" name="month" required
                                    class="block w-full px-4 py-2.5 text-gray-900 border border-gray-300 rounded-lg bg-gray-50 focus:ring-indigo-500 focus:border-indigo-500">
                                <option value="1">1月</option>
                                <option value="2">2月</option>
                                <option value="3">3月</option>
                                <option value="4">4月</option>
                                <option value="5">5月</option>
                                <option value="6">6月</option>
                                <option value="7">7月</option>
                                <option value="8">8月</option>
                                <option value="9">9月</option>
                                <option value="10">10月</option>
                                <option value="11" selected>11月</option>
                                <option value="12">12月</option>
                            </select>
                        </div>
                    </div>

                    <!-- 送信ボタン -->
                    <button type="submit" id="submitBtn"
                            class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-6 rounded-lg transition duration-200 shadow-md hover:shadow-lg">
                        <i class="fas fa-cog mr-2"></i>売上管理表を生成
                    </button>
                </form>

                <!-- ステータス表示 -->
                <div id="status" class="mt-6 hidden">
                    <div class="bg-blue-50 border border-blue-200 rounded-lg p-4">
                        <p class="text-blue-700 flex items-center">
                            <i class="fas fa-spinner fa-spin mr-2"></i>
                            <span id="statusText">処理中...</span>
                        </p>
                    </div>
                </div>

                <!-- エラー表示 -->
                <div id="error" class="mt-6 hidden">
                    <div class="bg-red-50 border border-red-200 rounded-lg p-4">
                        <p class="text-red-700 flex items-center">
                            <i class="fas fa-exclamation-circle mr-2"></i>
                            <span id="errorText"></span>
                        </p>
                    </div>
                </div>

                <!-- 成功表示 -->
                <div id="success" class="mt-6 hidden">
                    <div class="bg-green-50 border border-green-200 rounded-lg p-4">
                        <p class="text-green-700 flex items-center">
                            <i class="fas fa-check-circle mr-2"></i>
                            <span id="successText">Excelファイルのダウンロードが開始されました！</span>
                        </p>
                    </div>
                </div>
            </div>

            <!-- 使い方 -->
            <div class="bg-white rounded-lg shadow-lg p-8 mt-6">
                <h2 class="text-2xl font-bold text-gray-800 mb-4">
                    <i class="fas fa-info-circle text-blue-600 mr-2"></i>使い方
                </h2>
                <ol class="list-decimal list-inside space-y-2 text-gray-700">
                    <li>AirHostから予約データをCSV形式でエクスポート</li>
                    <li>CSVファイルをアップロード</li>
                    <li>対象の年月を選択</li>
                    <li>「売上管理表を生成」ボタンをクリック</li>
                    <li>自動的にExcelファイルがダウンロードされます</li>
                </ol>
            </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/axios@1.6.0/dist/axios.min.js"></script>
        <script>
            document.getElementById('uploadForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const submitBtn = document.getElementById('submitBtn');
                const status = document.getElementById('status');
                const error = document.getElementById('error');
                const success = document.getElementById('success');
                
                // 表示をリセット
                status.classList.remove('hidden');
                error.classList.add('hidden');
                success.classList.add('hidden');
                submitBtn.disabled = true;
                
                try {
                    const formData = new FormData(e.target);
                    
                    const response = await axios.post('/api/process', formData, {
                        responseType: 'blob',
                        headers: {
                            'Content-Type': 'multipart/form-data'
                        }
                    });
                    
                    // ダウンロード
                    const url = window.URL.createObjectURL(new Blob([response.data]));
                    const link = document.createElement('a');
                    link.href = url;
                    
                    const year = formData.get('year');
                    const month = formData.get('month');
                    link.setAttribute('download', \`売上計算書_\${year}年\${month}月.xlsx\`);
                    
                    document.body.appendChild(link);
                    link.click();
                    link.remove();
                    
                    status.classList.add('hidden');
                    success.classList.remove('hidden');
                    
                } catch (err) {
                    console.error(err);
                    status.classList.add('hidden');
                    error.classList.remove('hidden');
                    document.getElementById('errorText').textContent = 
                        'エラーが発生しました: ' + (err.response?.data?.error || err.message);
                } finally {
                    submitBtn.disabled = false;
                }
            });
        </script>
    </body>
    </html>
  `)
})

export default app
