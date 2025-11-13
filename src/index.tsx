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

// 施設名を取得
app.post('/api/facilities', async (c) => {
  try {
    const formData = await c.req.formData()
    const file = formData.get('csv') as File
    
    if (!file) {
      return c.json({ error: 'CSVファイルがありません' }, 400)
    }
    
    // CSVを読み込む
    const csvText = await file.text()
    const bookings = parseCSV(csvText)
    
    // 施設名を抽出
    const facilities = new Set<string>()
    bookings.forEach(row => {
      const facility = row['物件名']
      if (facility && facility.trim()) {
        facilities.add(facility.trim())
      }
    })
    
    return c.json({ facilities: Array.from(facilities).sort() })
    
  } catch (error: any) {
    console.error('Error:', error)
    return c.json({ error: error.message }, 500)
  }
})

// API: CSV処理とExcel生成
app.post('/api/process', async (c) => {
  try {
    const formData = await c.req.formData()
    const file = formData.get('csv') as File
    const year = formData.get('year') as string
    const month = formData.get('month') as string
    const facility = formData.get('facility') as string
    
    if (!file) {
      return c.json({ error: 'CSVファイルがありません' }, 400)
    }
    
    // CSVを読み込む
    const csvText = await file.text()
    const bookings = parseCSV(csvText)
    
    // 指定月と施設のデータをフィルタリング
    const filteredBookings = bookings.filter(row => {
      if (row['状態'] === 'システムキャンセル') return false
      
      // 施設フィルター（"全施設"の場合はフィルタリングしない）
      if (facility && facility !== '全施設') {
        const rowFacility = row['物件名']?.trim()
        if (rowFacility !== facility) return false
      }
      
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
    
    // 施設名を短縮（シート名が31文字を超えないように）
    let sheetName = `${year}年${month}月`
    if (facility && facility !== '全施設') {
      const shortFacility = facility.length > 20 ? facility.substring(0, 20) + '...' : facility
      sheetName = `${year}年${month}月_${shortFacility}`.substring(0, 31)
    }
    
    // Excelワークブックを作成
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet(sheetName)
    
    // 列幅を設定
    worksheet.columns = [
      { width: 2.57 },  // A
      { width: 10.0 },  // B
      { width: 11.0 },  // C
      { width: 12.57 }, // D
      { width: 23.71 }, // E
      { width: 10.0 },  // F
      { width: 10.14 }, // G
      { width: 10.29 }, // H
      { width: 7.43 },  // I
      { width: 7.57 },  // J
      { width: 7.43 },  // K
      { width: 10.57 }, // L
      { width: 10.0 },  // M
      { width: 10.86 }, // N
      { width: 10.29 }, // O
      { width: 10.29 }, // P
      { width: 10.0 },  // Q
      { width: 10.0 },  // R
      { width: 10.0 },  // S
      { width: 10.0 }   // T
    ]
    
    // 3行目のタイトル行
    worksheet.mergeCells('B3:C3')
    const titleCell = worksheet.getCell('B3')
    titleCell.value = '売上計算書'
    titleCell.font = { size: 14, bold: true }
    titleCell.alignment = { vertical: 'middle' }
    
    const dateCell = worksheet.getCell('D3')
    dateCell.value = new Date(parseInt(year), parseInt(month) - 1, 1)
    dateCell.numFmt = 'yyyy-mm-dd'
    dateCell.font = { size: 12, bold: true }
    dateCell.alignment = { horizontal: 'right', vertical: 'middle' }
    
    const propertyCell = worksheet.getCell('G3')
    propertyCell.value = facility && facility !== '全施設' ? facility : '今昔荘（全施設）'
    propertyCell.font = { size: 12, bold: true }
    propertyCell.alignment = { vertical: 'middle' }
    
    // 5行目のサブヘッダー（紫色の背景）
    const purpleFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFE5DFEC' } }
    const subHeaders = [
      '', '起算日', '決算日', '全日数', '', '', '', 'RevPER', '予約日数', '稼働率', 
      '平均客数', '月次売上\n（税込み）', 'ADR', '客単価', 'OTAサイト\n手数料', 
      'ADR\nOTA手数料\n差し引き後', 'OTAサイト\n手数料比率', '上代ADR\n（円/日）', 
      '清掃外注/リネン費', '清掃外注/リネン費'
    ]
    
    subHeaders.forEach((header, index) => {
      if (header) {
        const cell = worksheet.getCell(5, index + 1)
        cell.value = header
        cell.fill = purpleFill
        cell.font = { size: 11 }
        cell.alignment = { vertical: 'top', wrapText: true }
      }
    })
    
    // 6行目のサマリー行（オレンジ色の背景）
    const orangeFill = { type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FFFDE9D9' } }
    const lastRow = 9 + filteredBookings.length - 1
    
    // 起算日
    const startDateCell = worksheet.getCell('B6')
    startDateCell.value = new Date(parseInt(year), parseInt(month) - 1, 1)
    startDateCell.numFmt = 'yyyy-mm-dd'
    startDateCell.alignment = { horizontal: 'right', vertical: 'middle' }
    
    // 決算日
    const endDateCell = worksheet.getCell('C6')
    endDateCell.value = new Date(parseInt(year), parseInt(month), 0)
    endDateCell.numFmt = 'yyyy-mm-dd'
    endDateCell.alignment = { horizontal: 'right', vertical: 'middle' }
    
    // 全日数
    const daysCell = worksheet.getCell('D6')
    daysCell.value = { formula: '=C6-B6+1' }
    daysCell.alignment = { vertical: 'middle' }
    
    // RevPER
    const revperCell = worksheet.getCell('H6')
    revperCell.value = { formula: '=J6*M6' }
    revperCell.fill = orangeFill
    revperCell.alignment = { vertical: 'middle' }
    
    // 予約日数
    const bookingDaysCell = worksheet.getCell('I6')
    bookingDaysCell.value = { formula: `=SUM(I9:I${lastRow})` }
    bookingDaysCell.fill = orangeFill
    bookingDaysCell.alignment = { vertical: 'middle' }
    
    // 稼働率
    const occupancyCell = worksheet.getCell('J6')
    occupancyCell.value = { formula: '=I6/D6' }
    occupancyCell.fill = orangeFill
    occupancyCell.numFmt = '0%'
    occupancyCell.alignment = { vertical: 'middle' }
    
    // 平均客数
    const avgGuestsCell = worksheet.getCell('K6')
    avgGuestsCell.value = { formula: `=SUMPRODUCT(K9:K${lastRow},I9:I${lastRow})/I6` }
    avgGuestsCell.fill = orangeFill
    avgGuestsCell.numFmt = '0.0'
    avgGuestsCell.alignment = { vertical: 'middle' }
    
    // 月次売上
    const totalSalesCell = worksheet.getCell('L6')
    totalSalesCell.value = { formula: `=SUM(L9:L${lastRow})` }
    totalSalesCell.fill = orangeFill
    totalSalesCell.numFmt = '#,##0'
    totalSalesCell.alignment = { vertical: 'middle' }
    
    // ADR
    const adrCell = worksheet.getCell('M6')
    adrCell.value = { formula: `=SUM(L9:L${lastRow})/SUM(I9:I${lastRow})` }
    adrCell.fill = orangeFill
    adrCell.numFmt = '#,##0'
    adrCell.alignment = { vertical: 'middle' }
    
    // 客単価
    const guestPriceCell = worksheet.getCell('N6')
    guestPriceCell.value = { formula: '=M6/K6' }
    guestPriceCell.fill = orangeFill
    guestPriceCell.numFmt = '#,##0'
    guestPriceCell.alignment = { vertical: 'middle' }
    
    // OTAサイト手数料
    const otaFeeCell = worksheet.getCell('O6')
    otaFeeCell.value = { formula: `=SUM(O9:O${lastRow})` }
    otaFeeCell.fill = orangeFill
    otaFeeCell.numFmt = '#,##0'
    otaFeeCell.alignment = { vertical: 'middle' }
    
    // ADR OTA手数料差し引き後
    const adrAfterFeeCell = worksheet.getCell('P6')
    adrAfterFeeCell.value = { formula: '=(L6-O6)/I6' }
    adrAfterFeeCell.fill = orangeFill
    adrAfterFeeCell.numFmt = '#,##0'
    adrAfterFeeCell.alignment = { vertical: 'bottom' }
    
    // OTAサイト手数料比率
    const feeRatioCell = worksheet.getCell('Q6')
    feeRatioCell.value = { formula: '=O6/L6' }
    feeRatioCell.fill = orangeFill
    feeRatioCell.numFmt = '0.0%'
    feeRatioCell.alignment = { vertical: 'middle' }
    
    // 上代ADR
    const retailAdrCell = worksheet.getCell('R6')
    retailAdrCell.value = { formula: `=AVERAGE(R9:R${lastRow})` }
    retailAdrCell.fill = orangeFill
    retailAdrCell.numFmt = '#,##0'
    retailAdrCell.alignment = { vertical: 'middle' }
    
    // 清掃外注/リネン費（S列とT列）
    const cleaningCell1 = worksheet.getCell('S6')
    cleaningCell1.value = { formula: `=SUM(S9:S${lastRow})` }
    cleaningCell1.numFmt = '#,##0'
    cleaningCell1.alignment = { vertical: 'middle' }
    
    const cleaningCell2 = worksheet.getCell('T6')
    cleaningCell2.value = { formula: `=SUM(T9:T${lastRow})` }
    cleaningCell2.numFmt = '#,##0'
    cleaningCell2.alignment = { vertical: 'middle' }
    
    // 7行目（追加の計算式）
    const avgDaysCell = worksheet.getCell('I7')
    avgDaysCell.value = { formula: `=SUMIF(I9:I${lastRow},"<>0")/COUNTIF(I9:I${lastRow},"<>0")` }
    avgDaysCell.numFmt = '0.0'
    avgDaysCell.alignment = { vertical: 'middle' }
    
    // 8行目のデータヘッダー（紫色の背景）
    const dataHeaders = [
      '', '言語', '国籍', '販路', 'ゲスト名', '予約日', 'C/I', 'C/O', '滞在日数', 
      '予約間隔', '人数', '支払金額', 'ADR\n（円/日）', '客単価\n（円/日人）', 
      'OTAサイト\n手数料', 'ADR\nOTA手数料\n差し引き後', '注釈', '上代ADR\n（円/日）', 
      '清掃外注/リネン費', '付加価値利益'
    ]
    
    dataHeaders.forEach((header, index) => {
      if (header) {
        const cell = worksheet.getCell(8, index + 1)
        cell.value = header
        cell.fill = purpleFill
        cell.font = { size: 11 }
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
      }
    })
    
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
      
      const row = worksheet.getRow(currentRow)
      
      // データを設定
      row.getCell(2).value = language // B列: 言語
      row.getCell(3).value = country  // C列: 国籍
      row.getCell(4).value = channel  // D列: 販路
      row.getCell(5).value = guestWithId // E列: ゲスト名
      
      // 予約日
      if (bookingDate) {
        row.getCell(6).value = bookingDate
        row.getCell(6).numFmt = 'yyyy-mm-dd'
      }
      
      // チェックイン
      if (checkinDate) {
        row.getCell(7).value = checkinDate
        row.getCell(7).numFmt = 'yyyy-mm-dd'
      }
      
      // チェックアウト
      if (checkoutDate) {
        row.getCell(8).value = checkoutDate
        row.getCell(8).numFmt = 'yyyy-mm-dd'
      }
      
      // 滞在日数（計算式）
      row.getCell(9).value = { formula: `=H${currentRow}-G${currentRow}` }
      
      // 予約間隔（計算式）
      row.getCell(10).value = { formula: `=G${currentRow}-F${currentRow}` }
      
      // 人数
      row.getCell(11).value = guestCount
      
      // 支払金額
      row.getCell(12).value = sales
      row.getCell(12).numFmt = '#,##0'
      
      // ADR（計算式）
      row.getCell(13).value = { formula: `=IF(I${currentRow}=0,"",L${currentRow}/I${currentRow})` }
      row.getCell(13).numFmt = '#,##0'
      
      // 客単価（計算式）
      row.getCell(14).value = { formula: `=IF(I${currentRow}=0,"",M${currentRow}/K${currentRow})` }
      row.getCell(14).numFmt = '#,##0'
      
      // OTAサイト手数料
      row.getCell(15).value = otaFee
      row.getCell(15).numFmt = '#,##0'
      
      // ADR OTA手数料差し引き後（計算式）
      row.getCell(16).value = { formula: `=IF(I${currentRow}=0,"",(L${currentRow}-O${currentRow})/I${currentRow})` }
      row.getCell(16).numFmt = '#,##0'
      
      currentRow++
    })
    
    // Excelファイルを生成
    const buffer = await workbook.xlsx.writeBuffer()
    
    // ファイル名を生成
    let filename = `売上計算書_${year}年${month}月`
    if (facility && facility !== '全施設') {
      // ファイル名に使用できない文字を置換
      const safeFacility = facility.replace(/[<>:"/\\|?*]/g, '_')
      filename += `_${safeFacility}`
    } else {
      filename += '_全施設'
    }
    filename += '.xlsx'
    
    return new Response(buffer, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(filename)}"`
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

                    <!-- 施設選択 -->
                    <div id="facilitySelectContainer" class="hidden">
                        <label class="block text-sm font-semibold text-gray-700 mb-2">
                            <i class="fas fa-building mr-2"></i>施設
                        </label>
                        <select id="facility" name="facility" required
                                class="block w-full px-4 py-2.5 text-gray-900 border border-gray-300 rounded-lg bg-gray-50 focus:ring-indigo-500 focus:border-indigo-500">
                            <option value="">施設を選択してください</option>
                            <option value="全施設">全施設（統合）</option>
                        </select>
                        <p class="mt-1 text-sm text-gray-500">CSVファイルから自動検出された施設が表示されます</p>
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
                    <li>AirHostから予約データをCSV形式でエクスポート（全施設データ）</li>
                    <li>CSVファイルをアップロード</li>
                    <li>施設を選択（個別施設または全施設統合）</li>
                    <li>対象の年月を選択</li>
                    <li>「売上管理表を生成」ボタンをクリック</li>
                    <li>自動的にExcelファイルがダウンロードされます</li>
                </ol>
                <div class="mt-4 p-4 bg-blue-50 rounded-lg">
                    <p class="text-sm text-blue-800">
                        <i class="fas fa-lightbulb mr-2"></i>
                        <strong>ヒント:</strong> 全施設のCSVをアップロードすると、施設ごとに個別の売上管理表を作成できます。
                    </p>
                </div>
            </div>
        </div>

        <script src="https://cdn.jsdelivr.net/npm/axios@1.6.0/dist/axios.min.js"></script>
        <script>
            let currentFile = null;
            
            // CSVファイル選択時に施設一覧を取得
            document.getElementById('csvFile').addEventListener('change', async (e) => {
                const file = e.target.files[0];
                if (!file) return;
                
                currentFile = file;
                const facilitySelectContainer = document.getElementById('facilitySelectContainer');
                const facilitySelect = document.getElementById('facility');
                const statusText = document.getElementById('statusText');
                const status = document.getElementById('status');
                
                // ステータス表示
                status.classList.remove('hidden');
                statusText.textContent = '施設情報を読み込んでいます...';
                
                try {
                    const formData = new FormData();
                    formData.append('csv', file);
                    
                    const response = await axios.post('/api/facilities', formData, {
                        headers: {
                            'Content-Type': 'multipart/form-data'
                        }
                    });
                    
                    const facilities = response.data.facilities;
                    
                    // 施設選択肢を更新
                    facilitySelect.innerHTML = '<option value="">施設を選択してください</option>';
                    facilitySelect.innerHTML += '<option value="全施設">全施設（統合）</option>';
                    
                    facilities.forEach(facility => {
                        const option = document.createElement('option');
                        option.value = facility;
                        option.textContent = facility;
                        facilitySelect.appendChild(option);
                    });
                    
                    // 施設選択を表示
                    facilitySelectContainer.classList.remove('hidden');
                    status.classList.add('hidden');
                    
                    // 施設数を表示
                    statusText.textContent = \`\${facilities.length}施設を検出しました\`;
                    
                } catch (err) {
                    console.error(err);
                    status.classList.add('hidden');
                    document.getElementById('error').classList.remove('hidden');
                    document.getElementById('errorText').textContent = 
                        '施設情報の読み込みに失敗しました: ' + (err.response?.data?.error || err.message);
                }
            });
            
            document.getElementById('uploadForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                
                const submitBtn = document.getElementById('submitBtn');
                const status = document.getElementById('status');
                const error = document.getElementById('error');
                const success = document.getElementById('success');
                const statusText = document.getElementById('statusText');
                
                // 表示をリセット
                status.classList.remove('hidden');
                error.classList.add('hidden');
                success.classList.add('hidden');
                submitBtn.disabled = true;
                statusText.textContent = '売上管理表を生成中...';
                
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
                    
                    // ファイル名はContent-Dispositionヘッダーから取得
                    const contentDisposition = response.headers['content-disposition'];
                    let filename = '売上計算書.xlsx';
                    if (contentDisposition) {
                        const filenameMatch = contentDisposition.match(/filename[^;=\\n]*=((['"]).*?\\2|[^;\\n]*)/);
                        if (filenameMatch && filenameMatch[1]) {
                            filename = decodeURIComponent(filenameMatch[1].replace(/['"]/g, ''));
                        }
                    }
                    
                    link.setAttribute('download', filename);
                    
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
