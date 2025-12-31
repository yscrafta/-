#!/bin/bash
# 売上管理表作成ツール - 再起動スクリプト

echo "=== 売上管理表作成ツール 再起動 ==="

# 1. PM2を完全停止
echo "PM2を停止中..."
pm2 delete all 2>/dev/null || true
pm2 kill 2>/dev/null || true

# 2. ポート3000をクリーンアップ
echo "ポート3000をクリーンアップ中..."
fuser -k 3000/tcp 2>/dev/null || true
sleep 2

# 3. distディレクトリの確認
if [ ! -d "/home/user/webapp/dist" ]; then
    echo "distディレクトリが見つかりません。ビルドを実行します..."
    cd /home/user/webapp && npm run build
fi

# 4. PM2で起動
echo "PM2でサービスを起動中..."
cd /home/user/webapp
pm2 start ecosystem.config.cjs

# 5. 起動待機
echo "起動を待機中..."
sleep 8

# 6. ステータス確認
echo "=== サービスステータス ==="
pm2 status

# 7. 接続テスト
echo ""
echo "=== 接続テスト ==="
if curl -s -I http://localhost:3000 | grep "200 OK" > /dev/null; then
    echo "✅ サービスは正常に起動しています"
    echo ""
    echo "アクセスURL: https://3000-ig9ldweph8q90uxshkn1h-00000000.sandbox.novita.ai"
else
    echo "❌ サービスの起動に失敗しました"
    echo "ログを確認してください: pm2 logs webapp --nostream"
fi
