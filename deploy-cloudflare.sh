#!/bin/bash
# Cloudflare Pages デプロイスクリプト

echo "=== 売上管理表作成ツール - Cloudflare Pagesデプロイ ==="

# 1. ビルド
echo "📦 ビルド中..."
cd /home/user/webapp
npm run build

if [ $? -ne 0 ]; then
    echo "❌ ビルドに失敗しました"
    exit 1
fi

echo "✅ ビルド完了"

# 2. デプロイ
echo "🚀 Cloudflare Pagesにデプロイ中..."
CLOUDFLARE_ACCOUNT_ID="ceefb3b24472663aa536faa0cab664aa" \
CLOUDFLARE_API_TOKEN="c88QlVJwvObG7tKNE5pe8ZlwXgKm60mWwTH7BFNc" \
npx wrangler pages deploy dist --project-name webapp --commit-dirty=true

if [ $? -ne 0 ]; then
    echo "❌ デプロイに失敗しました"
    exit 1
fi

echo ""
echo "✅ デプロイ完了！"
echo "🌐 公開URL: https://webapp-4kf.pages.dev"
echo ""
