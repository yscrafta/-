#!/bin/bash
export CLOUDFLARE_API_TOKEN="ya68JCrz9r9Dua4tG15CKkIsXo9wVJSToph9X7Yu"
npx wrangler pages deploy dist --project-name webapp --branch main --commit-dirty=true
