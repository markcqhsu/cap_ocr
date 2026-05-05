#!/bin/bash
set -e

gcloud run deploy cap-ocr-api \
  --source . \
  --region asia-east1 \
  --project cap-ocr

echo "✅ 部署完成：https://cap-ocr-api-65578120182.asia-east1.run.app"
