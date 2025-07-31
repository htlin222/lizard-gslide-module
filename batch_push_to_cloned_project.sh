#!/bin/bash

# Save original .clasp.json
cp .clasp.json .clasp.json.bak

if [ ! -f cloned.txt ]; then
  echo "❌ cloned.txt not found!"
  exit 1
fi

# Function to update .clasp.json with new scriptId
update_clasp_json() {
  local new_id=$1
  # Use jq if available for cleaner JSON editing
  if command -v jq >/dev/null 2>&1; then
    jq --arg sid "$new_id" '.scriptId = $sid' .clasp.json >.clasp.tmp && mv .clasp.tmp .clasp.json
  else
    # fallback to sed (basic replacement, assumes compact JSON format)
    sed -i '' "s/\"scriptId\": *\"[^\"]*\"/\"scriptId\": \"$new_id\"/" .clasp.json
  fi
}

# Loop through each script ID
while read -r scriptId; do
  if [[ -z "$scriptId" ]]; then
    continue
  fi

  echo "🚀 Pushing to scriptId: $scriptId"
  update_clasp_json "$scriptId"
  clasp push --force
  echo "✅ Done pushing to $scriptId"
  echo "--------------------------"
done <cloned.txt

# Restore original .clasp.json
mv .clasp.json.bak .clasp.json
echo "🔄 Restored original .clasp.json"

clasp status
