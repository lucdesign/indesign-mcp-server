#!/bin/bash

echo "🔍 Detailliertes GitHub Push Debugging"
echo "====================================="
echo ""
echo "📝 Bitte gib dein GitHub Personal Access Token ein:"
read -s -p "Token: " TOKEN
echo ""

if [ -z "$TOKEN" ]; then
    echo "❌ Kein Token eingegeben!"
    exit 1
fi

echo "🔗 Konfiguriere Remote mit Token..."
git remote remove origin 2>/dev/null
git remote add origin https://lucdesign:$TOKEN@github.com/lucdesign/indesign-mcp-server.git

echo "📋 Git Status:"
git status

echo ""
echo "📋 Git Remote:"
git remote -v

echo ""
echo "📋 Git Log (lokale Commits):"
git log --oneline -3

echo ""
echo "🔍 Teste GitHub API mit deinem Token..."
curl -s -H "Authorization: token $TOKEN" https://api.github.com/user | head -5

echo ""
echo "📤 Versuche Push mit detaillierten Fehlermeldungen..."
git push -u origin main --verbose

echo ""
echo "🧹 Entferne Token aus URL..."
git remote remove origin
git remote add origin https://github.com/lucdesign/indesign-mcp-server.git