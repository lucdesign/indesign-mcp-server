#!/bin/bash

echo "🚀 Finaler Push-Versuch mit korrigierter Git-Konfiguration"
echo "========================================================"
echo ""
echo "📝 Gib dein GitHub Personal Access Token ein:"
read -s -p "Token: " TOKEN
echo ""

if [ -z "$TOKEN" ]; then
    echo "❌ Kein Token eingegeben!"
    exit 1
fi

echo "🔗 Setze Remote mit Token..."
git remote remove origin 2>/dev/null
git remote add origin https://lucdesign:$TOKEN@github.com/lucdesign/indesign-mcp-server.git

echo "📤 Push zu GitHub..."
if git push -u origin main; then
    echo ""
    echo "🎉 ✅ ERFOLG! Repository ist live!"
    echo "🌟 https://github.com/lucdesign/indesign-mcp-server"
    echo ""
    echo "🧹 Entferne Token für Sicherheit..."
    git remote remove origin
    git remote add origin https://github.com/lucdesign/indesign-mcp-server.git
    echo "✅ Fertig!"
else
    echo ""
    echo "❌ Push ist fehlgeschlagen."
    echo "🔍 Führe './debug-push.sh' für detaillierte Informationen aus."
fi