#!/bin/bash

echo "🔑 GitHub Token Setup und Push"
echo "=============================="
echo ""
echo "📝 Bitte gib dein GitHub Personal Access Token ein:"
echo "   (Das Token wird nicht sichtbar angezeigt)"
echo ""
read -s -p "Token: " TOKEN
echo ""
echo ""

if [ -z "$TOKEN" ]; then
    echo "❌ Kein Token eingegeben!"
    exit 1
fi

echo "🔗 Konfiguriere Remote-URL mit Token..."
git remote add origin https://lucdesign:$TOKEN@github.com/lucdesign/indesign-mcp-server.git

echo "📤 Pushe Repository zu GitHub..."
git push -u origin main

if [ $? -eq 0 ]; then
    echo ""
    echo "✅ ERFOLG! Repository ist live auf GitHub!"
    echo "🌟 https://github.com/lucdesign/indesign-mcp-server"
    echo ""
    echo "🧹 Entferne Token aus der Remote-URL für Sicherheit..."
    git remote remove origin
    git remote add origin https://github.com/lucdesign/indesign-mcp-server.git
    echo "✅ Token-basierte Remote-URL entfernt."
else
    echo ""
    echo "❌ Push fehlgeschlagen!"
    echo ""
    echo "🔍 Mögliche Ursachen:"
    echo "   • Token hat keine 'repo' Berechtigung"
    echo "   • Token ist abgelaufen oder ungültig"
    echo "   • Repository-Name stimmt nicht überein"
    echo ""
    echo "💡 Überprüfe dein Token auf GitHub:"
    echo "   Settings → Developer settings → Personal access tokens"
fi