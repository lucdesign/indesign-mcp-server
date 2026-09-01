#!/bin/bash

echo "🔑 GitHub Push mit Personal Access Token"
echo "======================================="

# Entferne die alte Remote URL
git remote remove origin

# Füge neue Remote URL mit Token hinzu
# ERSETZE 'DEIN-TOKEN-HIER' mit deinem echten Personal Access Token!
echo "📝 Gib deinen GitHub Personal Access Token ein:"
read -s TOKEN

git remote add origin https://lucdesign:$TOKEN@github.com/lucdesign/indesign-mcp-server.git

# Push durchführen
git push -u origin main

echo "✅ Repository erfolgreich hochgeladen!"
