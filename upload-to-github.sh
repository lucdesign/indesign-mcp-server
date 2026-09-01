#!/bin/bash

# GitHub Upload Script
# Führe diese Befehle nacheinander aus, nachdem du das Repository auf GitHub erstellt hast

echo "🚀 InDesign MCP Server GitHub Upload"
echo "====================================="

# 1. Remote Origin hinzufügen (ERSETZE 'lucdesign' mit deinem GitHub Username!)
git remote add origin https://github.com/lucdesign/indesign-mcp-server.git

# 2. Branch als main setzen
git branch -M main

# 3. Auf GitHub pushen
git push -u origin main

echo "✅ Repository erfolgreich auf GitHub veröffentlicht!"
echo "🌟 Dein Repository ist jetzt live auf:"
echo "    https://github.com/lucdesign/indesign-mcp-server"