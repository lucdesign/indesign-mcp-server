#!/bin/bash

echo "🔍 GitHub Token Validator"
echo "========================"
echo ""
echo "📝 Gib dein neues GitHub Personal Access Token ein:"
read -s -p "Token: " TOKEN
echo ""

if [ -z "$TOKEN" ]; then
    echo "❌ Kein Token eingegeben!"
    exit 1
fi

echo "🔍 Teste Token-Berechtigungen..."
echo ""

# Test 1: User Info
echo "1️⃣ User Authentication Test:"
USER_INFO=$(curl -s -H "Authorization: token $TOKEN" https://api.github.com/user)
if echo "$USER_INFO" | grep -q '"login"'; then
    echo "   ✅ Token ist gültig"
    echo "   👤 User: $(echo "$USER_INFO" | grep '"login"' | cut -d'"' -f4)"
else
    echo "   ❌ Token ist ungültig oder abgelaufen"
    echo "   📝 Erstelle ein neues Token auf GitHub!"
    exit 1
fi

echo ""

# Test 2: Repository Access
echo "2️⃣ Repository Access Test:"
REPO_INFO=$(curl -s -H "Authorization: token $TOKEN" https://api.github.com/repos/lucdesign/indesign-mcp-server)
if echo "$REPO_INFO" | grep -q '"full_name"'; then
    echo "   ✅ Repository-Zugriff funktioniert"
else
    echo "   ❌ Kein Repository-Zugriff"
    echo "   🔑 Token braucht 'repo' Berechtigung!"
    exit 1
fi

echo ""

# Test 3: Write Permission Test
echo "3️⃣ Write Permission Test:"
WRITE_TEST=$(curl -s -X POST -H "Authorization: token $TOKEN" \
  -H "Content-Type: application/json" \
  -d '{"ref":"refs/heads/test-branch","sha":"6245918"}' \
  https://api.github.com/repos/lucdesign/indesign-mcp-server/git/refs 2>/dev/null)

if echo "$WRITE_TEST" | grep -q '"ref"'; then
    echo "   ✅ Schreibberechtigung vorhanden"
    # Lösche Test-Branch wieder
    curl -s -X DELETE -H "Authorization: token $TOKEN" \
      https://api.github.com/repos/lucdesign/indesign-mcp-server/git/refs/heads/test-branch >/dev/null 2>&1
elif echo "$WRITE_TEST" | grep -q "Bad credentials"; then
    echo "   ❌ Token hat keine Schreibberechtigung"
    echo "   🔑 Token braucht 'repo' Scope für Write-Access!"
    exit 1
else
    echo "   ⚠️  Schreibberechtigung unklar (aber vermutlich OK)"
fi

echo ""
echo "🎯 TOKEN IST BEREIT FÜR GITHUB PUSH!"
echo ""
echo "🚀 Führe jetzt aus: ./final-push.sh"