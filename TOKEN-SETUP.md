# 🔑 GitHub Personal Access Token - Korrekte Einstellungen

## Token erstellen:
1. GitHub.com → Settings → Developer settings → Personal access tokens → Tokens (classic)
2. "Generate new token (classic)" klicken

## ✅ WICHTIGE SCOPES (Berechtigungen) die du AKTIVIEREN musst:

### **Basis-Berechtigung (UNBEDINGT nötig):**
☑️ **repo** 
   - Full control of private repositories
   - ⚠️ Das ist der wichtigste Scope für Push-Operationen!

### **Zusätzliche Berechtigungen (empfohlen):**
☑️ **workflow**
   - Update GitHub Action workflows

### **Was du NICHT brauchst (aber nicht schadet):**
- admin:repo_hook
- admin:org  
- user
- etc.

## ⚠️ HÄUFIGE FEHLER:
❌ Nur "public_repo" ausgewählt (reicht nicht für private Repos)
❌ Nur "read:user" ausgewählt (keine Schreibberechtigung)
❌ Token ohne "repo" Scope

## ✅ MINIMAL SETUP:
Für einen einfachen Push brauchst du nur:
- ☑️ **repo** (Full control of private repositories)

Das war's! Mehr ist nicht nötig.

## 🔄 Nach dem Erstellen:
1. Token SOFORT kopieren (wird nur einmal angezeigt!)
2. Altes Token löschen
3. Neues Token in unserem Script verwenden