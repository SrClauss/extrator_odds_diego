# INSTRUÇÕES DE DEPLOY

## ✅ Código enviado para GitHub

**Repositório:** https://github.com/SrClauss/extrator_odds_diego

---

## 🚀 Como o CI/CD funciona

### Build Automático

Toda vez que você fizer um push para `master` ou criar uma **tag**, o GitHub Actions irá:

1. ✅ Instalar Python 3.11
2. ✅ Instalar todas as dependências
3. ✅ Ofuscar o código com PyArmor (protege contra engenharia reversa)
4. ✅ Compilar com PyInstaller (gera executável único)
5. ✅ Criar pacote ZIP com todos os arquivos
6. ✅ Disponibilizar para download

### Como acessar os builds

**Opção 1: Artifacts (pushes normais)**
- Vá para: https://github.com/SrClauss/extrator_odds_diego/actions
- Clique no workflow mais recente
- Baixe o artifact `ExtractOdds-Windows`

**Opção 2: Releases (com tags)**
```bash
git tag v1.0.0
git push origin v1.0.0
```
- O executável será anexado automaticamente à release
- Acesse: https://github.com/SrClauss/extrator_odds_diego/releases

---

## 📦 O que foi implementado

### 1. Limitação de Licença
- ✅ Código expira em **31/03/2026**
- ✅ Aviso 7 dias antes
- ✅ Bloqueio automático após a data

### 2. Ofuscação de Código
- ✅ PyArmor ofusca o código Python
- ✅ Dificulta engenharia reversa
- ✅ Protege credenciais e lógica

### 3. Build Automatizado
- ✅ GitHub Actions CI/CD
- ✅ PyInstaller gera executável único
- ✅ Sem necessidade de Python no cliente
- ✅ Inclui ChromeDriver automático

### 4. Scripts de Instalação
- ✅ `run_extract_odds.bat` - Instalador completo
- ✅ `run_quick.bat` - Execução rápida
- ✅ `build.bat` - Build local para testes

---

## 🔧 Para fazer alterações

1. **Clone o repositório:**
```bash
git clone https://github.com/SrClauss/extrator_odds_diego.git
cd extrator_odds_diego
```

2. **Faça suas alterações em `extract_odds.py`**

3. **Commit e push:**
```bash
git add .
git commit -m "Descrição da alteração"
git push
```

4. **Aguarde o build automático** (5-10 minutos)

5. **Baixe o executável dos Actions**

---

## 📝 Alterando a data de expiração

Edite em `extract_odds.py`:

```python
# Linha 12
DATA_EXPIRACAO = datetime(2026, 3, 31, 23, 59, 59)  # Altere esta data
```

Depois faça commit e push.

---

## ⚙️ Testando localmente (Windows)

```bash
# Instalar dependências
pip install beautifulsoup4 openpyxl selenium webdriver-manager pyinstaller pyarmor

# Build local
build.bat

# Executável estará em: dist/ExtractOdds.exe
```

---

## 📊 Status do Projeto

- ✅ Código completo e funcional
- ✅ Automação Selenium implementada
- ✅ Extração para Excel com formatação
- ✅ Limitação de tempo ativa
- ✅ CI/CD configurado
- ✅ Ofuscação de código ativa
- ✅ Documentação completa

---

## 🔐 Segurança

**⚠️ IMPORTANTE:**
- As credenciais estão no código (`diegosantosdeassis5@gmail.com` / `bomba2022`)
- O código é ofuscado, mas não é 100% seguro
- Para maior segurança, considere:
  - Mover credenciais para variáveis de ambiente
  - Usar sistema de autenticação externo
  - Criptografar credenciais

---

## 📞 Próximos Passos

1. ✅ Testar o executável gerado pelo CI/CD
2. ✅ Distribuir para o cliente
3. ✅ Renovar licença conforme necessário
4. ✅ Monitorar logs de uso

---

**Repositório:** https://github.com/SrClauss/extrator_odds_diego
**Data de Criação:** 30/03/2026
**Licença:** Válida até 31/03/2026
