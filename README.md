# Extrator de Odds - UltraVirtual para Excel

## 📦 Instalação e Uso (Windows)

### Método Automático (Recomendado)

1. **Baixe os arquivos:**
   - `extract_odds.py`
   - `run_extract_odds.bat`
   - `requirements_minimo.txt` (renomeie para `requirements.txt`)

2. **Execute o instalador:**
   - Duplo clique em `run_extract_odds.bat`
   - O script irá:
     - ✅ Verificar/Instalar Python
     - ✅ Criar ambiente virtual
     - ✅ Instalar dependências
     - ✅ Executar o extrator

3. **Pronto!**
   - Digite o nome do arquivo Excel quando solicitado
   - Aguarde a automação completar

---

## 🚀 O que o script faz?

1. Abre o navegador Chrome
2. Acessa ultravirtual.com.br
3. Faz login automaticamente
4. Navega até Bet365 → Placar FT → Odds
5. Captura todas as odds do grid
6. Exporta para Excel com formatação:
   - Verde: odds positivas
   - Vermelho: odds negativas
   - Estatísticas completas

---

## 📋 Requisitos

- **Windows 7 ou superior**
- **Google Chrome instalado** (navegador)
- **Conexão com internet**
- Python será instalado automaticamente se necessário

---

## ⚠️ Licença

**Válida até: 31/03/2026**

O software expira automaticamente após essa data.

---

## 🔧 Build do Projeto

### Para Desenvolvedores

**Build local (Windows):**
```bash
build.bat
```

**Build via GitHub Actions:**
- Push para branch `main` ou `master`
- Ou crie uma tag: `git tag v1.0.0 && git push --tags`
- O executável será gerado automaticamente

---

## 📁 Arquivos Gerados

- `ExtractOdds.exe` - Executável standalone
- `*.xlsx` - Arquivos Excel com as odds extraídas

---

## 📞 Suporte

Em caso de problemas:
1. Verifique se o Chrome está atualizado
2. Verifique sua conexão com internet
3. Entre em contato para suporte técnico
