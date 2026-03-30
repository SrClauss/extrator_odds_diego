# Extrator de Odds - UltraVirtual para Excel

Script automatizado para extrair odds do UltraVirtual (Bet365) e exportar para Excel com formatação.

---

## 📦 Instalação

### Requisitos
- Python 3.11+
- Google Chrome instalado

### Setup

```bash
# Clone o repositório
git clone https://github.com/SrClauss/extrator_odds_diego.git
cd extrator_odds_diego

# Instale as dependências
pip install beautifulsoup4 openpyxl selenium webdriver-manager

# Execute
python extract_odds.py
```

---

## 🚀 O que o script faz?

1. Abre o navegador Chrome automaticamente
2. Acessa ultravirtual.com.br
3. Faz login (credenciais embutidas)
4. Navega até Bet365 → Placar FT → Odds
5. **Exibe menu para escolher o mercado** (17 opções disponíveis)
6. Captura todas as odds do mercado selecionado
7. Exporta para Excel com formatação:
   - 🟢 Verde: odds positivas
   - 🔴 Vermelho: odds negativas
   - 📊 Estatísticas completas

### Mercados Disponíveis

- Resultado Final
- Over Gols
- Under Gols
- Para o Time Marcar Sim/Não
- Resultado Correto - FT
- Primeiro Marcador de Gol
- Resultado Correto - Grupo
- Resultado/Para Ambos Times Marcarem
- Dupla Hipótese
- Intervalo/Final do Jogo
- Total de Gols Exatos
- Intervalo - Resultado
- Resultado Correto - HT
- Margem de Vitória
- Time a Marcar
- Time - Gols
- Handicap - Resultado

---

## ⚠️ Licença

**Válida até: 31/03/2026**

O software expira automaticamente após essa data.

---

## 🔧 Build Executável (Windows)

### PyInstaller

```bash
pip install pyinstaller

pyinstaller --onefile --name "ExtractOdds" \
  --hidden-import=selenium \
  --hidden-import=webdriver_manager \
  --hidden-import=bs4 \
  --hidden-import=openpyxl \
  extract_odds.py
```

Executável gerado em: `dist/ExtractOdds.exe`

### Build via GitHub Actions

1. Vá para: https://github.com/SrClauss/extrator_odds_diego/actions
2. Clique em "Build Extrator de Odds"
3. Clique em "Run workflow"
4. Aguarde ~5-10 minutos
5. Baixe o artifact gerado

Ou crie uma tag de versão:
```bash
git tag v1.0.0
git push origin v1.0.0
```

---

## 📁 Saída

- Arquivo `.xlsx` com nome definido pelo usuário
- Colunas: Odd | Verde | Vermelho | Total
- Headers formatados
- Dados com cores por categoria

---

## 🔐 Recursos

- ✅ Automação completa via Selenium
- ✅ Login automático
- ✅ Captura inteligente de odds
- ✅ Agregação por valor
- ✅ Formatação Excel profissional
- ✅ Contagem de verdes/vermelhos
- ✅ Estatísticas resumidas
- ✅ Limitação temporal de uso

---

## 📞 Suporte

Em caso de problemas:
1. Verifique se o Chrome está atualizado
2. Verifique sua conexão com internet
3. Verifique a validade da licença

**Repositório:** https://github.com/SrClauss/extrator_odds_diego
