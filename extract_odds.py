#!/usr/bin/env python3
"""
Script de extração de odds para Excel usando Selenium
Automatiza login no UltraVirtual, navega até Bet365 e extrai odds para Excel.
"""

import re
import time
import subprocess
import sys
from datetime import datetime
from collections import defaultdict

# Verificação de licença - expira em 31/03/2026
DATA_EXPIRACAO = datetime(2026, 3, 31, 23, 59, 59)

def verificar_licenca():
    """Verifica se a licença está válida"""
    agora = datetime.now()
    if agora > DATA_EXPIRACAO:
        print("\n" + "=" * 60)
        print("⚠️  LICENÇA EXPIRADA")
        print("=" * 60)
        print(f"\nEste software expirou em: {DATA_EXPIRACAO.strftime('%d/%m/%Y')}")
        print(f"Data atual: {agora.strftime('%d/%m/%Y')}")
        print("\n💡 Entre em contato para renovar sua licença.")
        print("=" * 60 + "\n")
        input("Pressione ENTER para sair...")
        sys.exit(1)
    
    dias_restantes = (DATA_EXPIRACAO - agora).days
    if dias_restantes <= 7:
        print(f"\n⚠️  Atenção: Licença expira em {dias_restantes} dias ({DATA_EXPIRACAO.strftime('%d/%m/%Y')})\n")

# Função para instalar pacotes caso não existam
def instalar_se_necessario(pacote, import_name=None):
    """Instala pacote se não estiver disponível"""
    if import_name is None:
        import_name = pacote
    try:
        __import__(import_name)
    except ImportError:
        print(f"📦 Instalando {pacote}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pacote, "-q"])

# Verificar e instalar dependências
instalar_se_necessario("beautifulsoup4", "bs4")
instalar_se_necessario("openpyxl")
instalar_se_necessario("selenium")
instalar_se_necessario("webdriver-manager", "webdriver_manager")

from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def extrair_odds_do_html(html_source):
    """
    Extrai odds do HTML fornecido.
    Retorna lista de dicts com 'odd' e 'cor'.
    """
    soup = BeautifulSoup(html_source, 'html.parser')
    odds_data = []
    
    # Encontrar todos os divs com classe br-tb e cor de fundo
    all_divs = soup.find_all('div', recursive=True)
    
    for div in all_divs:
        style = div.get('style', '')
        
        # Verificar se tem cor de fundo
        if 'background-color' not in style:
            continue
        if '#dc3545' not in style and '#28a745' not in style:
            continue
        
        cor = 'Vermelho' if '#dc3545' in style else 'Verde'
        
        # Procurar texto com a odd
        text_content = div.get_text(strip=True)
        
        # Verificar se contém número com ponto (é uma odd)
        if text_content and '.' in text_content:
            # Extrair apenas a parte numérica
            match = re.search(r'\d+\.?\d*', text_content)
            if match:
                odd_value = match.group()
                odds_data.append({
                    'odd': odd_value,
                    'cor': cor
                })
    
    return odds_data


def agregar_odds(odds_data):
    """
    Agrega odds idênticas e conta ocorrências por cor.
    Retorna lista de dicts ordenada por valor numérico.
    """
    resumo = defaultdict(lambda: {'verde': 0, 'vermelho': 0})
    
    for item in odds_data:
        odd_val = float(item['odd'])
        cor_type = 'verde' if item['cor'] == 'Verde' else 'vermelho'
        resumo[odd_val][cor_type] += 1
    
    # Converter para lista e ordenar por valor numérico
    dados_planilha = []
    for odd_val in sorted(resumo.keys()):
        total = resumo[odd_val]['verde'] + resumo[odd_val]['vermelho']
        dados_planilha.append({
            'Odd': str(odd_val),
            'Verde': resumo[odd_val]['verde'],
            'Vermelho': resumo[odd_val]['vermelho'],
            'Total': total
        })
    
    return dados_planilha


def criar_excel(dados_planilha, nome_arquivo):
    """
    Cria workbook Excel com dados formatados e salva.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Odds"
    
    # Definir estilos
    header_fill = PatternFill(start_color="0f273d", end_color="0f273d", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    header_alignment = Alignment(horizontal='center', vertical='center')
    
    verde_fill = PatternFill(start_color="28a745", end_color="28a745", fill_type="solid")
    vermelho_fill = PatternFill(start_color="dc3545", end_color="dc3545", fill_type="solid")
    
    center_alignment = Alignment(horizontal='center', vertical='center')
    
    # Cabeçalho
    headers = ['Odd', 'Verde', 'Vermelho', 'Total']
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
    
    # Dados
    for row_idx, item in enumerate(dados_planilha, 2):
        # Coluna Odd
        cell = ws.cell(row=row_idx, column=1)
        cell.value = item['Odd']
        cell.alignment = center_alignment
        
        # Coluna Verde
        cell = ws.cell(row=row_idx, column=2)
        cell.value = item['Verde']
        cell.fill = verde_fill
        cell.alignment = center_alignment
        
        # Coluna Vermelho
        cell = ws.cell(row=row_idx, column=3)
        cell.value = item['Vermelho']
        cell.fill = vermelho_fill
        cell.alignment = center_alignment
        
        # Coluna Total
        cell = ws.cell(row=row_idx, column=4)
        cell.value = item['Total']
        cell.alignment = center_alignment
    
    # Ajustar largura das colunas
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 10
    
    # Salvar arquivo
    wb.save(nome_arquivo)
    return nome_arquivo


def capturar_html_selenium(driver):
    """
    Captura o HTML do grid usando o elemento já carregado.
    """
    try:
        # Pegar o elemento do grid com a classe br-tb
        grid_element = driver.find_element(By.CLASS_NAME, 'br-tb')
        
        # Capturar o HTML completo da página
        html_grid = driver.page_source
        
        return html_grid
    except Exception as e:
        print(f"❌ Erro ao capturar grid: {e}")
        return None


def fazer_login(driver):
    """
    Faz login no site UltraVirtual.
    """
    login = "diegosantosdeassis5@gmail.com"
    senha = "bomba2022"
    
    print("\n🔐 Fazendo login...")
    
    try:
        # Clicar no botão ENTRAR
        button_ENTRAR = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(text(), 'ENTRAR')]"))
        )
        button_ENTRAR.click()
        
        # Preencher e-mail
        input1 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='email' or @type='text']"))
        )
        input1.clear()
        input1.send_keys(login)
        
        # Preencher senha
        input2 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@type='password']"))
        )
        input2.clear()
        input2.send_keys(senha)
        input2.submit()
        
        print("✓ Login realizado com sucesso!")
        time.sleep(2)
        return True
        
    except Exception as e:
        print(f"❌ Erro no login: {e}")
        return False


def navegar_ate_odds(driver):
    """
    Navega até a página de odds do Bet365.
    """
    try:
        print("\n🎯 Navegando até Bet365...")
        
        # Clicar em Bet365
        span_bet365 = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[text()='Bet365']"))
        )
        span_bet365.click()
        time.sleep(1)
        
        # Ir direto para a URL
        driver.get("https://ultravirtual.com.br/dashboard/bet365/world/hourly")
        time.sleep(2)
        
        # Clicar em Placar FT
        print("📊 Abrindo Placar FT...")
        placarFT_button_span = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Placar FT')]"))
        )
        placarFT_button_span.click()
        time.sleep(1)
        
        # Clicar em Odds
        print("💰 Abrindo grid de Odds...")
        odds_button_span = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Odds')]"))
        )
        odds_button_span.click()
        
        print("⏳ Aguardando odds carregarem completamente...")
        time.sleep(5)  # Esperar mais tempo para as odds carregarem
        
        # Aguardar até o grid estar presente
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "br-tb"))
        )
        
        print("✓ Navegação concluída! Grid carregado.")
        return True
        
    except Exception as e:
        print(f"❌ Erro na navegação: {e}")
        return False


def main():
    """
    Função principal - fluxo automatizado completo.
    """
    # Verificar licença primeiro
    verificar_licenca()
    
    print("=" * 60)
    print(" EXTRATOR AUTOMÁTICO DE ODDS PARA EXCEL")
    print(" UltraVirtual → Bet365 → Placar FT → Odds")
    print("=" * 60)
    
    # Pedir nome do arquivo
    while True:
        nome_arquivo = input("\n📝 Digite o nome do arquivo de saída (ex: resultado.xlsx): ").strip()
        if nome_arquivo:
            if not nome_arquivo.endswith('.xlsx'):
                nome_arquivo += '.xlsx'
            break
        print("⚠️  Nome do arquivo não pode estar vazio!")
    
    print(f"✓ Arquivo será salvo como: {nome_arquivo}")
    
    # Inicializar Selenium
    print("\n🌐 Iniciando navegador Chrome...")
    try:
        # Usar webdriver_manager para instalar/atualizar ChromeDriver automaticamente
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        print("✓ Chrome iniciado!")
    except Exception as e:
        print(f"❌ Erro ao iniciar Chrome: {e}")
        print("💡 Certifique-se de ter o Google Chrome instalado")
        return
    
    try:
        # Abrir site
        print("\n🔗 Acessando UltraVirtual...")
        driver.get("https://ultravirtual.com.br/")
        time.sleep(2)
        
        # Fazer login
        if not fazer_login(driver):
            driver.quit()
            return
        
        # Navegar até odds
        if not navegar_ate_odds(driver):
            driver.quit()
            return
        
        # Capturar HTML
        print("\n📥 Capturando HTML do grid...")
        html_grid = capturar_html_selenium(driver)
        
        if not html_grid:
            print("❌ Falha ao capturar HTML!")
            driver.quit()
            return
        
        print(f"✓ HTML capturado ({len(html_grid)} caracteres)")
        
    finally:
        print("\n🔒 Fechando navegador...")
        driver.quit()
    
    # Processar dados
    print("\n" + "=" * 60)
    print(" PROCESSANDO DADOS")
    print("=" * 60)
    
    print(f"\n🔍 Extraindo odds...")
    odds_data = extrair_odds_do_html(html_grid)
    print(f"✓ {len(odds_data)} odds extraídas")
    
    if not odds_data:
        print("❌ Nenhuma odd foi encontrada!")
        return
    
    print(f"\n📊 Agregando odds...")
    dados_planilha = agregar_odds(odds_data)
    print(f"✓ {len(dados_planilha)} odds únicas identificadas")
    
    # Estatísticas
    total_ocorrencias = len(odds_data)
    total_verde = sum(item['Verde'] for item in dados_planilha)
    total_vermelho = sum(item['Vermelho'] for item in dados_planilha)
    
    print(f"\n📈 ESTATÍSTICAS:")
    print(f"   • Odds únicas: {len(dados_planilha)}")
    print(f"   • Total de ocorrências: {total_ocorrencias}")
    print(f"   • Verde (ganhos): {total_verde} ({100*total_verde/total_ocorrencias:.1f}%)")
    print(f"   • Vermelho (perdas): {total_vermelho} ({100*total_vermelho/total_ocorrencias:.1f}%)")
    
    # Gerar Excel
    print(f"\n💾 Gerando arquivo Excel...")
    arquivo_salvo = criar_excel(dados_planilha, nome_arquivo)
    print(f"✅ Arquivo salvo com sucesso: {arquivo_salvo}")
    
    # Mostrar amostra
    print(f"\n📋 AMOSTRA DOS DADOS (primeiras 10 linhas):")
    print(f"{'Odd':<8} {'Verde':<10} {'Vermelho':<12} {'Total':<8}")
    print("-" * 40)
    for item in dados_planilha[:10]:
        print(f"{item['Odd']:<8} {item['Verde']:<10} {item['Vermelho']:<12} {item['Total']:<8}")
    
    if len(dados_planilha) > 10:
        print(f"... e mais {len(dados_planilha) - 10} linhas")
    
    print("\n✨ Processo concluído com sucesso!")
    print("=" * 60)


if __name__ == "__main__":
    main()
