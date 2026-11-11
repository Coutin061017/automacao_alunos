import argparse
import logging
import os
import shutil
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


BASE_DIR = Path(__file__).resolve().parent
CAMPOS_OBRIGATORIOS = ("Nome", "CPF", "WhatsApp")
CAMPOS_HISTORICO = ("Nome", "WhatsApp")
SELETOR_LINHAS = (By.CSS_SELECTOR, "#corpo_tabela > tr")


def ler_planilha(arquivo):
    return pd.read_excel(
        arquivo, engine="openpyxl", dtype=str, keep_default_na=False
    )


def validar_alunos(planilha):
    ausentes = []
    for campo in CAMPOS_OBRIGATORIOS:
        if campo not in planilha.columns:
            ausentes.append(campo)
    if ausentes:
        raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(sorted(ausentes))}")

    alunos = []
    cpfs_vistos = set()
    # A primeira linha do Excel é o cabeçalho.
    for numero, (_, linha) in enumerate(planilha.iterrows(), start=2):
        aluno = {"Linha": numero}
        faltas = []
        for campo in CAMPOS_OBRIGATORIOS:
            aluno[campo] = linha[campo].strip()
            if not aluno[campo]:
                faltas.append(f"{campo} não informado")

        # Removemos a pontuação só para comparar CPFs.
        cpf = aluno["CPF"].replace(".", "").replace("-", "").replace(" ", "")
        if not faltas:
            if cpf in cpfs_vistos:
                faltas.append("CPF duplicado no lote")
            else:
                cpfs_vistos.add(cpf)
        aluno["Status do erro"] = ", ".join(faltas)
        if faltas:
            logging.warning("Linha %d: %s", numero, aluno["Status do erro"])
        alunos.append(aluno)
    return alunos


def cadastrar_aluno(driver, nome, whatsapp, timeout=5):
    espera = WebDriverWait(driver, timeout)
    campo_nome = espera.until(EC.visibility_of_element_located((By.ID, "nome")))
    campo_whatsapp = espera.until(EC.visibility_of_element_located((By.ID, "whatsapp")))
    campo_nome.clear()
    campo_nome.send_keys(nome)
    campo_whatsapp.clear()
    campo_whatsapp.send_keys(whatsapp)

    botao = espera.until(EC.element_to_be_clickable((By.ID, "cadastrar")))
    posicao = len(driver.find_elements(*SELETOR_LINHAS))
    botao.click()

    def linha_confirmada(navegador):
        linhas = navegador.find_elements(*SELETOR_LINHAS)
        if len(linhas) != posicao + 1:
            return False
        celulas = linhas[posicao].find_elements(By.TAG_NAME, "td")
        # textContent preserva espaços internos que o Chrome normaliza em .text.
        return [celula.get_property("textContent") for celula in celulas] == [nome, whatsapp]

    # O clique só conta quando a tabela mostra os dados enviados.
    espera.until(linha_confirmada)
    return posicao


def extrair_historico(html):
    tabela = BeautifulSoup(html, "html.parser").select_one("tbody#corpo_tabela")
    if tabela is None:
        raise ValueError("Tabela de histórico não encontrada.")
    historico = []
    for linha in tabela.find_all("tr", recursive=False):
        valores = []
        for celula in linha.find_all("td", recursive=False):
            valores.append(celula.get_text())
        historico.append(tuple(valores))
    return historico


def executar_cadastros(alunos, arquivo_html, arquivo_driver, headless=False, timeout=5):
    validos = [aluno for aluno in alunos if not aluno["Status do erro"]]
    if not validos:
        return [], False

    driver = None
    cadastros_na_pagina = []
    historico = []
    falha_navegador = False
    try:
        if not arquivo_html.is_file():
            raise FileNotFoundError("web/index.html não encontrado.")
        if not arquivo_driver.is_file():
            raise FileNotFoundError("ChromeDriver local não encontrado. Confira --driver.")
        opcoes = webdriver.ChromeOptions()
        if headless:
            opcoes.add_argument("--headless=new")
        driver = webdriver.Chrome(service=Service(str(arquivo_driver)), options=opcoes)
        driver.get(arquivo_html.as_uri())
        for aluno in validos:
            try:
                posicao = cadastrar_aluno(driver, aluno["Nome"], aluno["WhatsApp"], timeout)
                cadastros_na_pagina.append((aluno, posicao))
            except TimeoutException:
                aluno["Status do erro"] = "Erro ao cadastrar: formulário ou confirmação indisponível"
            except WebDriverException as erro:
                aluno["Status do erro"] = "Falha na interação com a página"
                logging.warning("Linha %d: %s", aluno["Linha"], type(erro).__name__)
        historico = extrair_historico(driver.page_source)
    except (OSError, ValueError, WebDriverException) as erro:
        falha_navegador = True
        if isinstance(erro, FileNotFoundError):
            print(erro)
        else:
            print(f"Falha no navegador ({type(erro).__name__}). Confira o HTML e as versões do Chrome e do ChromeDriver.")
        logging.error("Falha no navegador: %s", type(erro).__name__)
    finally:
        if driver is not None:
            try:
                driver.quit()
            except WebDriverException:
                logging.warning("Falha ao encerrar o Chrome.")
                falha_navegador = True

    sucessos = []
    # A posição liga o cadastro à linha extraída no fim da execução.
    for aluno, posicao in cadastros_na_pagina:
        esperado = (aluno["Nome"], aluno["WhatsApp"])
        if posicao < len(historico) and historico[posicao] == esperado:
            aluno["Confirmado"] = True
            nome, whatsapp = historico[posicao]
            sucessos.append({"Nome": nome, "WhatsApp": whatsapp})

    for aluno in validos:
        if aluno.get("Confirmado"):
            logging.info("Linha %d confirmada no histórico.", aluno["Linha"])
            continue
        if not aluno["Status do erro"]:
            if falha_navegador:
                aluno["Status do erro"] = "Histórico indisponível"
            else:
                aluno["Status do erro"] = "Cadastro ausente do histórico"
        logging.warning("Linha %d: %s", aluno["Linha"], aluno["Status do erro"])
    return sucessos, falha_navegador


def gravar_excel(arquivo, abas):
    # Força valores de entrada a permanecerem texto, inclusive se começam com "=".
    with pd.ExcelWriter(arquivo, engine="openpyxl") as writer:
        for nome, tabela in abas.items():
            tabela.to_excel(writer, sheet_name=nome, index=False)
            aba = writer.sheets[nome]
            for linha in aba.iter_rows(min_row=2):
                for celula in linha:
                    if isinstance(celula.value, str):
                        celula.data_type = "s"
                        celula.number_format = "@"
            aba.freeze_panes = "A2"
            aba.auto_filter.ref = aba.dimensions
        if "Indicadores" in writer.sheets:
            for coluna in ("D", "E"):
                writer.sheets["Indicadores"][f"{coluna}2"].number_format = "0.00%"


def gerar_relatorios(pasta, lote, alunos, sucessos):
    pendencias = [aluno for aluno in alunos if not aluno.get("Confirmado")]
    total = len(alunos)
    if total != len(sucessos) + len(pendencias):
        raise ValueError("A contagem do histórico diverge dos registros classificados.")
    indicadores = {
        "Total processado": total,
        "Sucessos": len(sucessos),
        "Pendências": len(pendencias),
        "Taxa de sucesso": len(sucessos) / total if total else 0,
        "Taxa de pendência": len(pendencias) / total if total else 0,
    }
    base_local = []
    base_painel = []
    for aluno in alunos:
        if aluno.get("Confirmado"):
            status = "Sucesso"
            sucesso = 1
        else:
            status = "Pendência"
            sucesso = 0

        resultado = aluno.copy()
        resultado["Status"] = status
        base_local.append(resultado)

        # O painel identifica o registro pelo lote e pela linha do Excel.
        base_painel.append({
            "Lote": lote,
            "Linha": aluno["Linha"],
            "Status": status,
            "Status do erro": aluno["Status do erro"],
            "Sucesso": sucesso,
            "Pendência": 1 - sucesso,
        })
    relatorios = {
        "relatorio_sucesso_historico.xlsx": {
            "Histórico": pd.DataFrame(sucessos, columns=CAMPOS_HISTORICO)
        },
        "relatorio_pendencias.xlsx": {
            "Pendências": pd.DataFrame(pendencias, columns=["Nome", "CPF", "Status do erro"])
        },
        "base_consolidada.xlsx": {
            "Resultados": pd.DataFrame(
                base_local, columns=["Linha", *CAMPOS_OBRIGATORIOS, "Status", "Status do erro"]
            ),
            "Indicadores": pd.DataFrame([indicadores]),
        },
        "base_painel.xlsx": {
            "Painel": pd.DataFrame(
                base_painel, columns=["Lote", "Linha", "Status", "Status do erro", "Sucesso", "Pendência"]
            ),
        },
    }
    originais = pasta / "relatorios" / lote
    processados = pasta / "processados" / lote
    os.makedirs(originais.parent, exist_ok=True)
    os.makedirs(processados.parent, exist_ok=True)
    with TemporaryDirectory(dir=pasta) as temporaria:
        temporaria = Path(temporaria)
        gerados = temporaria / "originais"
        copias = temporaria / "copias"
        gerados.mkdir()
        copias.mkdir()
        for nome, abas in relatorios.items():
            gravar_excel(gerados / nome, abas)
            logging.info("Relatório gravado: %s", nome)
            shutil.copy2(gerados / nome, copias / nome)
            logging.info("Cópia preparada: %s", nome)
        shutil.move(gerados, originais)
        logging.info("Originais disponíveis em relatorios/%s.", lote)
        shutil.move(copias, processados)
        logging.info("Cópias movidas para processados/%s.", lote)
    return indicadores


def main(argv=None):
    parser = argparse.ArgumentParser(description="Automação de cadastro de alunos")
    parser.add_argument("--so-validar", action="store_true", help="Valida sem abrir o Chrome.")
    parser.add_argument("--pasta-projeto", type=Path, default=BASE_DIR)
    parser.add_argument("--driver", type=Path, help="Caminho do ChromeDriver local.")
    parser.add_argument("--headless", action="store_true", help="Executa o Chrome sem janela.")
    args = parser.parse_args(argv)
    pasta = args.pasta_projeto.resolve()
    arquivo_driver = args.driver or pasta / "drivers" / "chromedriver.exe"
    lote = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    try:
        os.makedirs(pasta / "logs", exist_ok=True)
        logging.basicConfig(
            filename=pasta / "logs" / "automacao.log", level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8",
        )
        logging.info("Início do lote %s.", lote)
        # 1. Ler a planilha e validar os campos obrigatórios.
        arquivo_alunos = pasta / "dados" / "alunos_projeto_final.xlsx"
        planilha = ler_planilha(arquivo_alunos)
        alunos = validar_alunos(planilha)
        if args.so_validar:
            pendentes = sum(bool(aluno["Status do erro"]) for aluno in alunos)
            print(f"Lidos: {len(alunos)} | Válidos: {len(alunos) - pendentes} | Pendências: {pendentes}")
            return 0
        # 2. Cadastrar os alunos válidos e conferir o histórico.
        arquivo_html = pasta / "web" / "index.html"
        sucessos, falha = executar_cadastros(
            alunos, arquivo_html, arquivo_driver.resolve(), args.headless
        )

        # 3. Gravar os relatórios e mostrar o resumo do lote.
        indicadores = gerar_relatorios(pasta, lote, alunos, sucessos)
        for nome, valor in indicadores.items():
            valor_exibido = f"{valor:.2%}" if nome.startswith("Taxa") else str(valor)
            print(f"{nome}: {valor_exibido}")
        print(f"Arquivos do lote: relatorios/{lote} e processados/{lote}")
        logging.info("Lote %s concluído. %s", lote, indicadores)
        return int(falha)
    except FileNotFoundError:
        print("Arquivo ou pasta não encontrado. Confira dados/alunos_projeto_final.xlsx e as pastas do projeto.")
        logging.error("Arquivo ou pasta não encontrado.")
    except ValueError as erro:
        print(f"Execução interrompida: {erro}")
        logging.error("Dados ou estrutura inválidos.")
    except (OSError, ImportError) as erro:
        print(f"Execução interrompida ({type(erro).__name__}). Confira colunas, dependências e permissões de gravação.")
        logging.error("Execução interrompida: %s")
    except Exception:
         print("Erro inesperado durante a execução. Consulte o arquivo de log.")
    logging.exception("Erro inesperado durante a execução.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
