import argparse
import logging
import os
import shutil
from datetime import datetime
from pathlib import Path
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

BASE = Path(__file__).resolve().parent
CAMPOS = ["Nome", "CPF", "WhatsApp"]
SELETOR_LINHAS = (By.CSS_SELECTOR, "#corpo_tabela > tr")

def validar_alunos(arquivo):
    df = pd.read_excel(arquivo, dtype=str, keep_default_na=False, engine="openpyxl")
    if not set(CAMPOS).issubset(df.columns):
        raise ValueError(f"O Excel deve conter as colunas: {', '.join(CAMPOS)}")

    alunos, cpfs = [], set()
    for numero, (_, linha) in enumerate(df.iterrows(), start=2):
        aluno = {campo: str(linha[campo]).strip() for campo in CAMPOS}
        aluno["Linha"] = numero
        erros = [f"{campo} não informado" for campo in CAMPOS if not aluno[campo]]
        cpf_limpo = "".join(c for c in aluno["CPF"] if c.isdigit())
        if not erros:
            if cpf_limpo in cpfs:
                erros.append("CPF duplicado no lote")
            cpfs.add(cpf_limpo)
        aluno["Status do erro"] = ", ".join(erros)
        aluno["Confirmado"] = False
        if erros:
            logging.warning("Linha %s ignorada: %s", numero, aluno["Status do erro"])
        alunos.append(aluno)
    return alunos

def cadastrar_aluno(driver, aluno):
  # Preenche o formulário e devolve a posição da nova linha confirmada no DOM.
    espera = WebDriverWait(driver, 10)
    for campo, identificador in (("Nome", "nome"), ("WhatsApp", "whatsapp")):
        entrada = espera.until(EC.visibility_of_element_located((By.ID, identificador)))
        entrada.clear()
        entrada.send_keys(aluno[campo])

    posicao = len(driver.find_elements(*SELETOR_LINHAS))
    espera.until(EC.element_to_be_clickable((By.ID, "cadastrar"))).click()

    def confirmado(navegador):
        linhas = navegador.find_elements(*SELETOR_LINHAS)
        if len(linhas) != posicao + 1:
            return False
        celulas = linhas[posicao].find_elements(By.TAG_NAME, "td")
        return [c.get_property("textContent") for c in celulas] == [aluno["Nome"], aluno["WhatsApp"]]

    espera.until(confirmado)
    return posicao

def executar_cadastros(alunos, arquivo_html, arquivo_driver):
    validos = [a for a in alunos if not a["Status do erro"]]
    if not validos:
        return []

    tentativas, historico = [], []
    falha_geral = False
    try:
        if not arquivo_html.is_file() or not arquivo_driver.is_file():
            raise FileNotFoundError("Confira web/index.html e drivers/chromedriver.exe")

        # O contexto fecha o Chrome mesmo quando ocorrer uma exceção.
        with webdriver.Chrome(service=Service(str(arquivo_driver))) as driver:
            driver.get(arquivo_html.as_uri())
            for aluno in validos:
                try:
                    posicao = cadastrar_aluno(driver, aluno)
                    tentativas.append((aluno, posicao))
                    logging.info("Linha %s adicionada à página", aluno["Linha"])
                except Exception:
                    aluno["Status do erro"] = "Falha na interação com a página"
                    logging.exception("Erro no cadastro da linha %s", aluno["Linha"])
            html = driver.page_source

        corpo = BeautifulSoup(html, "html.parser").select_one("tbody#corpo_tabela")
        if corpo is None:
            raise ValueError("Histórico de inscrições não encontrado no HTML")
        for tr in corpo.find_all("tr", recursive=False):
            historico.append(tuple(td.get_text() for td in tr.find_all("td", recursive=False)))
    except Exception:
        falha_geral = True
        logging.exception("Falha no navegador ou na extração do histórico")

    sucessos = []
    for aluno, posicao in tentativas:
        esperado = (aluno["Nome"], aluno["WhatsApp"])
        if posicao < len(historico) and historico[posicao] == esperado:
            aluno["Confirmado"] = True
            sucessos.append({"Nome": esperado[0], "WhatsApp": esperado[1]})

    for aluno in validos:
        if not aluno["Confirmado"]:
            aluno["Status do erro"] = aluno["Status do erro"] or (
                "Falha na interação com a página" 
                if falha_geral else "Cadastro ausente do histórico")
            logging.warning("Linha %s: %s", aluno["Linha"], aluno["Status do erro"])
    return sucessos

def salvar_excel(caminho, abas):
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        for nome, df in abas.items():
            df.to_excel(writer,sheet_name=nome,index=False)
            folha = writer.sheets[nome]
            folha.freeze_panes = "A2"
            folha.auto_filter.ref = folha.dimensions

def gerar_relatorios(alunos, sucessos, pasta):
    pasta.mkdir(parents=True, exist_ok=True)
    pendentes = [a for a in alunos if not a["Confirmado"]]
    total = len(alunos)
    if total != len(sucessos) + len(pendentes):
        raise ValueError("Divergência entre sucessos e pendências")
    indicadores = {
        "Total processado": total,
        "Sucessos": len(sucessos),
        "Pendências": len(pendentes),
        "Taxa de sucesso": len(sucessos) / total if total else 0,
        "Taxa de pendência": len(pendentes) / total if total else 0,
    }
    registros = [
        {**a, "Status": "Sucesso" if a["Confirmado"] else "Pendência"} for a in alunos
    ]
    painel = [
        {"Linha": a["Linha"], "Status": a["Status"], "Status do erro": a["Status do erro"],
         "Sucesso": int(a["Confirmado"]), "Pendência": int(not a["Confirmado"])}
        for a in registros
    ]
    arquivos = {
        "relatorio_sucesso_historico.xlsx": {"Histórico": pd.DataFrame(sucessos, columns=["Nome", "WhatsApp"])},
        "relatorio_pendencias.xlsx": {"Pendências": pd.DataFrame(pendentes, columns=["Nome", "CPF", "Status do erro"])},
        "base_consolidada.xlsx": {
            "Resultados": pd.DataFrame(registros, columns=["Linha", *CAMPOS, "Status", "Status do erro"]),
            "Indicadores": pd.DataFrame([indicadores]),
        },
        "base_painel.xlsx": {"Painel": pd.DataFrame(painel)},
    }
    for nome, abas in arquivos.items():
        salvar_excel(pasta / nome, abas)
        logging.info("Relatório criado: %s", pasta / nome)
    return indicadores

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--so-validar", action="store_true", help="Valida o Excel sem abrir o Chrome")
    args = parser.parse_args()
    lote = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    os.makedirs(BASE / "logs", exist_ok=True)
    logging.basicConfig(filename=BASE / "logs" / "automacao.log", level=logging.INFO,
                        format="%(asctime)s - %(levelname)s - %(message)s", encoding="utf-8")
    try:
        arquivo = BASE / "dados" / "alunos_projeto_final.xlsx"
        alunos = validar_alunos(arquivo)
        if args.so_validar:
            pendentes = sum(bool(a["Status do erro"]) for a in alunos)
            print(f"Total: {len(alunos)} | Válidos: {len(alunos)-pendentes} | Pendentes: {pendentes}")
            return 0

        # Uso útil do shutil: backup da entrada, sem recriar processados/.
        backup = BASE / "backup"
        backup.mkdir(exist_ok=True)
        shutil.copy2(arquivo, backup / f"alunos_{lote}.xlsx")
        logging.info("Backup da planilha criado para o lote %s", lote)

        sucessos = executar_cadastros(alunos, BASE / "web" / "index.html",
                                     BASE / "drivers" / "chromedriver.exe")
        indicadores = gerar_relatorios(alunos, sucessos, BASE / "relatorios" / lote)
        for nome, valor in indicadores.items():
            print(f"{nome}: {valor:.2%}" if nome.startswith("Taxa") else f"{nome}: {valor}")
        print(f"Relatórios: relatorios/{lote}")
        return 0
    except Exception:
        logging.exception("Erro geral na execução")
        print("O robô encontrou um erro. Veja logs/automacao.log")
        return 1

if __name__ == "__main__":
    raise SystemExit(main())