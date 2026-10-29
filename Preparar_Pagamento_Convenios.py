# -*- coding: utf-8 -*-
# Por Izabel Galvão :3
"""PP Despesa Empenhada — revisão do fluxo original de Louise/SESDEC.

Execute na mesma pasta de certificar_liquidar_data_pagamento_v2.py.
Usa as validações e o preenchimento de datas já existentes nesse módulo.
Não executa certificação nem liquidação ao importá-lo.
"""
from pathlib import Path
from datetime import datetime
import json
import re
import time
import io
from contextlib import redirect_stdout
from playwright.sync_api import Error as PlaywrightError

import Certificar_Liquidar_Convenios as base

PASTA = Path(__file__).resolve().parent
PLANILHA = PASTA / "Pagamentos_Tratado.xlsx"
PASTA_RECUPERACAO = PASTA / "RECUPERACAO"
ARQUIVO_RECUPERACAO_PP = PASTA_RECUPERACAO / "pp_pendente.json"
STATUS_PESQUISAR_PP = "PESQUISAR_PP"
# Preserva as escolhas do código original.
ORDEM_CRONOLOGICA = "Não Obedece Ordem Cronológica"
TIPO_ORDEM_BANCARIA = "Regularização"

# Velocidade do robô: milissegundos por ação e segundos entre etapas.
ATRASO_ACAO_MS = 100
PAUSA_ETAPA = 0.9
PAUSA_NUMERO_GERADO = 1.0
PAUSA_APOS_SALVAR = 1.0


def pausar(segundos):
    limite = time.monotonic() + segundos
    while True:
        base.verificar_panico()
        restante = limite - time.monotonic()
        if restante <= 0:
            return
        time.sleep(min(0.3, restante))


def normalizar_codigo(valor):
    codigo = re.sub(r"[\s.\-]", "", base.texto(valor)).upper()
    if not re.fullmatch(r"[0-9]+X?", codigo):
        raise base.ErroDados(f"Código bancário inválido: {valor!r}")
    return codigo.lstrip("0") or "0"


def extrair_pp(texto):
    numeros = set(re.findall(r"\b\d{4}PP\d{6}\b", texto.upper()))
    if len(numeros) != 1:
        raise base.ErroValidacaoSIGEF("A resposta não contém uma única PP válida.")
    return numeros.pop()


def titulo(texto):
    print("\n" + "=" * 64)
    print(texto)
    print("=" * 64)


def pendentes(planilhas):
    titulo("PREPARAÇÃO DE PAGAMENTO | CONFERÊNCIA DA PLANILHA")
    print(f"Arquivo: {PLANILHA.name}")
    saida = io.StringIO()
    try:
        with redirect_stdout(saida):
            ultima = base.prevalidar_planilha(planilhas)
    except BaseException:
        print(saida.getvalue(), end="")
        raise
    # Preserva os resultados e avisos, retirando apenas cabeçalhos repetidos.
    for mensagem in saida.getvalue().splitlines():
        if not mensagem.strip() or set(mensagem.strip()) == {"="}:
            continue
        if mensagem in ("PRÉ-VALIDAÇÃO DA PLANILHA", "PLANILHA APROVADA. Iniciando SIGEF..."):
            continue
        print("  " + mensagem)
    print("\n[OK] Planilha aprovada. Conferindo pagamentos pendentes...\n")
    registros, vistos = [], set()
    existentes = revisar = sem_nl = 0
    for linha in range(2, ultima + 1):
        if base.linha_vazia(planilhas["entrada"], linha):
            continue
        r = base.validar_e_formatar_registro(base.ler_registro(planilhas, linha))
        pp = r["preparacao_pagamento"].upper()
        if pp:
            if pp != STATUS_PESQUISAR_PP:
                if not re.fullmatch(r"\d{4}PP\d{6}", pp) or pp in vistos:
                    raise base.ErroDados(f"Linha {linha}: PP inválida ou duplicada: {pp}")
                vistos.add(pp)
            if pp == STATUS_PESQUISAR_PP:
                revisar += 1
                print(f"[ATENÇÃO] Linha {linha} | {r['nome']} | Localizar PP no SIGEF.")
            else:
                existentes += 1
            continue
        if r["ordem_bancaria"]:
            raise base.ErroDados(f"Linha {linha}: existe OB sem PP; confira a planilha.")
        if not r["liquidacao"] or r["liquidacao"] == base.STATUS_PESQUISAR_NL:
            sem_nl += 1
            continue
        if not re.fullmatch(r"\d{4}NL\d{6}", r["liquidacao"]):
            raise base.ErroDados(f"Linha {linha}: NL inválida.")
        for campo in ("banco", "agencia", "conta"):
            normalizar_codigo(r[campo])
        registros.append(r)
    titulo("PREPARAÇÃO DE PAGAMENTOS")
    print(f"Já preparados: {existentes} | A preparar: {len(registros)} ")
    if sem_nl:
        print(f"Aguardando NL válida: {sem_nl}")
    return registros


def salvar_pp(planilhas, r, numero):
    ws = planilhas["pp"]
    atual = base.texto(ws.cell(r["linha"], base.COL_PP).value)
    if atual and atual not in (STATUS_PESQUISAR_PP, numero):
        raise base.ErroDados(f"Linha {r['linha']}: já contém outra PP: {atual}")
    base.sincronizar_dados_base_saida(ws, r)
    for coluna, valor in ((base.COL_CE, r["despesa_certificada"]),
                          (base.COL_NL, r["liquidacao"]), (base.COL_PP, numero)):
        base._escrever_texto(ws, r["linha"], coluna, valor)
    base.salvar_planilha_segura(planilhas["book"])
    if numero == STATUS_PESQUISAR_PP:
        base.mensagem_detalhada("  [PROTEÇÃO] Linha marcada para conferência em caso de interrupção.")
    else:
        print(f"  [SALVO] {numero} ")


def registrar_pp_pendente(r, numero, domicilio):
    PASTA_RECUPERACAO.mkdir(parents=True, exist_ok=True)
    temporario = ARQUIVO_RECUPERACAO_PP.with_suffix(".tmp")
    dados = {
        "linha": r["linha"],
        "nl": r["liquidacao"],
        "pp": numero,
        "data": datetime.now().isoformat(),
        "domicilio_selecionado": domicilio,
    }
    temporario.write_text(
        json.dumps(dados, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporario.replace(ARQUIVO_RECUPERACAO_PP)


def limpar_pp_pendente():
    if ARQUIVO_RECUPERACAO_PP.exists():
        ARQUIVO_RECUPERACAO_PP.unlink()


def recuperar_pp_pendente(planilhas):
    if not ARQUIVO_RECUPERACAO_PP.exists():
        return

    try:
        dados = json.loads(ARQUIVO_RECUPERACAO_PP.read_text(encoding="utf-8"))
        linha = int(dados["linha"])
        numero = base.texto(dados["pp"]).upper()
        nl = base.texto(dados["nl"]).upper()
    except Exception as exc:
        raise base.ErroDados(
            f"O arquivo de recuperação {ARQUIVO_RECUPERACAO_PP.name} "
            "está inválido. Confira-o antes de executar o robô."
        ) from exc

    if not re.fullmatch(r"\d{4}PP\d{6}", numero):
        raise base.ErroDados(f"PP inválida no arquivo de recuperação: {numero!r}.")

    r = base.validar_e_formatar_registro(base.ler_registro(planilhas, linha))
    if r["liquidacao"].upper() != nl:
        raise base.ErroDados(
            f"A NL da linha {linha} mudou desde a geração da PP {numero}. "
            "Confira o documento no SIGEF antes de continuar."
        )

    resposta = base.pyautogui.confirm(
        text=(
            f"Existe uma PP confirmada no SIGEF aguardando gravação.\n\n"
            f"Linha: {linha}\nFavorecido: {r['nome']}\n"
            f"NL: {nl}\nPP: {numero}\n\n"
            "Deseja gravar essa PP na planilha agora?"
        ),
        title="Recuperar preparação de pagamento",
        buttons=["Recuperar", "Cancelar"],
    )
    if resposta != "Recuperar":
        raise base.ErroDados(
            "Recuperação de PP pendente. Automação encerrada sem gerar nova PP."
        )

    salvar_pp(planilhas, r, numero)
    limpar_pp_pendente()
    print(f"[RECUPERADO] PP {numero} gravada na linha {linha}.")


def esperar_campo(campo, esperado, normalizar=base.texto, timeout_segundos=15):
    campo.wait_for(state="visible", timeout=int(timeout_segundos * 1000))
    fim = time.monotonic() + timeout_segundos
    obtido = ""
    while time.monotonic() < fim:
        base.verificar_panico()
        obtido = campo.input_value()
        try:
            if normalizar(obtido) == normalizar(esperado):
                return
        except base.ErroDados:
            pass
        time.sleep(0.3)
    raise base.ErroValidacaoSIGEF(f"Campo divergente: esperado {esperado!r}, recebido {obtido!r}.")


def selecionar_nl(pagina, r):
    print(
        f"  [FAVORECIDO] Conferindo o CPF carregado pela NL: "
        f"{r['cpf_formatado']}",
        flush=True,
    )
    with pagina.expect_popup(timeout=15000) as evento:
        pagina.locator("#txtNotaLancamento_lnkBtnPesquisa").click()
    ordem = evento.value
    base.aguardar_rede(ordem)
    # Capturar a nova janela independentemente de qual popup é seu opener.
    with pagina.context.expect_page(timeout=15000) as evento:
        ordem.get_by_text(ORDEM_CRONOLOGICA, exact=True).click()
    pesquisa = evento.value
    base.aguardar_rede(pesquisa)
    ano, numero = r["liquidacao"].split("NL")
    base.preencher_sequencial(pesquisa.locator("#txtNotaLancamento_SIGEFPesquisa"), str(int(numero)))
    base.preencher_sequencial(pesquisa.locator('[name="txtNotaLancamentoSigla"]'), ano)
    pesquisa.get_by_role("button", name="Confirma a Consulta", exact=True).click()
    base.aguardar_rede(pesquisa)
    # O original selecionava sempre o índice 0; só fazê-lo se houver uma opção.
    opcoes = pesquisa.locator('td[onclick^="SelecionarItem("]')
    opcoes.first.wait_for(state="visible", timeout=15000)
    visiveis = {}
    for i in range(opcoes.count()):
        opcao = opcoes.nth(i)
        if opcao.is_visible():
            visiveis.setdefault(opcao.get_attribute("onclick"), opcao)
    if len(visiveis) != 1:
        raise base.ErroValidacaoSIGEF("A NL retornou mais de uma opção de fonte. Selecione e confira manualmente.")
    next(iter(visiveis.values())).click()
    base.aguardar_rede(pagina)
    esperar_campo(pagina.locator("#txtCredor_SIGEFPesquisa"), r["cpf_formatado"], base.somente_digitos)
    print(f"  [FAVORECIDO] CPF confirmado: {r['cpf_formatado']}", flush=True)


def selecionar_conta_automaticamente(pagina, r):
    print(
        f"  [DOMICÍLIO] Pesquisando banco {r['banco']}, "
        f"agência {r['agencia']}, conta {r['conta']}",
        flush=True,
    )
    pagina.locator("#cboTipoOrdemBancaria").select_option(label=TIPO_ORDEM_BANCARIA)
    base.preencher_sequencial(pagina.locator("#txtBanco"), r["banco"])
    with pagina.expect_popup(timeout=15000) as evento:
        pagina.locator("#txtConta_lnkBtnPesquisa").click()
    pesquisa = evento.value
    base.aguardar_rede(pesquisa, timeout=1000)
    pesquisa.get_by_role("button", name="Confirmar a Consulta").click()
    base.aguardar_rede(pesquisa, timeout=1000)
    # Células DIRETAS evitam contar novamente a tabela dentro de uma linha
    # externa do layout. A ordem vista no SIGEF é Banco / Agência / Conta.
    esperados = [normalizar_codigo(r[k]) for k in ("banco", "agencia", "conta")]
    limite = time.monotonic() + 1
    while True:
        base.verificar_panico()
        linhas = pesquisa.locator("tr")
        candidatas = []
        for i in range(linhas.count()):
            linha = linhas.nth(i)
            if not linha.is_visible():
                continue
            celulas = linha.locator("xpath=./td")
            valores = celulas.all_inner_texts()
            if len(valores) < 3:
                continue
            try:
                codigos = [normalizar_codigo(v) for v in valores[:3]]
            except base.ErroDados:
                continue
            if codigos == esperados:
                candidatas.append(celulas.nth(0))
        if len(candidatas) == 1:
            candidatas[0].click()
            break
        if len(candidatas) > 1:
            raise base.ErroValidacaoSIGEF(
                "Mais de um domicílio bancário possui os mesmos banco, agência e conta."
            )
        if time.monotonic() >= limite:
            raise base.ErroValidacaoSIGEF(
                "Domicílio bancário não encontrado após aguardar a consulta: "
                f"banco {r['banco']}, agência {r['agencia']}, conta {r['conta']}."
            )
        time.sleep(0.3)
    base.aguardar_rede(pagina, timeout=1000)
    for seletor, chave in (("#txtBanco", "banco"), ("#txtAgencia", "agencia"),
                           ("#txtConta_SIGEFPesquisa", "conta")):
        esperar_campo(pagina.locator(seletor), r[chave], normalizar_codigo, timeout_segundos=1)
    print("  [DOMICÍLIO] Banco, agência e conta conferidos.", flush=True)


def selecionar_conta(pagina, r):
    try:
        selecionar_conta_automaticamente(pagina, r)
        return {k: r[k] for k in ("banco", "agencia", "conta")}
    except (base.ErroValidacaoSIGEF, base.PlaywrightTimeoutError) as exc:
        base.verificar_panico()
        base.mensagem_detalhada(f"[SELEÇÃO MANUAL] {exc}")
        print("  [AÇÃO NECESSÁRIA] Selecione o domicílio no SIGEF e confirme na janela do robô.", flush=True)
        resposta = base.pyautogui.confirm(
            text=(f"Não foi possível concluir a seleção automática para {r['nome']}.\n"
                  f"Planilha: banco {r['banco']}, agência {r['agencia']}, conta {r['conta']}.\n\n"
                  "Selecione o domicílio bancário na janela do SIGEF.\n"
                  "Depois de selecionar, clique em Continuar nesta janela."),
            title="Seleção manual do domicílio bancário",
            buttons=["Continuar", "Cancelar"],
        )
        base.verificar_panico()
        if resposta != "Continuar":
            raise base.ErroValidacaoSIGEF("Seleção manual cancelada. A PP não foi confirmada.")
        base.aguardar_rede(pagina)
        selecionado = {}
        for seletor, chave in (("#txtBanco", "banco"), ("#txtAgencia", "agencia"),
                               ("#txtConta_SIGEFPesquisa", "conta")):
            campo = pagina.locator(seletor)
            campo.wait_for(state="visible", timeout=15000)
            selecionado[chave] = campo.input_value().strip()
            normalizar_codigo(selecionado[chave])
        if any(normalizar_codigo(selecionado[k]) != normalizar_codigo(r[k])
               for k in selecionado):
            resposta = base.pyautogui.confirm(
                text=("O domicílio selecionado é diferente da planilha.\n\n"
                      f"Planilha: {r['banco']} / {r['agencia']} / {r['conta']}\n"
                      f"Selecionado: {selecionado['banco']} / {selecionado['agencia']} / {selecionado['conta']}\n\n"
                      "Continuar a preparação com o domicílio selecionado?"),
                title="Confirmar domicílio bancário diferente",
                buttons=["SIM", "NÃO"],
            )
            base.verificar_panico()
            if resposta != "SIM":
                raise base.ErroValidacaoSIGEF("Domicílio diferente não autorizado. A PP não foi confirmada.")
        return selecionado


def preparar_linha(pagina, planilhas, r):
    base.verificar_panico()
    base.aguardar_rede(pagina)
    base.preencher_data_sigef(pagina.locator("#txtDataReferencia_SIGEFData"), r["data_pagamento"])
    base.preencher_sequencial(pagina.locator("#txtGestao_SIGEFPesquisa"), r["gestao"])
    pausar(PAUSA_ETAPA)
    print("  [1/5] Consultando a NL e conferindo o favorecido...")
    selecionar_nl(pagina, r)
    pausar(PAUSA_ETAPA)
    print("  [2/5] Selecionando o domicílio bancário...")
    domicilio = selecionar_conta(pagina, r)
    pausar(PAUSA_ETAPA)
    observacao = base.montar_observacao_sigef(
        "Preparação de Pagamento",
        r,
    )
    base.preencher_sequencial(pagina.locator("#txtObservacao"), observacao)
    print("  [3/5] Consultando retenções e conferindo a confirmação...")
    pagina.get_by_role("button", name="Sugerir Retenções").click()
    pagina.get_by_text("Não existem sugestões para").wait_for(state="visible", timeout=30000)
    pagina.locator("#menun7").get_by_role("link").click()
    base.aguardar_rede(pagina)
    for seletor, chave in (("#txtBancoConf", "banco"), ("#txtAgenciaConf", "agencia"),
                           ("#txtContaConf", "conta")):
        esperar_campo(pagina.locator(seletor), domicilio[chave], normalizar_codigo)
    botao = pagina.get_by_role("button", name="Confirmar a Operação")
    botao.wait_for(state="visible", timeout=15000)
    base.verificar_panico()
    # Persistir antes do clique: uma interrupção não pode gerar uma segunda PP.
    salvar_pp(planilhas, r, STATUS_PESQUISAR_PP)
    print("  [4/5] Confirmando a preparação no SIGEF...", flush=True)
    botao.click()
    mensagem = pagina.get_by_text("Operação realizada com")
    mensagem.wait_for(state="visible", timeout=30000)
    pausar(PAUSA_NUMERO_GERADO)
    numero = extrair_pp(mensagem.inner_text())
    # Preservar o número também se o Excel impedir a gravação do workbook.
    registrar_pp_pendente(r, numero, domicilio)
    print("  [5/5] Gravando a PP gerada na planilha...", flush=True)
    salvar_pp(planilhas, r, numero)
    limpar_pp_pendente()
    pausar(PAUSA_APOS_SALVAR)
    pagina.get_by_role("link", name="Limpar a Tela").click()


def solicitar_parada():
    # O teclado pode repetir o evento enquanto ESC estiver pressionado.
    if not base.robo_deve_parar:
        base.robo_deve_parar = True
        print("\n!!! TECLA ESC ACIONADA! ENCERRANDO AUTOMAÇÃO COM SEGURANÇA !!!")


def executar():
    planilhas, hotkey = None, None
    total = concluidos = 0
    base.robo_deve_parar = False
    base.PLANILHA = str(PLANILHA)
    try:
        if not base.mostrar_inicio("pp"):
            print("[CANCELADO] Início cancelado. Nenhuma operação realizada.")
            return
        planilhas = base.abrir_planilha()
        recuperar_pp_pendente(planilhas)
        registros = pendentes(planilhas)
        total = len(registros)
        if not registros:
            print("Nenhuma PP pendente com NL válida.")
            base.pyautogui.alert(
                text="Encerrei por aqui.",
                title="Fim",
                button="OK",
            )
            return
        hotkey = base.keyboard.add_hotkey("Esc", solicitar_parada)
        print("Robô iniciado. Pressione 'Esc' a qualquer momento para abortar com segurança.")
        with base.sync_playwright() as p:
            browser = p.chromium.connect_over_cdp(base.CHROME_DEBUG_URL, slow_mo=ATRASO_ACAO_MS)
            guias = [g for c in browser.contexts for g in c.pages if "SIGEF" in g.url.upper()]
            candidatas = [(g, f) for g in guias for f in g.frames
                          if f.get_by_placeholder("Pesquisar funcionalidades do sistema...").count() == 1]
            if len(candidatas) != 1:
                raise base.ErroAutomacao("Não foi localizado um único menu do SIGEF. Confira as guias abertas.")
            guia, frame = candidatas[0]
            pagina = base.abrir_funcionalidade(frame, guia, "PP Despesa Empenhada")
            for indice, r in enumerate(registros, start=1):
                base.verificar_panico()
                print(f"\n[{indice}/{total}] Linha {r['linha']} | {r['nome']}", flush=True)
                base.exibir_dados_registro(r, "preparacao")
                preparar_linha(pagina, planilhas, r)
                concluidos += 1
        print("Preparações concluídas.")
        base.pyautogui.alert(
            text="Encerrei por aqui.",
            title="Fim",
            button="OK",
        )
    except KeyboardInterrupt:
        print("\n[ENCERRADO] Tecla ESC acionada. A planilha será fechada com segurança.")
    except Exception as exc:
        # Durante uma espera, o navegador pode fechar o popup com ESC e
        # responder antes da próxima chamada de verificar_panico().
        if base.robo_deve_parar and isinstance(exc, PlaywrightError):
            print("\n[ENCERRADO] Tecla ESC acionada. A planilha será fechada com segurança.")
        else:
            print(f"[ERRO] {exc}\nConfira eventuais registros PESQUISAR_PP antes de repetir.")
            raise
    finally:
        if hotkey is not None:
            base.keyboard.remove_hotkey(hotkey)
        if planilhas:
            planilhas["book"].close()
        print(f"\n[RESUMO] Preparações concluídas nesta etapa: {concluidos} de {total}")
        if total > concluidos:
            print("  Se houve interrupção, confira a última linha antes de retomar.")
        print("[FINALIZAÇÃO] Recursos liberados com segurança.")


if __name__ == "__main__":
    executar()
