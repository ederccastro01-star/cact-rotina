# -*- coding: utf-8 -*-
"""Parsing puro dos e-mails de Push/Jusbrasil (spec §9c). Corpos sintéticos: nomes e números fictícios."""
import os
from datetime import datetime

import pytest

from pc import push_fontes as pf
from tests.conftest import REAIS

PJE = """________________________________
De: Processo Judicial Eletrônico - Push <nao-responda@trt18.jus.br>
Enviado: sexta-feira, 16 de outubro de 2026 05:41:31 (UTC-03:00) Brasilia
Para: fulano@hotmail.com <fulano@hotmail.com>
Assunto: [TRT18] [PUSH] Atualizações de Informações Processuais do Processo 0000001-11.2026.5.18.0011

Processo Judicial Eletrônico

Tribunal Regional do Trabalho da 18ª Região - 1º Grau

Acompanhamento de Atualizações de Informações Processuais

Prezado(a) FULANO DE TAL,

Informamos que o processo a seguir sofreu movimentação:

Número do Processo: 0000001-11.2026.5.18.0011

Classe Judicial: Cumprimento de Sentença de Ações Coletivas

Órgão Julgador: 11ª VARA DO TRABALHO DE GOIÂNIA

Data de Autuação: 17/05/2026 11:17:44

Autor: BELTRANO DA SILVA

Advogados do Autor: FULANO DE TAL, OAB: 11111

Réu: EMPRESA EXEMPLO S/A

Advogados do Réu: SICRANO, OAB: 22222

Eventos:

Data Evento
15/10/2026 23:40 Expedido(a) intimação a(o) BELTRANO DA SILVA
15/10/2026 23:40 Documento sigiloso
15/10/2026 23:39 Disponibilizado
(a) o(a) intimação no Diário da Justiça Eletrônico do dia 15/10/2026

Para acessar este processo na consulta pública, clique em https://pje.trt18.jus.br/consultaprocessual/.

ATENÇÃO: este e-mail é gerado de forma automatizada, por gentileza, não o responda.

AVENIDA T1, QUADRA T22, LOTE 01/24, N. 1698, SETOR BUENO, GOIÂNIA-GO - CEP: 74215-901

08006206622
"""

PJE_OUTLOOK = """
Weliton Marques
parceiro@gmail.com

\tInício da mensagem encaminhada:

\tDe: Processo Judicial Eletrônico - Push <nao-responda@trt18.jus.br>
\t
\tAssunto: [TRT18] [PUSH] Atualizações de Informações Processuais do Processo 0000002-22.2024.5.18.0008
\t
\tTribunal Regional do Trabalho da 18ª Região - 2º Grau

\tNúmero do Processo: 0000002-22.2024.5.18.0008

\tClasse Judicial: Ação Trabalhista - Rito Ordinário

\tÓrgão Julgador: JUÍZO DE EXECUÇÃO

\tData de Autuação: 02/09/2024 13:15:37

\tAutor: JOSE EXEMPLO

\tAdvogados do Autor: PARCEIRO, OAB: 21877

\tRéu: MUNICIPIO EXEMPLO e outros (1)

\tEventos:

Data\tEvento\t
16/10/2026 16:12\tExpedido(a) intimação a(o) MUNICIPIO EXEMPLO\t
16/10/2026 16:11\tSentença - Sentença\t

\tPara acessar este processo na consulta pública, clique em https://pje.trt18.jus.br/consultaprocessual/ <https://x>
"""

LEGADO = """________________________________
De: trtpush@trt18.jus.br <trtpush@trt18.jus.br>
Assunto: TRT 18 - Push - Andamento Processual

Sr(a) FULANO,

Encaminhamos abaixo os andamentos processuais referentes aos processos cadastrados por V.S�. no sistema TRT-Push.

Processo: ATOrd-0000003-33.2016.5.18.0015<http://sistemas.trt18.jus.br/x?p_num_pje=1&p_grau_pje=1&> Ver todos Andamentos >><http://sistemas.trt18.jus.br/y>
AUTOR: ALAN EXEMPLO
RÉU: JOSE EXEMPLO + 3
2 Andamento(s)

Data de Autuação Tramitação
16/10/2026 10:21:46 Expedido(a) intimação a(o) HELOIZA EXEMPLO
16/10/2026 09:10:37 Decisão | Decisão <http://sistemas.trt18.jus.br/visualizador/z>

Ocorreu um erro ou voc� n�o tem permiss�o para acessar este processo.
Processo: ATOrd-0000004-44.2014.5.18.0016<http://sistemas.trt18.jus.br/w> Ver todos Andamentos >><http://sistemas.trt18.jus.br/v>
AUTOR: SERGIO EXEMPLO
RÉU: BANCO EXEMPLO + 17
1 Andamento(s)

Data de Autuação Tramitação
16/10/2026 05:49:10 Publicado(a) o(a) intima��o em 16/10/2026

Obs.: Informa��es dispon�veis em tempo real e sujeitas a altera��es at� o final do dia.

Tribunal Regional do Trabalho - 18� Regi�o
"""

JUS_ANTIGO = ("Jusbrasil\n        Eder, os processos que você acompanha possuem novas informações! Os termos que você acompanha possuem 2 "
              "publicações recentes nos Diários Oficiais. Confira o que há de novo: Processo 0000005-55.2024.5.18.0014 Termos encontrados "
              "23147GO FULANO Data de disponibilização 15/10/26 Data de publicação 16/10/26 Diário TRT18 · TRT da 18ª Região Publicação "
              "COORDENADORIA DA 1ª TURMA JULGADORA Pauta Pauta de Julgamento PAUTA DA 36ª SESSÃO... Analisar publicação no Jus IA ➔ Acessar publicação "
              "Processo 0000006-66.2015.5.18.0010 Termos encontrados 23147GO FULANO Data de disponibilização 15/10/26 Data de publicação 16/10/26 "
              "Diário TJGO · Tribunal de Justiça de Goiás Publicação 2ª VARA CÍVEL NÚMERO ÚNICO: 0000006-66.2015.5.18.0010 Intimação da parte autora... "
              "Analisar publicação no Jus IA ➔ Acessar publicação Atenciosamente, Equipe Jusbrasil Este e-mail foi enviado para Eder (x@gmail.com) © 2026, Jusbrasil.")

JUS_NOVO = """Veja as publicações e comece a trabalhar em cada caso com o Jus IA.
 <https://na01.safelinks.protection.outlook.com/?url=x>
Jus IA Monitora\t

Movimentação recente


Eder, 4 novas publicações no seu nome


OAB/GO 23.147  Publicações dos Diários Oficiais de 19/10

TRT18 · 10ª VARA DO TRABALHO DE GOIÂNIA\t 19/10/26\t

Notificação

Celmo Exemplo x Empresa Exemplo S/A · 0000007-77.2026.5.18.0010

10ª VARA DO TRABALHO DE GOIÂNIA NÚMERO ÚNICO: 0000007-77.2026.5.18.0010 POLO ATIVO CELMO EXEMPLO POLO PASSIVO EMPRESA EXEMPLO S/A ADVOGADO(A/S) FULANO…

Ler publicação na íntegra

Jus IA ● pronto para começar a trabalhar

Com base no histórico do processo 0000007-77.2026.5.18.0010, liste os próximos passos a partir desta notificação: o que preciso fazer, em que prazo e qual peça cabe apresentar.

Próximos passos →\t

TRT18 · COORDENADORIA DA 1ª TURMA JULGADORA Pauta\t 16/10/26\t

Publicação

0000008-88.2024.5.18.0014

COORDENADORIA DA 1ª TURMA JULGADORA Pauta Pauta de Julgamento PAUTA DA 36ª SESSÃO…

Ler publicação na íntegra

Próximos passos →\t

Mais 2 publicações na sua Caixa de entrada

<https://na01.safelinks.protection.outlook.com/?url=y>

Notificação · TRT18 · 10ª VARA DO TRABALHO DE GOIÂNIA · 19/10/26

Celmo Exemplo x Empresa Exemplo S/A · 0000007-77.2026.5.18.0010

— Equipe Jusbrasil
"""

PROJUDI = """Sr(a). FULANO,

O Poder Judiciário do Estado de Goiás informa que está disponível uma intimação/citação referente ao processo nº. 5312551.73 em sua caixa de entrada no Sistema PROJUDI.

Sistema PROJUDI
"""


def _email(remetente, assunto, corpo, data=datetime(2026, 10, 16, 8, 52), mid="<a@b>"):
    return pf.Email(remetente=remetente, assunto=assunto, corpo=corpo, data=data, message_id=mid, conta="x", pasta="Caixa de Entrada")


# --------------------------------------------------------------------------- remetente real e origem
def test_remetente_real_vem_do_corpo_encaminhado():
    assert pf.remetente_real("ederccastro@hotmail.com", PJE) == "nao-responda@trt18.jus.br"
    assert pf.remetente_real("parceiro@gmail.com", PJE_OUTLOOK) == "nao-responda@trt18.jus.br"
    assert pf.remetente_real("publicacoes-diarios@jusbrasil.com.br", JUS_ANTIGO) == "publicacoes-diarios@jusbrasil.com.br"


def test_origem_por_remetente_real_e_corpo():
    assert pf.origem("nao-responda@trt18.jus.br", "x", PJE) == "PJe Push"
    assert pf.origem("push-nao-responda@mail.trt10.jus.br", "x", "") == "PJe Push"
    assert pf.origem("trtpush@trt18.jus.br", "x", LEGADO) == "TRT-Push (legado)"
    assert pf.origem("publicacoes-diarios@jusbrasil.com.br", "2 novas publicações", JUS_ANTIGO) == "Jusbrasil"
    assert pf.origem("sistema-projudi@tjgo.jus.br", "[PROJUDI]Informação de intimação/citação", PROJUDI) == "Outros"


def test_marketing_do_jusbrasil_nao_e_elegivel():
    assert pf.elegivel(_email("nao-responda@jusbrasil.com.br", "Conheça o Jus IA", "promo")) is False
    assert pf.elegivel(_email("ola@descubra.jusbrasil.com.br", "Suas publicações agora têm um lugar", "promo")) is False
    assert pf.elegivel(_email("publicacoes-diarios@jusbrasil.com.br", "4 novas publicações no seu nome", JUS_NOVO)) is True
    assert pf.elegivel(_email("ederccastro@hotmail.com", "ENC: [TRT18] [PUSH] Atualizações", PJE)) is True
    assert pf.elegivel(_email("ederccastro@hotmail.com", "RE: contrato", "De: kelly@escritorio.com.br <kelly@escritorio.com.br>\noi")) is False
    assert pf.elegivel(_email("sistema-projudi@tjgo.jus.br", "[PROJUDI]Informação de intimação/citação", PROJUDI)) is True


# --------------------------------------------------------------------------- PJe Push
def test_pje_uma_linha_por_evento_com_cabecalho():
    regs = pf.extrair(_email("ederccastro@hotmail.com", "ENC: [TRT18] [PUSH] Atualizações ... 0000001-11.2026.5.18.0011", PJE))
    assert len(regs) == 3
    r = regs[0]
    assert r.processo == "0000001-11.2026.5.18.0011" and r.tribunal == "TRT18" and r.ramo == "Justiça do Trabalho" and r.grau == "1º Grau"
    assert r.classe == "Cumprimento de Sentença de Ações Coletivas" and r.orgao == "11ª VARA DO TRABALHO DE GOIÂNIA"
    assert r.autuacao == datetime(2026, 5, 17, 11, 17, 44) and r.autor == "BELTRANO DA SILVA" and r.reu == "EMPRESA EXEMPLO S/A"
    assert r.data == datetime(2026, 10, 15, 23, 40) and r.teor == "Expedido(a) intimação a(o) BELTRANO DA SILVA"
    assert r.tipo == "Intimação / Expedição de documento" and r.origem == "PJe Push"
    assert regs[1].tipo == "Documento sigiloso"
    assert regs[2].teor == "Disponibilizado (a) o(a) intimação no Diário da Justiça Eletrônico do dia 15/10/2026"   # evento quebrado em duas linhas


def test_pje_no_corpo_do_outlook_com_tabs_e_encaminhamento_de_parceiro():
    regs = pf.extrair(_email("parceiro@gmail.com", "Fwd: [TRT18] [PUSH] Atualizações ... 0000002-22.2024.5.18.0008", PJE_OUTLOOK))
    assert [r.teor for r in regs] == ["Expedido(a) intimação a(o) MUNICIPIO EXEMPLO", "Sentença - Sentença"]
    assert regs[0].grau == "2º Grau" and regs[0].orgao == "JUÍZO DE EXECUÇÃO" and regs[0].remetente == "nao-responda@trt18.jus.br"
    assert regs[1].tipo == "Sentença"


# --------------------------------------------------------------------------- TRT-Push legado
def test_legado_varios_processos_urls_removidas_acentos_corrigidos():
    regs = pf.extrair(_email("ederccastro@hotmail.com", "ENC: TRT 18 - Push - Andamento Processual", LEGADO))
    assert [r.processo for r in regs] == ["0000003-33.2016.5.18.0015", "0000003-33.2016.5.18.0015", "0000004-44.2014.5.18.0016"]
    assert regs[1].teor == "Decisão | Decisão" and "http" not in regs[1].teor
    assert regs[0].autor == "ALAN EXEMPLO" and regs[0].reu == "JOSE EXEMPLO + 3" and regs[0].grau == "1º Grau"
    assert regs[0].data == datetime(2026, 10, 16, 10, 21, 46) and regs[0].origem == "TRT-Push (legado)"
    assert regs[2].teor == "Publicado(a) o(a) intimação em 16/10/2026"          # acento inequívoco corrigido


def test_normalizar_acentos_so_quando_inequivoco():
    assert pf.normalizar_acentos("Informa��es dispon�veis at�") == "Informações disponíveis até"
    assert pf.normalizar_acentos("xyz�w") == "xyz�w"                    # desconhecido fica como está


# --------------------------------------------------------------------------- Jusbrasil
def test_jusbrasil_layout_antigo_linha_unica():
    regs = pf.extrair(_email("publicacoes-diarios@jusbrasil.com.br", "2 novas publicações encontradas para você", JUS_ANTIGO))
    assert len(regs) == 2
    assert regs[0].processo == "0000005-55.2024.5.18.0014" and regs[0].tribunal == "TRT18" and regs[0].data == datetime(2026, 10, 16)
    assert regs[0].teor.startswith("COORDENADORIA DA 1ª TURMA JULGADORA Pauta") and "Analisar publicação" not in regs[0].teor
    assert regs[1].tribunal == "TJGO" and regs[1].ramo == "Justiça Estadual" and regs[1].orgao == "2ª VARA CÍVEL"
    assert regs[0].tipo == "Pauta/Audiência" and regs[0].origem == "Jusbrasil"


def test_jusbrasil_layout_novo_em_blocos_e_conta_faltantes():
    em = _email("publicacoes-diarios@jusbrasil.com.br", "4 novas publicações no seu nome", JUS_NOVO)
    regs = pf.extrair(em)
    assert [r.processo for r in regs] == ["0000007-77.2026.5.18.0010", "0000008-88.2024.5.18.0014"]
    assert regs[0].data == datetime(2026, 10, 19) and regs[0].orgao == "10ª VARA DO TRABALHO DE GOIÂNIA" and regs[0].tribunal == "TRT18"
    assert regs[0].teor.startswith("10ª VARA DO TRABALHO DE GOIÂNIA NÚMERO ÚNICO") and "Ler publicação" not in regs[0].teor
    assert regs[0].autor == "Celmo Exemplo" and regs[0].reu == "Empresa Exemplo S/A"
    assert regs[1].data == datetime(2026, 10, 16) and regs[1].tipo == "Pauta/Audiência"
    assert pf.anunciadas(em.assunto) == 4 and pf.faltantes(em) == 2                   # "Mais 2 publicações na sua Caixa de entrada"


# --------------------------------------------------------------------------- outros e derivados
def test_outros_projudi_sem_cnj_mantem_numero_bruto():
    regs = pf.extrair(_email("sistema-projudi@tjgo.jus.br", "[PROJUDI]Informação de intimação/citação", PROJUDI))
    assert len(regs) == 1 and regs[0].processo == "5312551.73" and regs[0].tribunal == "TJGO" and regs[0].origem == "Outros"
    assert regs[0].tipo == "Intimação / Citação" and regs[0].data == datetime(2026, 10, 16, 8, 52)      # sem data no corpo → data do e-mail


def test_tipo_de_andamento_dois_primeiros_na_ordem():
    assert pf.tipo_andamento("Expedido(a) intimação a(o) X") == "Intimação / Expedição de documento"
    assert pf.tipo_andamento("Conclusos os autos para despacho") == "Despacho / Conclusão"
    assert pf.tipo_andamento("Reativado o processo") == "Outros"
    assert pf.tipo_andamento("Disponibilizado(a) o(a) Pauta de Julgamento") == "Pauta/Audiência / Publicação/Disponibilização"
    assert pf.tipo_andamento("POLO PASSIVO EQUATORIAL GOIAS DISTRIBUIDORA DE ENERGIA S/A") == "Outros"     # nome de parte não é distribuição
    assert pf.tipo_andamento("Distribuído por sorteio") == "Remessa/Distribuição" and pf.tipo_andamento("Redistribuição do feito") == "Remessa/Distribuição"


def test_ramo_tribunal_superior_antes_dos_demais():
    assert pf.ramo("TST") == "Tribunal Superior" and pf.ramo("TRT18") == "Justiça do Trabalho"
    assert pf.ramo("TRF1") == "Justiça Federal" and pf.ramo("TJGO") == "Justiça Estadual" and pf.ramo("CNJ") == "Outros"
    assert pf.sigla_tribunal("nao-responda@trtsp.jus.br", "[TRT02] [PUSH] x", "") == "TRT02"
    assert pf.sigla_tribunal("nao-responda@tst.jus.br", "[PUSH] x", "Tribunal Superior do Trabalho") == "TST"


def test_limpar_rodape():
    assert pf.limpar_rodape("Sentença\n\nPara acessar este processo na consulta pública, clique em x\n08006206622") == "Sentença"
    assert pf.limpar_rodape("teor ... Analisar publicação no Jus IA ➔ Acessar publicação") == "teor ..."


# --------------------------------------------------------------------------- fixtures reais (gitignored)
def _real(nome):
    c = os.path.join(REAIS, "push_emails", nome)
    if not os.path.exists(c):
        pytest.skip("fixture real ausente")
    txt = open(c, encoding="utf-8").read()
    cab, corpo = txt.split("=====BODY=====\n", 1)
    meta = dict(l.split(": ", 1) for l in cab.strip().splitlines() if ": " in l)
    return _email(meta["SENDER"], meta["SUBJECT"], corpo)


def test_reais_pje_gmail():
    regs = pf.extrair(_real("gmail_pje_0000894.txt"))
    assert len(regs) == 2 and regs[0].processo == "0000894-09.2026.5.18.0011" and regs[0].tipo == "Intimação / Expedição de documento"


def test_reais_legado_gmail():
    regs = pf.extrair(_real("gmail_trtpush_legado.txt"))
    assert len(regs) == 16 and len({r.processo for r in regs}) == 2 and all("http" not in r.teor for r in regs)
    assert regs[-1].teor.startswith("Disponibilizado(a) o(a) intimação no Diário")


def test_reais_jusbrasil_antigo_e_novo():
    assert len(pf.extrair(_real("gmail_jusbrasil_antigo.txt"))) == 2
    novo = pf.extrair(_real("jusbrasil_1.txt"))
    assert len(novo) == 3 and pf.faltantes(_real("jusbrasil_1.txt")) == 1


def test_reais_encaminhamento_parceiro_outlook():
    regs = pf.extrair(_real("fwd_welitonadvogado_1.txt"))
    assert len(regs) >= 5 and regs[0].processo.endswith(".5.18.0008") and regs[0].remetente == "nao-responda@trt18.jus.br"


def test_partes_do_jusbrasil_sem_links():
    teor = "2ª VARA NÚMERO ÚNICO: 0000006-66.2015.5.18.0010 POLO ATIVO MARCIO EXEMPLO POLO PASSIVO J J MOTORS EIRELI - ME <https://x.y/z> ADVOGADO(A/S) FULANO"
    assert pf._partes(teor) == ("MARCIO EXEMPLO", "J J MOTORS EIRELI - ME")


# --------------------------------------------------------------------------- e-mail de terceiro (parceiro) com número de processo
TERCEIRO = """Boa tarde, Dr. Eder!
A pedido do Dr. Luís Gustavo Nicoli, encaminho os seguintes autos para solicitação de orçamento referente à elaboração de cálculos trabalhistas:
Processo_0000477-40.2026.5.18.0081 (1).pdf <https://nicolisociedade.sharepoint.com/x>
Peço que nos informe o valor dos honorários para nos representar na fase de execução do referido processo.
Nosso prazo processual encerra-se em 28/10/2026.
Atenciosamente,
"""


def test_email_de_terceiro_com_cnj_no_assunto_e_elegivel_e_vira_registro():
    em = _email("lucas.resende@nicoliadvogados.adv.br", "Bianka Vitória da Silva x  Cooperativa Central de Proteção Patrimonial - 0000477-40.2026.5.18.0081", TERCEIRO,
                data=datetime(2026, 10, 16, 16, 24))
    assert pf.elegivel(em) is True and pf.origem(pf.remetente_real(em.remetente, em.corpo), em.assunto, em.corpo) == pf.ORIGEM_TERCEIRO
    regs = pf.extrair(em)
    assert len(regs) == 1
    r = regs[0]
    assert r.processo == "0000477-40.2026.5.18.0081" and r.tribunal == "TRT18" and r.origem == pf.ORIGEM_TERCEIRO
    assert r.data == datetime(2026, 10, 16, 16, 24) and r.tipo.startswith("Mensagem de terceiro")
    assert r.autor == "Bianka Vitória da Silva" and r.reu == "Cooperativa Central de Proteção Patrimonial"
    assert r.teor.startswith("Bianka Vitória da Silva x") and "28/10/2026" in r.teor and "sharepoint" not in r.teor


def test_email_comum_sem_cnj_continua_fora():
    assert pf.elegivel(_email("kelly@escritorio.com.br", "RE: contrato de honorários", "oi, segue minuta")) is False


def test_email_do_proprio_escritorio_nao_e_terceiro():
    """Resposta nossa numa conversa sobre o processo não é novidade: 18/10/2026 a primeira rodada gerou 18 linhas de ruído."""
    em = _email("ederccastro@hotmail.com", "RE: 0000941-04.2026.5.18.0004 EDILON x EQUATORIAL", "Boa tarde, segue a conta.\n", data=datetime(2026, 10, 16, 10, 0))
    assert pf.elegivel(em, proprios={"ederccastro@hotmail.com", "ederccastro01@gmail.com"}) is False
    assert pf.elegivel(em) is True                                                      # sem a lista, continua elegível


# --------------------------------------------------------------------------- parceiros: radar pelo remetente
PARCEIROS = {"@nicoliadvogados.adv.br", "welitonadvogado@gmail.com"}


def test_parceiro_sem_cnj_no_assunto_acha_o_numero_no_corpo_ou_no_anexo():
    em = _email("bruna.ferreira@nicoliadvogados.adv.br", "Orçamento de cálculos", "Dr. Eder, segue o processo 0000591-48.2023.5.06.0006 para orçamento.", data=datetime(2026, 10, 16, 9, 0))
    assert pf.elegivel(em, parceiros=PARCEIROS) is True and pf.elegivel(em) is False
    r = pf.extrair(em, parceiros=PARCEIROS)[0]
    assert r.processo == "0000591-48.2023.5.06.0006" and r.tribunal == "TRT6" and r.origem == pf.ORIGEM_TERCEIRO and r.tipo.startswith("Mensagem de terceiro")
    em2 = pf.Email("karinne@nicoliadvogados.adv.br", "Novo caso", "Segue anexo para orçamento.", datetime(2026, 10, 16, 9, 0), "<m2>", "x", "Caixa de Entrada",
                   anexos=["Processo_0000477-40.2026.5.18.0081 (1).pdf"])
    assert pf.extrair(em2, parceiros=PARCEIROS)[0].processo == "0000477-40.2026.5.18.0081"


def test_parceiro_sem_numero_algum_entra_so_se_falar_de_servico():
    em = _email("welitonadvogado@gmail.com", "Pedido", "Boa tarde! Pode me passar um orçamento para a liquidação de um acordo?", data=datetime(2026, 10, 16, 9, 0))
    assert pf.elegivel(em, parceiros=PARCEIROS) is True
    r = pf.extrair(em, parceiros=PARCEIROS)[0]
    assert r.processo == "" and "sem nº de processo" in r.tipo and r.teor.startswith("Pedido — Boa tarde")
    social = _email("welitonadvogado@gmail.com", "Feliz aniversário!", "Parabéns, Dr. Eder!", data=datetime(2026, 10, 16, 9, 0))
    assert pf.elegivel(social, parceiros=PARCEIROS) is False
    estranho = _email("alguem@qualquer.com", "Orçamento", "quero um orçamento de cálculos", data=datetime(2026, 10, 16, 9, 0))
    assert pf.elegivel(estranho, parceiros=PARCEIROS) is False                      # não é parceiro e não tem CNJ no assunto
