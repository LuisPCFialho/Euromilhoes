#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera o Guia de Utilização do Gerador Euromilhões, em PDF.

  python scripts/gerar_guia_pdf.py

Escrito para quem nunca jogou Euromilhões e nunca ouviu falar de filtros nem
de estratégias de aposta. Tudo o que afirma em números vem dos dados da
própria aplicação (1980 sorteios, 1978 com dados de prémios) ou de contagem
exaustiva sobre as 2.118.760 combinações possíveis.

NOTA SOBRE CARACTERES: as fontes Helvetica do reportlab usam WinAnsi, que
cobre todos os acentos portugueses mas não cobre sinais como a seta, o menor
ou igual e o sigma. No fim, o script verifica que todo o texto é codificável
em cp1252 e recusa gravar se não for, para não saírem quadrados em vez de
letras.
"""
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, KeepTogether,
                                NextPageTemplate, PageBreak, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle)
from reportlab.platypus.tableofcontents import TableOfContents

SAIDA = Path(__file__).resolve().parent.parent / "Guia-Gerador-Euromilhoes.pdf"

TITULO = "Gerador Euromilhões"
SUBTITULO = "Guia de Utilização"
AUTOR = "Luís Fialho"
COLABORACAO = "João Fialho"
ANO = "2026"
DATA_GUIA = "12 de Novembro de 2026"

# ── Paleta, a mesma da aplicação ────────────────────────────────────────────
AZUL_FUNDO = colors.HexColor("#1a1a2e")
AZUL_MEIO = colors.HexColor("#16213e")
AZUL_CLARO = colors.HexColor("#74c0fc")
DOURADO = colors.HexColor("#c9a227")
DOURADO_CLARO = colors.HexColor("#ffd700")
VERDE = colors.HexColor("#1f7a3d")
VERDE_CLARO = colors.HexColor("#e8f5ec")
VERMELHO = colors.HexColor("#b3331a")
CINZA = colors.HexColor("#5a5a6a")
CINZA_CLARO = colors.HexColor("#f2f3f7")
CINZA_LINHA = colors.HexColor("#d5d7e0")
TEXTO = colors.HexColor("#24242e")

LARGURA = 170 * mm
base = getSampleStyleSheet()

E = {}
E["corpo"] = ParagraphStyle("corpo", parent=base["Normal"], fontName="Helvetica",
                            fontSize=9.6, leading=14.2, textColor=TEXTO,
                            alignment=TA_JUSTIFY, spaceAfter=7)
E["corpo_pequeno"] = ParagraphStyle("corpo_pequeno", parent=E["corpo"],
                                    fontSize=8.6, leading=12.4, spaceAfter=5)
E["h1"] = ParagraphStyle("h1", parent=base["Normal"], fontName="Helvetica-Bold",
                         fontSize=16, leading=20, textColor=colors.white,
                         backColor=AZUL_MEIO, borderPadding=(7, 8, 7, 8),
                         spaceBefore=4, spaceAfter=13)
# keepWithNext impede títulos órfãos no fundo da página.
E["h2"] = ParagraphStyle("h2", parent=base["Normal"], fontName="Helvetica-Bold",
                         fontSize=11.6, leading=15, textColor=AZUL_MEIO,
                         spaceBefore=11, spaceAfter=6, keepWithNext=1)
E["legenda"] = ParagraphStyle("legenda", parent=base["Normal"],
                              fontName="Helvetica-Oblique", fontSize=8,
                              leading=11, textColor=CINZA, spaceAfter=10)
E["cel"] = ParagraphStyle("cel", parent=base["Normal"], fontName="Helvetica",
                          fontSize=8.1, leading=10.6, textColor=TEXTO)
E["cel_cab"] = ParagraphStyle("cel_cab", parent=E["cel"],
                              fontName="Helvetica-Bold", textColor=colors.white)
E["quadro_t"] = ParagraphStyle("quadro_t", parent=base["Normal"],
                               fontName="Helvetica-Bold", fontSize=9.4,
                               leading=12.5, textColor=AZUL_MEIO, spaceAfter=4)
E["quadro_c"] = ParagraphStyle("quadro_c", parent=base["Normal"],
                               fontName="Helvetica", fontSize=8.8, leading=12.6,
                               textColor=TEXTO, alignment=TA_JUSTIFY)
# Entrelinha apertada de propósito: são 50 entradas e o índice tem de caber
# numa página só.
E["toc1"] = ParagraphStyle("toc1", parent=base["Normal"],
                           fontName="Helvetica-Bold", fontSize=9.5, leading=13.4,
                           textColor=AZUL_MEIO)
E["toc2"] = ParagraphStyle("toc2", parent=base["Normal"], fontName="Helvetica",
                           fontSize=8.6, leading=11.2, textColor=TEXTO,
                           leftIndent=14)


# ═══════════════════════════════════════════════════════════════════════════
# Blocos reutilizáveis
# ═══════════════════════════════════════════════════════════════════════════
def p(texto, estilo="corpo"):
    return Paragraph(texto, E[estilo])


def h1(texto):
    return Paragraph(texto, E["h1"])


def h2(texto):
    return Paragraph(texto, E["h2"])


def tabela(cabecalho, linhas, larguras=None, alinhar_dir=(), destaques=(),
           tamanho=8.1, cor_cab=AZUL_MEIO):
    """Tabela com cabeçalho escuro e linhas alternadas.

    alinhar_dir : índices de colunas a alinhar à direita (números)
    destaques   : índices de linhas de corpo a realçar
    """
    est = ParagraphStyle("c%s" % tamanho, parent=E["cel"], fontSize=tamanho,
                         leading=tamanho * 1.32)
    dados = [[Paragraph(c, E["cel_cab"]) for c in cabecalho]]
    for linha in linhas:
        dados.append([c if hasattr(c, "wrap") else Paragraph(str(c), est)
                      for c in linha])
    t = Table(dados, colWidths=larguras, repeatRows=1, hAlign="LEFT")
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), cor_cab),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, CINZA_LINHA),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]
    for i in range(1, len(dados)):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), CINZA_CLARO))
    for col in alinhar_dir:
        estilo.append(("ALIGN", (col, 1), (col, -1), "RIGHT"))
    for i in destaques:
        estilo.append(("BACKGROUND", (0, i + 1), (-1, i + 1),
                       colors.HexColor("#fff6d6")))
    t.setStyle(TableStyle(estilo))
    return t


def quadro(titulo, texto, cor=AZUL_CLARO, fundo=colors.HexColor("#eef5fd")):
    """Caixa de destaque, para as ideias que não devem passar ao lado."""
    t = Table([[Paragraph(titulo, E["quadro_t"])],
               [Paragraph(texto, E["quadro_c"])]],
              colWidths=[LARGURA - 10 * mm], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), fundo),
        ("LINEBEFORE", (0, 0), (0, -1), 2.6, cor),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("TOPPADDING", (0, 0), (0, 0), 8),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 8),
        ("TOPPADDING", (0, 1), (-1, 1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
    ]))
    return t


def aviso(titulo, texto):
    return quadro(titulo, texto, cor=VERMELHO, fundo=colors.HexColor("#fdf0ec"))


def bom(titulo, texto):
    return quadro(titulo, texto, cor=VERDE, fundo=VERDE_CLARO)


def lista(itens, estilo="corpo"):
    """Lista com marcas, como parágrafos recuados."""
    est = ParagraphStyle("li_" + estilo, parent=E[estilo], leftIndent=12,
                         bulletIndent=2, spaceAfter=3.5)
    saida = [Paragraph(i, est, bulletText="•") for i in itens]
    saida.append(Spacer(1, 5))
    return saida


def bolas(numeros, estrelas):
    """Desenha uma chave como 5 bolas e 2 estrelas."""
    est_n = ParagraphStyle("bn", parent=E["cel"], alignment=TA_CENTER,
                           fontSize=9.5, textColor=colors.white)
    est_e = ParagraphStyle("be", parent=E["cel"], alignment=TA_CENTER,
                           fontSize=9.5, textColor=colors.HexColor("#3a2d00"))
    cels = [Paragraph("<b>%d</b>" % n, est_n) for n in numeros]
    cels += [Paragraph("<b>%d</b>" % s, est_e) for s in estrelas]
    t = Table([cels], colWidths=[11 * mm] * 7, rowHeights=[8 * mm], hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (4, 0), AZUL_MEIO),
        ("BACKGROUND", (5, 0), (6, 0), DOURADO_CLARO),
        ("GRID", (0, 0), (-1, -1), 1.2, colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t


# ═══════════════════════════════════════════════════════════════════════════
# Capa e rodapé
# ═══════════════════════════════════════════════════════════════════════════
def desenhar_capa(canvas, doc):
    canvas.saveState()
    L, A = A4
    canvas.setFillColor(AZUL_FUNDO)
    canvas.rect(0, 0, L, A, fill=1, stroke=0)
    canvas.setFillColor(AZUL_MEIO)
    canvas.rect(0, A - 96 * mm, L, 96 * mm, fill=1, stroke=0)
    canvas.setFillColor(DOURADO_CLARO)
    canvas.rect(0, A - 97.6 * mm, L, 1.6 * mm, fill=1, stroke=0)

    canvas.setFillColor(DOURADO_CLARO)
    canvas.setFont("Helvetica-Bold", 10)
    canvas.drawCentredString(L / 2, A - 34 * mm,
                             "A P L I C A Ç Ã O   D E   A P O S T A S")

    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 42)
    canvas.drawCentredString(L / 2, A - 57 * mm, "GERADOR")
    canvas.setFont("Helvetica-Bold", 37)
    canvas.drawCentredString(L / 2, A - 76 * mm, "EUROMILHÕES")

    y = A - 142 * mm
    exemplo = [7, 18, 23, 37, 44]
    estrelas_ex = [3, 9]
    passo = 23 * mm
    todos = exemplo + estrelas_ex
    x0 = L / 2 - (len(todos) - 1) * passo / 2
    for i, v in enumerate(todos):
        x = x0 + i * passo
        canvas.setFillColor(colors.white if i < 5 else DOURADO_CLARO)
        canvas.circle(x, y, 9.2 * mm, fill=1, stroke=0)
        canvas.setFillColor(AZUL_FUNDO)
        canvas.setFont("Helvetica-Bold", 16)
        canvas.drawCentredString(x, y - 5.6, str(v))

    canvas.setFillColor(colors.HexColor("#c8c8da"))
    canvas.setFont("Helvetica", 18)
    canvas.drawCentredString(L / 2, y - 32 * mm, SUBTITULO)
    canvas.setFont("Helvetica-Oblique", 11.5)
    canvas.drawCentredString(L / 2, y - 44 * mm,
                             "Como escolher chaves, e o que os números dizem sobre isso")
    canvas.setFont("Helvetica", 10)
    canvas.setFillColor(colors.HexColor("#9a9ab0"))
    canvas.drawCentredString(L / 2, y - 55 * mm,
                             "Para quem nunca jogou e para quem joga há anos")

    canvas.setStrokeColor(DOURADO)
    canvas.setLineWidth(0.8)
    canvas.line(L / 2 - 26 * mm, 70 * mm, L / 2 + 26 * mm, 70 * mm)

    canvas.setFillColor(colors.white)
    canvas.setFont("Helvetica-Bold", 14)
    canvas.drawCentredString(L / 2, 58 * mm, AUTOR)
    canvas.setFont("Helvetica", 9)
    canvas.setFillColor(colors.HexColor("#9a9ab0"))
    canvas.drawCentredString(L / 2, 50 * mm, "com colaboração de %s" % COLABORACAO)
    canvas.setFont("Helvetica-Bold", 12)
    canvas.setFillColor(DOURADO_CLARO)
    canvas.drawCentredString(L / 2, 38 * mm, ANO)
    canvas.restoreState()


def desenhar_rodape(canvas, doc):
    canvas.saveState()
    L, _ = A4
    canvas.setStrokeColor(CINZA_LINHA)
    canvas.setLineWidth(0.5)
    canvas.line(20 * mm, 16 * mm, L - 20 * mm, 16 * mm)
    canvas.setFont("Helvetica", 7.6)
    canvas.setFillColor(CINZA)
    canvas.drawString(20 * mm, 11.5 * mm, "%s — %s" % (TITULO, SUBTITULO))
    canvas.drawRightString(L - 20 * mm, 11.5 * mm, "%d" % canvas.getPageNumber())
    canvas.restoreState()


class Guia(BaseDocTemplate):
    """Documento com capa própria e índice com números de página."""

    def __init__(self, caminho):
        super().__init__(caminho, pagesize=A4,
                         leftMargin=20 * mm, rightMargin=20 * mm,
                         topMargin=20 * mm, bottomMargin=22 * mm,
                         title="%s — %s" % (TITULO, SUBTITULO), author=AUTOR,
                         subject="Guia de utilização da aplicação Gerador Euromilhões")
        self.addPageTemplates([
            # A capa é toda desenhada no canvas; esta moldura só recebe o
            # espaçador que antecede a quebra de página. Sem padding, senão o
            # espaçador não cabe e o reportlab abre uma página em branco.
            PageTemplate(id="capa",
                         frames=[Frame(20 * mm, 22 * mm, LARGURA, 20 * mm,
                                       id="capa", leftPadding=0, rightPadding=0,
                                       topPadding=0, bottomPadding=0)],
                         onPage=desenhar_capa),
            PageTemplate(id="normal",
                         frames=[Frame(20 * mm, 22 * mm, LARGURA,
                                       A4[1] - 42 * mm, id="texto")],
                         onPage=desenhar_rodape),
        ])

    def afterFlowable(self, flowable):
        """Alimenta o índice: cada título anuncia a página onde caiu."""
        if not isinstance(flowable, Paragraph):
            return
        if flowable.style.name == "h1":
            self.notify("TOCEntry", (0, flowable.getPlainText(), self.page))
        elif flowable.style.name == "h2":
            self.notify("TOCEntry", (1, flowable.getPlainText(), self.page))


# ═══════════════════════════════════════════════════════════════════════════
# CONTEÚDO
# ═══════════════════════════════════════════════════════════════════════════
def construir():
    s = []

    def add(bloco):
        """Aceita um flowable ou uma lista deles."""
        if isinstance(bloco, list):
            s.extend(bloco)
        else:
            s.append(bloco)

    # ── capa ────────────────────────────────────────────────────────────────
    add(NextPageTemplate("normal"))
    add(Spacer(1, 0.2))
    add(PageBreak())

    # ── índice ──────────────────────────────────────────────────────────────
    add(Paragraph("Índice", ParagraphStyle(
        "idx", parent=E["h1"], backColor=None, textColor=AZUL_MEIO,
        borderPadding=0, fontSize=20, spaceAfter=14)))
    toc = TableOfContents()
    toc.levelStyles = [E["toc1"], E["toc2"]]
    toc.dotsMinLevel = 0
    add(toc)
    add(PageBreak())

    # ═══════════════════════════════════════════════════════════════════════
    add(h1("1. Comece por aqui"))
    add(p("Este guia explica uma aplicação que gera chaves do Euromilhões. Não é "
          "preciso saber nada do jogo para o ler: o capítulo 2 explica as regras, "
          "o capítulo 3 explica o que a aplicação pode e não pode fazer, e os "
          "restantes mostram, ecrã por ecrã, como se usa."))
    add(p("Há uma coisa que vale a pena dizer logo, porque é a razão de ser de tudo "
          "o que vem a seguir — e porque é o contrário do que costuma prometer-se "
          "neste tipo de programas."))
    add(quadro(
        "O que esta aplicação não faz",
        "<b>Não aumenta a probabilidade de acertar.</b> Nenhuma aplicação o pode "
        "fazer. As bolas do Euromilhões não têm memória: a probabilidade de uma "
        "chave sair é exactamente a mesma para todas as 2.118.760 combinações, "
        "quer tenham sido escolhidas com todo o cuidado, quer tenham sido escritas "
        "ao calhas. Quem lhe disser o contrário está enganado, ou está a enganá-lo."))
    add(bom(
        "O que esta aplicação faz, e que é real",
        "<b>Reduz a probabilidade de partilhar o prémio</b> no dia em que acertar. "
        "Milhões de pessoas jogam datas de aniversário, sequências e desenhos no "
        "boletim. Se a sua chave for parecida com a dessas pessoas, quando acertar "
        "vai dividir o prémio com muita gente; se for diferente, fica com uma fatia "
        "maior. Isto não é opinião: está medido nos prémios de 1978 sorteios, e os "
        "números estão no capítulo 3.<br/><br/>"
        "<b>E organiza o trabalho:</b> guarda as chaves, confere-as automaticamente "
        "contra o resultado oficial, mantém o histórico completo dos sorteios desde "
        "2004 e exporta tudo para Excel."))
    add(p("Os dois efeitos têm tamanhos muito diferentes. O primeiro é zero. O "
          "segundo é pequeno, mas existe. Este guia procura ser claro sobre qual é "
          "qual — inclusive quando isso obriga a dizer que uma funcionalidade da "
          "própria aplicação não serve para aquilo que parece servir."))

    add(h2("Como ler este guia"))
    add(tabela(
        ["A sua situação", "Leia"],
        [["Nunca jogou Euromilhões",
          "Capítulos 2, 3 e 4. Depois experimente a aplicação."],
         ["Já joga e quer usar a aplicação",
          "Capítulo 5 (gerar chaves) e capítulo 6 (os filtros)."],
         ["Quer saber se isto funciona",
          "Capítulo 3. É o capítulo dos números, e não poupa ninguém."],
         ["Quer saber quanto vale uma aposta",
          "Capítulo 10. Talvez seja a informação mais útil do programa."]],
        larguras=[60 * mm, LARGURA - 60 * mm]))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("2. O Euromilhões em três minutos"))
    add(p("Uma aposta é composta por <b>5 números entre 1 e 50</b> e <b>2 estrelas "
          "entre 1 e 12</b>. Custa 2,50 €. Há sorteio à terça e à sexta-feira."))
    add(Spacer(1, 3))
    add(bolas([7, 18, 23, 37, 44], [3, 9]))
    add(p("Uma aposta: cinco números (a azul) e duas estrelas (a dourado).", "legenda"))
    add(p("Ganha-se prémio acertando pelo menos 2 números, ou 1 número e as 2 "
          "estrelas. Há 13 níveis de prémio, a que se chama <b>escalões</b>. A "
          "tabela mostra-os todos, com a probabilidade de cada um e o prémio típico "
          "— a mediana efectivamente paga nos 1039 sorteios desde Setembro de 2016, "
          "quando as estrelas passaram a ser 12."))
    add(tabela(
        ["Escalão", "O que acertou", "Probabilidade", "Prémio típico"],
        [["1", "5 números + 2 estrelas (jackpot)", "1 em 139.838.160", "acumulado"],
         ["2", "5 números + 1 estrela", "1 em 6.991.908", "259.068 €"],
         ["3", "5 números", "1 em 3.107.515", "27.759 €"],
         ["4", "4 números + 2 estrelas", "1 em 621.503", "1.711 €"],
         ["5", "4 números + 1 estrela", "1 em 31.075", "139 €"],
         ["6", "3 números + 2 estrelas", "1 em 14.125", "69 €"],
         ["7", "4 números", "1 em 13.811", "46 €"],
         ["8", "2 números + 2 estrelas", "1 em 985", "15,73 €"],
         ["9", "3 números + 1 estrela", "1 em 706", "12,47 €"],
         ["10", "3 números", "1 em 314", "10,37 €"],
         ["11", "1 número + 2 estrelas", "1 em 188", "7,69 €"],
         ["12", "2 números + 1 estrela", "1 em 49", "6,03 €"],
         ["13", "2 números", "1 em 22", "4,15 €"]],
        larguras=[17 * mm, 62 * mm, 44 * mm, LARGURA - 123 * mm],
        alinhar_dir=(2, 3), destaques=(0,)))
    add(p("Somando os treze escalões, a probabilidade de ganhar <i>alguma coisa</i> "
          "é de 1 em 13. A de ganhar o jackpot é de 1 em 139.838.160.", "legenda"))

    add(h2("Quantas combinações existem"))
    add(p("Os cinco números podem sair de 2.118.760 maneiras diferentes. Este número "
          "aparece em todo o guia, porque é o universo onde a aplicação trabalha: "
          "quando se diz que um filtro rejeita 12%, quer dizer que rejeita 12% "
          "destas 2.118.760 combinações."))
    add(tabela(
        ["", "Combinações", "De onde vem"],
        [["5 números de 50", "2.118.760", "50 x 49 x 48 x 47 x 46, a dividir por 120"],
         ["2 estrelas de 12", "66", "12 x 11, a dividir por 2"],
         ["Aposta completa", "139.838.160", "2.118.760 x 66"]],
        larguras=[40 * mm, 32 * mm, LARGURA - 72 * mm], alinhar_dir=(1,)))
    add(quadro(
        "Uma imagem para o 1 em 139.838.160",
        "Se comprasse uma aposta por segundo, sem parar para dormir, levaria mais de "
        "<b>quatro anos e meio</b> a cobrir todas as combinações possíveis. "
        "Comprando duas apostas por semana, levaria 1,3 milhões de anos."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("3. O capítulo dos números: o que está medido"))
    add(p("Este é o capítulo mais importante do guia, e é o único onde a aplicação "
          "admite aquilo que não consegue fazer. Se só ler um capítulo, leia este."))

    add(h2("Porque é que a partilha se consegue medir"))
    add(p("Quanto se recebe de um prémio depende de quantas pessoas acertaram o "
          "mesmo. O número de vencedores de um escalão depende de três coisas: "
          "quantos bilhetes foram vendidos, a probabilidade do escalão, e <b>quão "
          "popular era aquela combinação</b>. É esta terceira que interessa."))
    add(p("As vendas variam muito de sorteio para sorteio, mas — e é isto que torna "
          "a medição possível — <b>as vendas são independentes dos números que "
          "saem</b>, porque os números saem ao acaso. O próprio sorteio faz a "
          "distribuição aleatória por grupos, o que elimina o enviesamento. "
          "Comparar os sorteios em que saíram números baixos com aqueles em que "
          "saíram números altos mede directamente quais são mais jogados."))
    add(p("A tabela resume o que saiu desse cruzamento, feito sobre <b>1978 dos 1980 "
          "sorteios</b> (13-02-2004 a 10-11-2026; faltam dois que a fonte não "
          "publica). Um valor positivo significa mais co-vencedores, ou seja, pior "
          "para quem acerta."))
    add(tabela(
        ["Característica do sorteio", "Efeito nos<br/>co-vencedores",
         "Intervalo de<br/>confiança 95%", "Amostra"],
        [["Cada número <b>até 31</b> na chave (escalão 2+0)",
          "+1,24%", "+0,19% a +2,29%", "1601 sorteios"],
         ["Cada número <b>até 31</b> na chave (escalão 3+0)",
          "+1,30%", "+0,00% a +2,59%", "1978 sorteios"],
         ["<b>Chave inteira de datas</b> (os 5 números até 31)",
          "+4,14%", "+0,02% a +8,27%", "130 sorteios"],
         ["Soma dos números entre 180 e 209",
          "-5,9%", "-11,8% a 0,0%", "1601 sorteios"],
         ["Dois pares de números consecutivos",
          "+2,8%", "-2,7% a +8,3%", "1601 sorteios"],
         ["Números até 12 (os meses)", "sem efeito", "—", "1978 sorteios"],
         ["Chave num dos 16 padrões equilibrados",
          "-0,32%", "-2,34% a +1,70%", "1601 sorteios"]],
        larguras=[71 * mm, 25 * mm, 36 * mm, LARGURA - 132 * mm],
        alinhar_dir=(1, 2, 3), destaques=(0, 1, 2)))
    add(p("Quando o intervalo de confiança inclui o zero, o efeito não se distingue "
          "do acaso. Dos sete, só os três primeiros — todos sobre números até 31 — "
          "ficam claramente acima de zero.", "legenda"))

    add(h2("O aviso mais importante de todos"))
    add(aviso(
        "Não há números mais jogados que se consigam identificar",
        "A partir destes dados produz-se com facilidade uma lista de \"números mais "
        "jogados\" com um ar perfeitamente convincente — o 20, o 25, o 7. Essa "
        "lista foi produzida e depois testada: dividiram-se os sorteios em duas "
        "metades independentes e estimou-se tudo outra vez em cada uma.<br/><br/>"
        "<b>As duas metades não concordam.</b> A correlação entre elas é de -0,085 "
        "num escalão e +0,018 no outro, ou seja, zero. Se a lista medisse algo real, "
        "as duas metades teriam de dar aproximadamente o mesmo.<br/><br/>"
        "Qualquer lista de números preferidos — desta aplicação, de um jornal ou de "
        "um site — é ruído. O que se aguenta é apenas o contraste <b>agregado</b> "
        "entre 1-31 e 32-50, porque faz a média de 31 números contra 19 e o ruído "
        "cancela-se."))

    add(h2("O que isto dá, na prática"))
    add(p("Uma única recomendação sobrevive a todos os testes: <b>ter pelo menos um "
          "número acima de 31</b>. Elimina o pior caso, que são as chaves "
          "inteiramente feitas de datas de aniversário — as tais que atraem +4,1% de "
          "co-vencedores. Na aplicação é o filtro G, a Regra do 31, e vem ligado."))
    add(p("Exigir dois números acima de 31, em vez de um, baixaria a partilha "
          "esperada de -0,21% para -0,79% face a não filtrar nada: meio ponto "
          "percentual, dentro da margem de erro da própria medição. Chegou a ser "
          "experimentado e foi revertido, porque tornava impossíveis dois dos 16 "
          "padrões que a aplicação usa para construir as chaves — e não compensava."))

    add(h2("O jackpot é precisamente aquilo que não se consegue testar"))
    add(p("Nos 158 sorteios em que saíram apenas números até 31 houve 39 vencedores "
          "de jackpot, contra 43,5 esperados: um desvio de 0,68 erros-padrão, ou "
          "seja, nada. Mas isto não prova que não haja efeito. Com apenas 39 "
          "vencedores, só se detectaria um efeito acima de 30%."))
    add(p("O escalão que mais interessa é exactamente aquele onde 22 anos de dados "
          "não chegam. Os escalões baixos, com milhões de vencedores, são a melhor "
          "aproximação disponível — e é deles que vem toda a tabela acima."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("4. Primeiros passos"))
    add(h2("Entrar na aplicação"))
    add(p("A aplicação abre no endereço <b>euromilhoes-chi.vercel.app</b> e pede "
          "email e palavra-passe. Na primeira entrada a palavra-passe é <b>123</b>, "
          "e a aplicação obriga a definir uma nova antes de deixar entrar. A sessão "
          "fica guardada no navegador, por isso não é preciso repetir de cada vez."))

    add(h2("Os oito separadores"))
    add(p("O menu da esquerda tem oito secções. A tabela diz para que serve cada uma "
          "e em que capítulo deste guia está explicada."))
    add(tabela(
        ["Separador", "Para que serve", "Cap."],
        [["<b>Gerar Chaves</b>",
          "O ecrã principal: produz as chaves e aplica os filtros", "5"],
         ["<b>Último Sorteio</b>",
          "O resultado mais recente e a comparação com as suas chaves", "11"],
         ["<b>Estatísticas Históricas</b>",
          "Frequências, atrasos, somas e as tabelas dos filtros", "9"],
         ["<b>Estratégias e Filtros</b>",
          "Explicação dos quadrantes, dos padrões e de cada filtro", "6 a 8"],
         ["<b>Padrões BI-BP-AI-AP</b>",
          "A tabela completa dos 56 padrões combinatórios", "8"],
         ["<b>Histórico Padrões</b>", "Que padrão saiu em cada sorteio", "8"],
         ["<b>Histórico da Base de Dados</b>",
          "Os 1980 sorteios, agrupados por ano e por mês", "12"],
         ["<b>Actualizar Base de Dados</b>",
          "Importar, inserir à mão e recolher prémios", "12"]],
        larguras=[48 * mm, LARGURA - 62 * mm, 14 * mm]))

    add(h2("Dois detalhes do ecrã"))
    add(lista([
        "<b>Tema claro ou escuro:</b> o botão no canto superior direito troca entre "
        "os dois, e a escolha fica guardada para a próxima visita.",
        "<b>Próximo sorteio:</b> a barra lateral mostra sempre a data do próximo "
        "sorteio e quantas horas faltam.",
    ]))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("5. Gerar chaves"))
    add(p("É o ecrã onde se passa a maior parte do tempo. À esquerda ficam as "
          "opções e o botão de gerar; à direita aparecem as chaves produzidas."))

    add(h2("Passo a passo"))
    add(tabela(
        ["Passo", "O que fazer", "Nota"],
        [["1", "Escolher quantas chaves quer", "1 a 50 de cada vez"],
         ["2", "Deixar os filtros como estão, ou mexer neles", "Ver o capítulo 6"],
         ["3", "Se quiser, indicar números a incluir ou a excluir", "Opcional"],
         ["4", "Carregar em <b>Gerar Chaves</b>", "Demora menos de um segundo"],
         ["5", "Copiar, guardar como favorita ou exportar", "Ver mais abaixo"]],
        larguras=[15 * mm, 86 * mm, LARGURA - 101 * mm]))

    add(h2("Os três conjuntos de filtros pré-definidos"))
    add(p("Em vez de ligar e desligar filtros um a um, há três botões que o fazem de "
          "uma vez. A última coluna mostra quantas das 2.118.760 combinações "
          "sobrevivem a cada conjunto: é a melhor forma de ver o que cada um custa "
          "em variedade."))
    add(tabela(
        ["Conjunto", "Filtros ligados", "Chaves que sobrevivem", "Do universo"],
        [["<b>Todos</b>", "os nove (A a I)", "1.239.019", "58,5%"],
         ["<b>Conservador</b>", "A, B, D, F, G", "1.345.851", "63,5%"],
         ["<b>Agressivo</b>", "A e G", "1.799.854", "84,9%"],
         ["(nenhum)", "—", "2.118.760", "100%"]],
        larguras=[30 * mm, 46 * mm, 44 * mm, LARGURA - 120 * mm],
        alinhar_dir=(2, 3)))
    add(p("Os nomes enganam um pouco: \"Agressivo\" deixa passar <i>mais</i> chaves, "
          "e não menos. Refere-se a filtrar de forma agressivamente mínima.",
          "legenda"))

    add(h2("Incluir ou excluir números"))
    add(p("Pode obrigar um número a aparecer, ou impedi-lo de aparecer, em quantas "
          "chaves quiser. São quatro campos, cada um com um número e um intervalo "
          "de <b>mín</b> a <b>máx</b> chaves."))
    add(tabela(
        ["Se escrever", "Resultado"],
        [["<b>7</b> em <b>0</b> a <b>0</b> chaves", "O número 7 nunca aparece"],
         ["<b>7</b> com mín e máx em branco", "O número 7 aparece em todas as chaves"],
         ["<b>7</b> em <b>2</b> a <b>5</b> chaves",
          "O 7 aparece em 2 a 5 das chaves geradas"],
         ["<b>7</b> em <b>1</b> a <b>1</b> chave", "O 7 aparece exactamente numa chave"]],
        larguras=[64 * mm, LARGURA - 64 * mm]))
    add(aviso(
        "Isto não melhora as suas probabilidades",
        "Incluir o número da data de nascimento de um filho não muda nada na "
        "probabilidade de acertar — mas, se esse número for 31 ou menos, aumenta "
        "ligeiramente a probabilidade de partilhar o prémio, pelas razões do "
        "capítulo 3.<br/><br/>"
        "A funcionalidade existe porque muita gente gosta de ter os seus números nas "
        "chaves, e essa é uma razão perfeitamente legítima para jogar. Fica apenas "
        "dito o que custa, para a escolha ser informada."))

    add(h2("Chave ao calhas"))
    add(p("O botão <b>Lucky Pick</b> gera uma chave sem filtro nenhum, puramente ao "
          "acaso. Vale a pena dizê-lo com clareza: <b>a probabilidade de essa chave "
          "acertar é exactamente a mesma</b> das que passaram pelos nove filtros. A "
          "diferença está só na partilha, e é pequena."))

    add(h2("O que fazer com as chaves geradas"))
    add(tabela(
        ["Botão", "O que faz"],
        [["<b>Copiar</b>", "Põe as chaves na área de transferência, em texto"],
         ["<b>Partilhar</b>", "Abre o WhatsApp com a lista pronta a enviar"],
         ["<b>Guardar favorito</b>",
          "Guarda o conjunto com um nome, para o recuperar mais tarde"],
         ["<b>Exportar Excel</b>",
          "Ficheiro .xlsx com as chaves e as estatísticas da geração"],
         ["<b>Exportar CSV</b>",
          "O mesmo em texto simples, para outras folhas de cálculo"],
         ["<b>JSC</b>",
          "Envia as chaves para a aplicação AIJackpot (capítulo 13)"],
         ["<b>Comparar com sorteio</b>",
          "Mostra quantos números cada chave acertou no último sorteio"]],
        larguras=[46 * mm, LARGURA - 46 * mm]))
    add(p("Os favoritos ficam guardados na nuvem e não no computador, por isso estão "
          "disponíveis em qualquer dispositivo onde entre com a mesma conta.",
          "legenda"))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("6. Os nove filtros, um a um"))
    add(p("Um filtro é uma regra que recusa certas combinações. Todos se podem "
          "desligar. Ao lado de cada interruptor, a aplicação mostra <b>quanto esse "
          "filtro rejeita de facto</b> — uma informação que vale a pena olhar antes "
          "de decidir se o filtro importa."))

    add(h2("Os nove de uma vez"))
    add(p("A coluna <b>acertou</b> diz em quantos dos 1971 sorteios com dados "
          "completos a combinação sorteada teria passado o filtro. A coluna <b>ao "
          "acaso</b> diz quantos teriam passado se os sorteios fossem — como são — "
          "puramente aleatórios. A última coluna é a diferença entre as duas, em "
          "pontos percentuais."))
    add(tabela(
        ["", "Filtro", "Regra", "Rejeita", "Acertou", "Ao acaso", "Dif."],
        [["A", "Soma", "soma dos 5 entre 90 e 195", "12,6%", "88,13%", "87,44%", "+0,69"],
         ["B", "Consecutivos", "no máximo 1 par seguido e 0 triplos", "4,5%",
          "95,59%", "95,50%", "+0,09"],
         ["C", "Dígito Final", "no máximo 2 com o mesmo algarismo final", "4,8%",
          "95,64%", "95,22%", "+0,42"],
         ["D", "Dezenas", "pelo menos 3 dezenas diferentes", "7,1%",
          "93,25%", "92,86%", "+0,39"],
         ["E", "Anti-Repetição", "no máximo 2 números do sorteio anterior", "0,5%",
          "99,70%", "99,52%", "+0,18"],
         ["F", "Cores", "ver o capítulo 7", "19,5%", "78,23%", "79,03%", "-0,80"],
         ["G", "Regra do 31", "pelo menos 1 número acima de 31", "8,0%",
          "92,03%", "91,98%", "+0,05"],
         ["H", "Anti-Progressão", "recusa 5-10-15-20-25 e semelhantes", "0,014%",
          "100,00%", "99,99%", "+0,01"],
         ["I", "AC Value", "complexidade aritmética 4, 5 ou 6", "3,6%",
          "97,06%", "96,44%", "+0,62"]],
        larguras=[7 * mm, 25 * mm, 54 * mm, 16 * mm, 17 * mm, 17 * mm,
                  LARGURA - 136 * mm],
        alinhar_dir=(3, 4, 5, 6), tamanho=7.5))
    add(aviso(
        "Leia a última coluna com atenção",
        "Nenhum dos nove filtros se afasta do acaso mais de <b>0,8 pontos "
        "percentuais</b>. Em 1971 sorteios, isso é indistinguível de zero.<br/><br/>"
        "Quer dizer que <b>nenhum destes filtros capta qualquer padrão nos "
        "sorteios</b>. E isto não é um defeito da aplicação: é o que tem de "
        "acontecer, porque as bolas são sorteadas ao acaso. Um filtro que se "
        "afastasse muito do acaso seria sinal de que o sorteio não era honesto — ou, "
        "muito mais provavelmente, de que havia um erro nas contas.<br/><br/>"
        "Os filtros servem para <b>dar variedade às chaves e evitar as combinações "
        "que muita gente joga</b>. Não servem para prever."))

    add(h2("Filtro A: Soma"))
    add(p("Aceita apenas chaves cuja soma dos cinco números esteja entre 90 e 195. "
          "Cerca de 88% dos sorteios históricos caem neste intervalo. Há também a "
          "opção <b>Apertado</b>, que estreita o intervalo para 95-160."))
    add(tabela(
        ["Chave", "Soma", "Passa?"],
        [["1 - 2 - 3 - 4 - 5", "15", "Não, soma baixa demais"],
         ["3 - 11 - 24 - 29 - 41", "108", "Sim"],
         ["46 - 47 - 48 - 49 - 50", "240", "Não, soma alta demais"]],
        larguras=[52 * mm, 20 * mm, LARGURA - 72 * mm]))
    add(p("Nota curiosa: as somas que este filtro recusa por serem baixas têm, em "
          "média, 4,27 números até 31 — são chaves carregadas de datas. As que "
          "passou a aceitar no topo do intervalo têm 1,09. O filtro da soma acaba "
          "por empurrar na mesma direcção da Regra do 31.", "legenda"))

    add(h2("Filtro B: Consecutivos"))
    add(p("Permite no máximo um par de números seguidos e proíbe três seguidos. "
          "Pares acontecem em cerca de 42% dos sorteios reais; três seguidos em "
          "menos de 1%."))
    add(tabela(
        ["Chave", "Números seguidos", "Passa?"],
        [["4 - 12 - 13 - 27 - 44", "um par (12-13)", "Sim"],
         ["4 - 12 - 13 - 27 - 28", "dois pares", "Não"],
         ["4 - 12 - 13 - 14 - 27", "um triplo", "Não"]],
        larguras=[52 * mm, 36 * mm, LARGURA - 88 * mm]))
    add(p("Atenção a uma crença muito repetida: diz-se que os apostadores evitam "
          "números seguidos e que por isso convém jogá-los. <b>Os dados não "
          "confirmam isso.</b> Dois pares seguidos dão +2,8% de co-vencedores, com "
          "margem de erro de 2,8 pontos — se alguma coisa, ligeiramente <i>mais</i> "
          "jogados, e não menos.", "legenda"))

    add(h2("Filtro C: Dígito Final"))
    add(p("No máximo dois números com o mesmo algarismo final. Três ou mais "
          "representam menos de 4% dos sorteios."))
    add(tabela(
        ["Chave", "Algarismos finais", "Passa?"],
        [["3 - 13 - 28 - 34 - 45", "3, 3, 8, 4, 5", "Sim, só dois acabam em 3"],
         ["3 - 13 - 23 - 34 - 45", "3, 3, 3, 4, 5", "Não, três acabam em 3"]],
        larguras=[52 * mm, 38 * mm, LARGURA - 90 * mm]))

    add(h2("Filtro D: Dezenas"))
    add(p("Exige pelo menos três dezenas diferentes, sendo as dezenas 1-10, 11-20, "
          "21-30, 31-40 e 41-50. Evita chaves concentradas numa zona do boletim, que "
          "são visualmente atraentes e por isso mais jogadas."))
    add(tabela(
        ["Chave", "Dezenas usadas", "Passa?"],
        [["2 - 9 - 14 - 33 - 47", "4 dezenas diferentes", "Sim"],
         ["2 - 5 - 9 - 13 - 17", "2 dezenas diferentes", "Não"]],
        larguras=[52 * mm, 42 * mm, LARGURA - 94 * mm]))

    add(h2("Filtro E: Anti-Repetição"))
    add(p("No máximo dois números em comum com o sorteio imediatamente anterior. "
          "Rejeita muito pouco — apenas 0,5% das combinações — porque é raro haver "
          "três ou mais números repetidos por puro acaso."))

    add(h2("Filtro F: Cores"))
    add(p("É o filtro mais elaborado da aplicação e tem um capítulo próprio, o 7. Em "
          "resumo, classifica os 50 números conforme quantas vezes saíram nos "
          "últimos 9 sorteios e limita quantos de cada tipo podem entrar numa chave."))

    add(h2("Filtro G: Regra do 31"))
    add(p("Exige pelo menos um número acima de 31. <b>É o único filtro com um "
          "benefício medido</b> — e o benefício não é na probabilidade de acertar, é "
          "na partilha. A maioria dos apostadores joga datas de aniversário, que só "
          "usam números de 1 a 31."))
    add(tabela(
        ["Chave", "Números acima de 31", "Passa?", "Co-vencedores"],
        [["4 - 9 - 17 - 23 - 30", "nenhum", "Não", "+4,1% (o pior caso)"],
         ["4 - 9 - 17 - 23 - 44", "um (o 44)", "Sim", "próximo da média"],
         ["4 - 38 - 41 - 46 - 50", "quatro", "Sim", "abaixo da média"]],
        larguras=[46 * mm, 36 * mm, 18 * mm, LARGURA - 100 * mm]))
    add(bom(
        "Se tiver de escolher um único filtro para deixar ligado, é este",
        "Rejeita 8% das combinações — 169.911 das 2.118.760 — e são exactamente as "
        "8% que mais gente joga. Não mexe na probabilidade de acertar: se acertar, "
        "mexe no tamanho da fatia."))

    add(h2("Filtro H: Anti-Progressão"))
    add(p("Recusa progressões aritméticas perfeitas, como 5-10-15-20-25 ou "
          "3-9-15-21-27, que são muito jogadas por fazerem um desenho bonito no "
          "boletim."))
    add(quadro(
        "Este filtro quase não faz nada, e está tudo bem",
        "Existem apenas <b>288 progressões perfeitas</b> em 2.118.760 combinações. O "
        "filtro rejeita 0,014% do universo, e por isso aparece destacado a laranja no "
        "painel da aplicação. Deixá-lo ligado ou desligado não muda praticamente "
        "nada.<br/><br/>"
        "Em 1980 sorteios nunca saiu uma progressão perfeita — o que também não é "
        "notável: ao acaso esperavam-se 0,27."))

    add(h2("Filtro I: AC Value"))
    add(p("AC quer dizer <i>complexidade aritmética</i>. Tomam-se todas as "
          "diferenças entre pares de números da chave, conta-se quantas são "
          "distintas e subtraem-se 4. O resultado vai de 0 a 6. O filtro aceita 4, "
          "5 ou 6, o que cobre 96,4% das combinações."))
    add(tabela(
        ["Chave", "Diferenças distintas", "AC", "Passa?"],
        [["5 - 10 - 15 - 20 - 25", "4 (são 5, 10, 15 e 20)", "0", "Não"],
         ["1 - 2 - 3 - 5 - 8", "8", "4", "Sim"],
         ["3 - 11 - 24 - 29 - 41", "10", "6", "Sim"]],
        larguras=[46 * mm, 52 * mm, 12 * mm, LARGURA - 110 * mm]))
    add(tabela(
        ["AC", "Combinações", "Do universo", "Aceite?"],
        [["0", "288", "0,01%", "Não"],
         ["1", "1.084", "0,05%", "Não"],
         ["2", "26.224", "1,24%", "Não"],
         ["3", "47.874", "2,26%", "Não"],
         ["4", "384.832", "18,16%", "Sim"],
         ["5", "456.710", "21,56%", "Sim"],
         ["6", "1.201.748", "56,72%", "Sim"]],
        larguras=[14 * mm, 32 * mm, 28 * mm, LARGURA - 74 * mm],
        alinhar_dir=(1, 2), destaques=(4, 5, 6)))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("7. O sistema de cores"))
    add(p("É a parte mais original da aplicação. Olha para os <b>últimos 9 "
          "sorteios</b> e pinta cada um dos 50 números de uma cor, conforme quantas "
          "vezes saiu nesse período."))
    add(tabela(
        ["Cor", "Saiu quantas vezes nos últimos 9 sorteios",
         "Quantos pode ter a chave"],
        [["<b>Vermelho</b>", "nenhuma vez", "0 a 3"],
         ["<b>Verde</b>", "uma vez", "0 a 3"],
         ["<b>Azul</b>", "exactamente duas vezes", "0 a 2"],
         ["<b>Castanho</b>", "três ou mais vezes", "0 a 1"]],
        larguras=[26 * mm, 78 * mm, LARGURA - 104 * mm]))
    add(p("Como cada sorteio tira 5 números e a janela tem 9 sorteios, saem 45 "
          "números no total — por isso a maioria dos 50 fica vermelha ou verde e "
          "muito poucos ficam castanhos. Um exemplo real, a 11 de Novembro de 2026:"))
    add(tabela(
        ["Cor", "Quantos", "Números"],
        [["<b>Vermelho</b>", "21",
          "1 2 4 5 6 8 10 11 19 20 25 26 27 28 29 34 36 40 45 46 50"],
         ["<b>Verde</b>", "16", "3 9 12 14 17 18 22 24 30 31 32 33 38 41 42 43"],
         ["<b>Azul</b>", "11", "7 13 16 21 23 35 39 44 47 48 49"],
         ["<b>Castanho</b>", "2", "15 37"]],
        larguras=[26 * mm, 16 * mm, LARGURA - 42 * mm], tamanho=7.6))
    add(p("Nesse dia o filtro aceitava 1.705.870 das 2.118.760 combinações, ou seja "
          "80,5%. Como as cores mudam a cada sorteio, esta percentagem muda também.",
          "legenda"))

    add(h2("As cores servem para prever?"))
    add(p("Não. É o filtro cuja tabela mais chama a atenção, por isso vale a pena ser "
          "explícito: acertou em 78,23% dos sorteios, e ao acaso teria acertado em "
          "79,03%. A diferença é de -0,8 pontos, ou seja, nenhuma."))
    add(p("E tem de ser assim. As cores são calculadas a partir dos 9 sorteios "
          "anteriores, e o sorteio seguinte é independente desses: as bolas não "
          "sabem o que saiu antes. O valor do filtro é outro — distribui as chaves "
          "por números \"frios\" e \"quentes\" de forma equilibrada, o que dá "
          "variedade e evita que todas as chaves geradas se pareçam umas com as "
          "outras."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("8. Quadrantes e os 16 padrões"))
    add(p("A aplicação foi construída sobre uma metodologia chamada Lotterycodex, que "
          "divide os 50 números em quatro grupos, cruzando baixo/alto com par/ímpar."))
    add(tabela(
        ["Sigla", "Grupo", "Quantos", "Números"],
        [["<b>BI</b>", "Baixos Ímpares", "13", "1, 3, 5, ... 25"],
         ["<b>BP</b>", "Baixos Pares", "12", "2, 4, 6, ... 24"],
         ["<b>AI</b>", "Altos Ímpares", "12", "27, 29, 31, ... 49"],
         ["<b>AP</b>", "Altos Pares", "13", "26, 28, 30, ... 50"]],
        larguras=[18 * mm, 40 * mm, 20 * mm, LARGURA - 78 * mm]))
    add(p("Um <b>padrão</b> é a forma de uma chave escrita como quatro números: "
          "quantos BI, quantos BP, quantos AI, quantos AP. Somam sempre 5. Por "
          "exemplo, a chave 3-11-24-29-41 tem o padrão [2,1,2,0]: dois baixos "
          "ímpares (3 e 11), um baixo par (24), dois altos ímpares (29 e 41) e zero "
          "altos pares."))
    add(p("Dos 56 padrões possíveis, a aplicação usa 16 — aqueles a que a "
          "metodologia chama equilibrados. Cobrem 1.333.800 das 2.118.760 "
          "combinações, número que foi verificado por contagem exaustiva e está "
          "correcto."))
    add(tabela(
        ["Padrão", "Chaves possíveis", "Padrão", "Chaves possíveis"],
        [["[1,1,1,2]", "99.032", "[2,0,2,1]", "43.007"],
         ["[2,1,1,1]", "91.858", "[0,2,1,2]", "41.864"],
         ["[1,1,2,1]", "91.414", "[1,2,0,2]", "41.542"],
         ["[1,2,1,1]", "87.062", "[0,1,2,2]", "41.206"],
         ["[2,0,1,2]", "51.484", "[1,2,2,0]", "39.098"],
         ["[2,1,0,2]", "49.171", "[2,1,2,0]", "37.716"],
         ["[1,0,2,2]", "42.759", "[2,2,0,1]", "21.735"],
         ["[0,2,2,1]", "42.454", "[2,2,1,0]", "18.452"]],
        larguras=[28 * mm, 30 * mm, 28 * mm, LARGURA - 86 * mm],
        alinhar_dir=(1, 3)))
    add(p("Chaves possíveis depois de aplicados os nove filtros, no contexto de 11 de "
          "Novembro de 2026. Os dois últimos padrões são os mais restritos: chegaram "
          "a ficar sem chaves nenhumas quando a Regra do 31 exigia dois números "
          "altos, e foi por isso que a regra voltou a exigir apenas um.", "legenda"))

    add(h2("Os padrões servem para prever?"))
    add(p("Também não — e desta vez a medição é particularmente clara. Esta "
          "metodologia é apresentada na aplicação como a base de tudo, por isso foi "
          "testada com o mesmo método usado nos filtros."))
    add(tabela(
        ["O que foi medido", "Real", "Ao acaso", "Desvio"],
        [["Sorteios que caíram num padrão equilibrado", "62,98%", "62,95%",
          "0,03 erros-padrão"],
         ["Grupo \"best\" (os 4 melhores padrões)", "27,07%", "26,42%", "0,66"],
         ["Grupo \"middle-high\"", "35,91%", "36,53%", "0,58"],
         ["Grupo \"middle-low\"", "22,47%", "22,30%", "0,19"],
         ["Grupo \"worst\"", "14,55%", "14,75%", "0,25"]],
        larguras=[74 * mm, 20 * mm, 22 * mm, LARGURA - 116 * mm],
        alinhar_dir=(1, 2, 3)))
    add(aviso(
        "Os 16 padrões são decorativos",
        "Em 1980 sorteios, 62,98% caíram num padrão equilibrado, quando ao acaso "
        "seriam 62,95%. Os grupos \"best\", \"middle\" e \"worst\" <b>não se "
        "ordenam</b>: ficam todos a menos de 0,7 pontos do acaso. E dos 56 padrões "
        "possíveis apenas 1 se desvia mais de 2 erros-padrão, quando ao acaso se "
        "esperavam 2,8 — ou seja, há <i>menos</i> sinal do que o acaso "
        "produziria.<br/><br/>"
        "Testou-se também a única hipótese que os poderia salvar: se reduzem a "
        "partilha. Não reduzem — uma chave equilibrada atrai -0,3% de co-vencedores, "
        "com margem de erro de 1,0 ponto, ou seja, zero.<br/><br/>"
        "Servem para <b>variar a construção das chaves</b>, e nisso são úteis. Não "
        "são uma estratégia."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("9. Estatísticas"))
    add(p("O separador das estatísticas tem vários painéis. Os primeiros descrevem os "
          "sorteios; os últimos avaliam a própria aplicação."))
    add(tabela(
        ["Painel", "O que mostra", "Serve para prever?"],
        [["<b>Heatmap 1-50</b>", "Quantas vezes cada número saiu, em cores", "Não"],
         ["<b>Atrasos</b>", "Há quantos sorteios cada número não sai", "Não"],
         ["<b>Tendência de somas</b>", "A soma de cada sorteio e a média móvel", "Não"],
         ["<b>Linha das dezenas</b>",
          "Como se distribuem as dezenas ao longo do tempo", "Não"],
         ["<b>Condições dos filtros</b>",
          "A regra de cada filtro e quantas vezes acertou", "—"],
         ["<b>Resultados por sorteio</b>",
          "Tabela com Sim/Não de cada filtro, sorteio a sorteio", "—"],
         ["<b>Valor de uma linha</b>",
          "Quanto vale, de facto, apostar 2,50 €", "Capítulo 10"],
         ["<b>Análise de partilha</b>",
          "O resumo do capítulo 3, dentro da aplicação", "—"]],
        larguras=[40 * mm, LARGURA - 72 * mm, 32 * mm]))
    add(quadro(
        "Porque é que a coluna da direita diz sempre Não",
        "Números \"atrasados\" não estão em dívida. Números \"quentes\" não estão "
        "lançados. Cada sorteio é independente de todos os anteriores, e a "
        "probabilidade de o número 17 sair no próximo sorteio é a mesma, tenha ele "
        "saído cinco vezes no último mês ou nenhuma nos últimos dois anos.<br/><br/>"
        "Estes painéis existem porque são interessantes e porque ajudam a escolher "
        "números com alguma intenção, em vez de nenhuma. Não existem porque "
        "funcionam."))

    add(h2("A tabela dos filtros, sorteio a sorteio"))
    add(p("É a tabela maior da aplicação: uma linha por sorteio, com o mais recente "
          "no topo, e uma coluna por filtro. Verde quer dizer que o filtro teria "
          "aceitado a chave sorteada; castanho claro quer dizer que a teria "
          "recusado. Cada sorteio novo entra automaticamente no topo."))
    add(tabela(
        ["Quantos filtros erraram", "Sorteios", "Percentagem"],
        [["Nenhum — todos certos", "1115", "56,57%"],
         ["Pelo menos um errado", "856", "43,43%"],
         ["Só um errado", "673", "34,15%"],
         ["Dois errados", "150", "7,61%"],
         ["Três errados", "30", "1,52%"],
         ["Mais de três errados", "3", "0,15%"]],
        larguras=[62 * mm, 24 * mm, LARGURA - 86 * mm],
        alinhar_dir=(1, 2), destaques=(0,)))
    add(p("Calculado sobre os 1971 sorteios com contexto completo: os nove primeiros "
          "de 2004 não entram, porque o sistema de cores precisa de 9 sorteios "
          "anteriores para poder existir.", "legenda"))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("10. Quanto vale uma aposta de 2,50 €"))
    add(p("A aplicação calcula isto e mostra-o no separador das estatísticas. É "
          "provavelmente a informação mais útil de todo o programa, e é a que nenhum "
          "site de lotarias costuma dar."))
    add(p("A conta é simples: para cada um dos 13 escalões, multiplica-se a "
          "probabilidade pelo prémio, e somam-se os 13 resultados. As probabilidades "
          "são exactas. Os prémios são as medianas efectivamente pagas nos 1039 "
          "sorteios desde Setembro de 2016 — a mediana, e não a média, porque um "
          "punhado de sorteios com bolo muito acumulado puxaria a média para cima e "
          "daria um sorteio típico mais generoso do que ele é."))

    add(h2("Os escalões sem o jackpot"))
    add(p("Somados, os doze escalões abaixo do jackpot valem <b>0,481 €</b> por cada "
          "2,50 € apostados, ou seja <b>19,2%</b>. Este é o chão: o valor que uma "
          "aposta tem mesmo que o jackpot não saia a ninguém."))

    add(h2("Com o jackpot, e com a partilha"))
    add(p("O jackpot depende de duas coisas: do bolo acumulado, que se escreve no "
          "campo da aplicação, e de quantas pessoas o ganham ao mesmo tempo. A "
          "segunda é estimada a partir do próprio sorteio — contam-se os bilhetes "
          "vendidos a partir dos vencedores do escalão de 2 números, cuja "
          "probabilidade é conhecida, e aplica-se depois a matemática da partilha."))
    add(tabela(
        ["Bolo do jackpot", "Bilhetes estimados", "Vale a linha", "Retorno"],
        [["17 milhões (o mínimo)", "17,8 M", "0,60 €", "23,8%"],
         ["50 milhões", "24,0 M", "0,81 €", "32,4%"],
         ["100 milhões", "29,1 M", "1,13 €", "45,1%"],
         ["150 milhões", "32,6 M", "1,44 €", "57,5%"],
         ["200 milhões", "35,3 M", "1,74 €", "69,8%"],
         ["250 milhões (o tecto)", "37,6 M", "2,05 €", "81,9%"]],
        larguras=[48 * mm, 34 * mm, 28 * mm, LARGURA - 110 * mm],
        alinhar_dir=(1, 2, 3)))
    add(aviso(
        "Não existe bolo que torne a aposta favorável",
        "Mesmo com o jackpot no tecto de 250 milhões, uma linha de 2,50 € vale cerca "
        "de <b>2,05 €</b>. O retorno nunca chega aos 100%: perde-se sempre, em "
        "média, entre 45 cêntimos e 1,90 € por cada linha apostada.<br/><br/>"
        "É assim que as lotarias funcionam, e não há nada de errado nisso — parte do "
        "dinheiro vai para prémios e parte para o Estado e para causas sociais. Mas "
        "vale a pena saber o número antes de apostar, e não depois."))
    add(quadro(
        "Uma coisa que falta nesta conta: o M1LHÃO",
        "Cada aposta portuguesa de 2,50 € dá também um código para o sorteio do "
        "M1LHÃO, que acrescenta valor ao que está calculado acima. Não entra na "
        "conta porque a aplicação não tem o número de códigos portugueses em "
        "circulação em cada sorteio, e sem isso qualquer valor seria "
        "inventado.<br/><br/>"
        "O valor real de uma aposta feita em Portugal é portanto <b>maior</b> do que "
        "o das tabelas acima — mas continua a ser inferior aos 2,50 € que custa."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("11. Conferir as suas apostas"))
    add(p("O separador <b>Último Sorteio</b> mostra o resultado mais recente, os "
          "prémios por escalão e o padrão BI-BP-AI-AP da chave sorteada."))
    add(tabela(
        ["Funcionalidade", "O que faz"],
        [["<b>Comparar com sorteio</b>",
          "Pega nas chaves geradas e diz quantos números e estrelas cada uma acertou"],
         ["<b>Conferir apostas</b>",
          "Verifica uma lista de apostas suas contra o resultado, escalão a escalão"],
         ["<b>Importar apostas de Excel</b>",
          "Lê um ficheiro com as suas apostas e confere-as todas de uma vez"],
         ["<b>Recolher prémios</b>",
          "Vai buscar o detalhe de prémios de sorteios antigos que ainda não o tenham"]],
        larguras=[52 * mm, LARGURA - 52 * mm]))
    add(p("A conferência usa a tabela oficial dos 13 escalões, por isso identifica "
          "também os prémios pequenos que passam facilmente ao lado: acertar 2 "
          "números, ou 1 número e 2 estrelas, dá prémio — e é fácil não dar por isso."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("12. O histórico e a base de dados"))
    add(p("A aplicação tem a <b>totalidade dos sorteios do Euromilhões</b>: 1980 "
          "sorteios, desde o primeiro, a 13 de Fevereiro de 2004, até ao mais "
          "recente. Para 1978 deles tem também o detalhe de prémios — quantas "
          "pessoas ganharam em cada escalão e quanto receberam. Faltam dois "
          "sorteios cujos prémios a fonte não publica: 1 de Outubro de 2004 e 28 de "
          "Abril de 2025."))

    add(h2("Actualização automática"))
    add(p("Não é preciso fazer nada para a base de dados se manter actual. Um "
          "processo automático corre à terça e à sexta, pouco depois do sorteio, vai "
          "buscar o resultado novo e os prémios, e grava-os. Se os prémios ainda não "
          "estiverem publicados — costumam sair algumas horas depois do sorteio — a "
          "execução seguinte apanha-os."))

    add(h2("Actualização manual"))
    add(tabela(
        ["Opção", "Quando usar"],
        [["<b>Actualizar recentes</b>",
          "Para forçar a busca do sorteio mais recente"],
         ["<b>Inserir manualmente</b>",
          "Para acrescentar um sorteio escrevendo os números"],
         ["<b>Importar ficheiro</b>",
          "Para carregar um histórico completo de uma só vez"],
         ["<b>Importar Excel</b>", "O mesmo, a partir de uma folha de cálculo"],
         ["<b>Recolher prémios</b>",
          "Para preencher prémios em falta de sorteios antigos"],
         ["<b>Eliminar sorteio</b>",
          "Para corrigir um registo errado (a operação pode ser desfeita)"]],
        larguras=[50 * mm, LARGURA - 50 * mm]))
    add(p("O histórico apresenta-se agrupado por ano e por mês, em secções que se "
          "abrem e fecham, com o ano corrente já aberto."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("13. Ligação ao AIJackpot"))
    add(p("Quem joga em grupo tem um problema a mais: alguém tem de registar quem "
          "apostou o quê, quanto cada um pôs, e como se divide o que se ganhar. Para "
          "isso existe uma segunda aplicação, o <b>AIJackpot</b>, que gere um grupo "
          "de 32 apostadores."))
    add(p("O botão <b>JSC</b>, no ecrã de geração, envia as chaves geradas "
          "directamente para o AIJackpot. Do outro lado, a aplicação calcula "
          "automaticamente o valor apostado e, quando o resultado oficial chega, os "
          "acertos e os prémios de cada chave."))
    add(tabela(
        ["Passo", "O que acontece"],
        [["1", "Gera as chaves no Gerador Euromilhões"],
         ["2", "Carrega no botão <b>JSC</b>"],
         ["3",
          "As chaves são numeradas automaticamente com o ano e o mês do próximo sorteio"],
         ["4", "São enviadas para o AIJackpot e aparecem na conta do grupo"],
         ["5", "Pode também enviar a lista para o WhatsApp, para o grupo ver"]],
        larguras=[15 * mm, LARGURA - 15 * mm]))
    add(p("A numeração reinicia sozinha quando muda o mês. O botão <b>Configurar "
          "JSC</b> permite ajustar o contador, se por alguma razão for preciso "
          "saltar números.", "legenda"))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("14. Jogo responsável"))
    add(p("Este guia descreve uma ferramenta que ajuda a escolher números. Não é um "
          "incentivo a apostar, e seria desonesto terminá-lo sem dizer o que se "
          "segue — que decorre directamente das contas do capítulo 10."))
    add(lista([
        "<b>Apostar dá prejuízo, em média, sempre.</b> Entre 18% e 76% do valor "
        "apostado não volta. Não há estratégia, filtro ou padrão que mude isto, e "
        "esta aplicação não muda isto.",
        "<b>Nenhum filtro aumenta a probabilidade de acertar.</b> Está medido em "
        "1971 sorteios, filtro a filtro, e os resultados estão no capítulo 6.",
        "<b>O único efeito real é na partilha</b>, e vale poucos por cento.",
        "<b>Aposte apenas o que pode perder sem consequências</b>, e trate o valor "
        "como o preço de um entretenimento, não como um investimento.",
        "Em Portugal, apoio para problemas com o jogo: <b>SICAD, Linha Vida, "
        "1414</b> (chamada gratuita).",
    ]))
    add(quadro(
        "O que esta aplicação é, honestamente",
        "Uma forma organizada e informada de escolher números, de guardar as "
        "escolhas e de conferir os resultados, com a vantagem pequena mas real de "
        "evitar as combinações que muita gente joga. Se acertar o jackpot, há uma "
        "probabilidade ligeiramente maior de não o dividir com ninguém.<br/><br/>"
        "É isto — e é mais do que a maioria das aplicações do género pode dizer com "
        "números atrás."))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("15. Glossário"))
    add(tabela(
        ["Termo", "Significado"],
        [["<b>AC Value</b>",
          "Complexidade aritmética: quantas diferenças distintas há entre os pares da chave, menos 4. Vai de 0 a 6."],
         ["<b>BI, BP, AI, AP</b>",
          "Os quatro quadrantes: Baixos Ímpares, Baixos Pares, Altos Ímpares, Altos Pares."],
         ["<b>Bolo, ou acumulado</b>",
          "O valor do jackpot num determinado sorteio, que cresce quando ninguém acerta."],
         ["<b>Chave</b>",
          "Uma combinação de 5 números (e, na aposta completa, 2 estrelas)."],
         ["<b>Erro-padrão</b>",
          "A margem de incerteza de uma medição. Um resultado a menos de 2 erros-padrão de zero não se distingue do acaso."],
         ["<b>Escalão</b>",
          "Um dos 13 níveis de prémio, definido por quantos números e estrelas se acertou."],
         ["<b>Filtro</b>", "Uma regra que recusa certas combinações na geração."],
         ["<b>Intervalo de confiança</b>",
          "O intervalo onde o valor verdadeiro provavelmente está. Se inclui zero, não se pode afirmar que haja efeito."],
         ["<b>Jackpot</b>", "O primeiro escalão: 5 números e 2 estrelas."],
         ["<b>Mediana</b>",
          "O valor do meio de uma lista ordenada. Mais robusta que a média quando há valores extremos."],
         ["<b>Padrão</b>",
          "A forma de uma chave em termos de quadrantes, como [2,1,2,0]."],
         ["<b>Partilha</b>",
          "Dividir um prémio com outras pessoas que acertaram o mesmo."],
         ["<b>Valor esperado</b>",
          "Quanto, em média, uma aposta devolve. No Euromilhões é sempre inferior ao que custa."]],
        larguras=[36 * mm, LARGURA - 36 * mm], tamanho=7.9))

    # ═══════════════════════════════════════════════════════════════════════
    add(PageBreak())
    add(h1("16. Fontes e referências"))
    add(h2("De onde vêm os dados da aplicação"))
    add(tabela(
        ["Fonte", "O que fornece", "Endereço"],
        [["Jogos Santa Casa",
          "Operador oficial em Portugal; resultado do último sorteio",
          "jogossantacasa.pt/web/SCCartazResult/euroMilhoes"],
         ["euro-millions.com",
          "Histórico completo dos sorteios e detalhe de prémios por escalão",
          "euro-millions.com/results"],
         ["lotaria.net", "Resultados e arquivo, usado como fonte alternativa",
          "lotaria.net/euromilhoes/arquivo"],
         ["euromillones.com", "Fonte adicional de resultados e prémios",
          "euromillones.com/pt/resultados/euromilhoes"]],
        larguras=[32 * mm, 56 * mm, LARGURA - 88 * mm], tamanho=7.7))

    add(h2("Metodologia"))
    add(tabela(
        ["Referência", "Onde entra nesta aplicação"],
        [["<b>Lotterycodex</b> — metodologia dos quadrantes BI-BP-AI-AP e dos padrões combinatórios (lotterycodex.com)",
          "Base da construção das chaves: os 4 quadrantes e os 16 padrões equilibrados do capítulo 8. Foi testada nos dados da própria aplicação, com o resultado que ali se descreve."],
         ["<b>Distribuição hipergeométrica</b> — cálculo exacto das probabilidades de cada escalão",
          "As probabilidades do capítulo 2 e o cálculo do valor esperado do capítulo 10."],
         ["<b>Distribuição de Poisson</b> — modelo do número de vencedores simultâneos",
          "A matemática da partilha do jackpot, no capítulo 10."],
         ["<b>Fiabilidade por divisão em metades</b> — teste clássico para saber se uma estimativa mede algo real ou apenas ruído",
          "O aviso do capítulo 3 sobre a impossibilidade de identificar números mais jogados."]],
        larguras=[76 * mm, LARGURA - 76 * mm], tamanho=7.7))

    add(h2("Como foram obtidos os números deste guia"))
    add(lista([
        "<b>Percentagens de rejeição dos filtros:</b> contagem exaustiva das "
        "2.118.760 combinações, uma a uma. Não são estimativas.",
        "<b>Taxas de acerto históricas:</b> avaliação de cada filtro contra a chave "
        "sorteada, em todos os 1971 sorteios com contexto completo.",
        "<b>Efeitos na partilha:</b> cruzamento do número de vencedores de cada "
        "escalão com as características de cada sorteio, em 1978 sorteios, "
        "controlando o valor do jackpot e as mudanças de regras do jogo.",
        "<b>Prémios típicos e valor esperado:</b> medianas dos 1039 sorteios desde "
        "24 de Setembro de 2016, quando as estrelas passaram de 11 para 12.",
    ], "corpo_pequeno"))
    add(Spacer(1, 4))
    add(p("Os dados estão guardados na própria aplicação e podem ser reexaminados a "
          "qualquer momento. Os números deste guia referem-se a 11 de Novembro de "
          "2026, com 1980 sorteios na base de dados; os que dependem do estado "
          "actual, como as cores, mudam a cada sorteio.", "legenda"))

    add(Spacer(1, 8))
    add(tabela(
        ["", ""],
        [["<b>Aplicação</b>", "euromilhoes-chi.vercel.app"],
         ["<b>Autor</b>", "%s, %s" % (AUTOR, ANO)],
         ["<b>Colaboração</b>", COLABORACAO],
         ["<b>Guia actualizado em</b>", DATA_GUIA]],
        larguras=[48 * mm, LARGURA - 48 * mm], cor_cab=colors.white))
    return s


# ═══════════════════════════════════════════════════════════════════════════
def verificar_texto(story):
    """As fontes Helvetica usam WinAnsi. Um caracter fora disso sai como
    quadrado no PDF, em silêncio — melhor falhar aqui."""
    maus = {}

    def olhar(obj):
        if isinstance(obj, Table):
            for linha in obj._cellvalues:
                for cel in linha:
                    olhar(cel)
            return
        if not isinstance(obj, Paragraph):
            return
        for ch in obj.getPlainText():
            try:
                ch.encode("cp1252")
            except UnicodeEncodeError:
                maus[ch] = maus.get(ch, 0) + 1

    for f in story:
        if isinstance(f, KeepTogether):
            for g in f._content:
                olhar(g)
        else:
            olhar(f)
    return maus


def main():
    story = construir()
    maus = verificar_texto(story)
    if maus:
        print("CARACTERES QUE NAO CABEM NA FONTE (sairiam como quadrados):")
        for ch, n in sorted(maus.items(), key=lambda kv: -kv[1]):
            print("   U+%04X x%d" % (ord(ch), n))
        return 1
    doc = Guia(str(SAIDA))
    doc.multiBuild(story)
    print("Gravado: %s" % SAIDA)
    print("Tamanho: %.0f KB" % (SAIDA.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    sys.exit(main())
