from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
import tempfile
from docx.shared import RGBColor
import re
from Condicoes import cidades_bahia,palavras_orientador
import unicodedata


app = Flask(__name__)
CORS(app)

def _normalizar_para_busca(texto):
    """Remove pontuação chata, sufixos como '- BA' ou '/BA', múltiplos espaços, e deixa minusculo."""
    if not texto:
        return ""
    t = texto.strip().lower()

    # remove sufixos comuns: "- ba", "/ba", ", ba", " - bahia", "/bahia", etc.
    t = re.sub(r'[-/]\s*ba(hia)?\b', '', t)            # "- ba", "/bahia"
    t = re.sub(r',\s*ba(hia)?\b', '', t)               # ", BA"
    t = re.sub(r'\s+ba(hia)?\b', '', t)                # " BA" ou " BAHIA"
    t = re.sub(r'\bestado da bahia\b', '', t)

    # remove pontuação (mantém letras e espaços)
    t = re.sub(r'[^0-9a-zà-ú\s]', ' ', t, flags=re.IGNORECASE)

    # colapsa espaços e trim
    t = re.sub(r'\s+', ' ', t).strip()
    return t

def detectar_tipo(linha, identificados):
    texto = linha.strip().lower()

    # Ordem de prioridade para classificação
    if texto.isdigit() and len(texto) == 4:
        return "ano"
    if "universidade" in texto or "instituto" in texto or "faculdade" in texto:
        return "instituicao"
    if len(texto.split()) >= 3 and texto[0].isupper() and identificados["autor"] is None:
        return "autor"
    if len(texto) > 40 and identificados["titulo"] is None:
        return "titulo"
    if identificados["titulo"] is not None and identificados["subtitulo"] is None and len(texto) > 15:
        return "subtitulo"
    if texto in ["feira de santana", "salvador", "são paulo", "rio de janeiro"] and identificados["cidade"] is None:
        return "cidade"

    return None


def identificar_linhas_da_capa(doc):
    linhas = [p.text for p in doc.paragraphs]

    identificados = {
        "instituicao": None,
        "autor": None,
        "curso": None,
        "titulo": None,
        "subtitulo": None,
        "cidade": None,
        "ano": None
    }

    classificados = set()  # <- impede alterar classificação depois

    for _ in range(7):  # repete para refinar a identificação
        for i, linha in enumerate(linhas):
            if not linha.strip():
                continue

            if i in classificados:
                continue  # <- já foi classificado, não muda mais

            tipo = detectar_tipo(linha, identificados)

            if tipo and identificados[tipo] is None:
                identificados[tipo] = i
                classificados.add(i)  # <- trava a linha como aquele tipo

    return identificados


def detectar_fonte_principal(doc):
    contagem_fontes = {}

    # Verifica a fonte configurada no estilo Normal
    try:
        estilo_normal = doc.styles["Normal"].font.name
    except:
        estilo_normal = None

    for p in doc.paragraphs:
        for run in p.runs:
            fonte = run.font.name or estilo_normal or "Arial"
            fonte = fonte.strip()

            if fonte not in contagem_fontes:
                contagem_fontes[fonte] = 0

            contagem_fontes[fonte] += len(run.text)

    # Descobre a fonte mais usada
    fonte_predominante = max(contagem_fontes, key=contagem_fontes.get)

    # Normaliza para Arial ou Times
    if "arial" in fonte_predominante.lower():
        return "Arial"

    if "times" in fonte_predominante.lower():
        return "Times New Roman"

    # Se for outra fonte → força Arial
    return "Arial"

def eh_instituicao(texto):
    """
    Retorna True apenas se o texto contiver palavras típicas de instituições.
    Nenhum outro critério é considerado.
    """
    if not texto:
        return False

    texto_limpo = texto.strip().lower()

    PALAVRAS_INSTITUICOES = [
        "universidade",
        "centro universit",
        "instituto",
        "faculdade",
        "escola",
        "colégio",
        "campus",
        "departamento",
        "programa de pós"
    ]

    return any(p in texto_limpo for p in PALAVRAS_INSTITUICOES)

def classificar_capa(texto):
    texto_limpo = texto.strip()

    if not texto_limpo:
        return "vazio"

    # Instituição: normalmente grande, não começa com número e tem várias palavras
    if len(texto_limpo.split()) >= 3 and texto_limpo.isupper():
        return "instituicao"

    # Autor: costuma ter nome + sobrenome iniciando maiúsculas
    if re.match(r'^[A-ZÁÀÃÂÉÍÓÚ][a-z].+ [A-ZÁÀÃÂÉÍÓÚ][a-z].+', texto_limpo):
        return "autor"

    # Curso
    if any(palavra in texto_limpo.lower() for palavra in ["curso", "bacharel", "licenciatura", "tecnólogo"]):
        return "curso"

    # Título da capa: mais longo, não precisa ser uppercase
    if len(texto_limpo.split()) >= 4 and not re.match(r'^\d', texto_limpo):
        return "titulo_capa"

    # Subtitulo (se houver)
    if texto_limpo.lower().startswith(("um estudo", "análise", "desenvolvimento", "projeto", "monografia", "trabalho")):
        return "subtitulo"

    # Cidade
    if texto_limpo.lower() in ["feira de santana", "salvador", "são paulo", "rio de janeiro"]:
        return "cidade"

    # Ano (4 dígitos)
    if re.match(r'^\d{4}$', texto_limpo):
        return "ano"

    return "outro"

def eh_ano(texto):
    texto = (texto or "").strip()
    # procura um ano isolado como 1999, 2025, 2010 (entre limites razoáveis)
    m = re.search(r"\b(19|20)\d{2}\b", texto)
    if m:
        # opcional: retorna o ano encontrado se precisar
        return True
    return False



from Condicoes import cidades_bahia


def remover_acentos(txt):
    """Remove acentos e normaliza texto para comparação segura."""
    return ''.join(
        c for c in unicodedata.normalize('NFD', txt)
        if unicodedata.category(c) != 'Mn'
    )

def eh_cidade(texto, cidades_bahia):
    """Retorna True se o texto corresponder a uma cidade baiana."""
    
    if not texto:
        return False

    texto_norm = remover_acentos(texto.strip().lower())

    for cidade in cidades_bahia:
        cidade_norm = remover_acentos(cidade.strip().lower())

        # Igualdade direta ou texto contém a cidade (ex: "Feira de Santana - BA")
        if cidade_norm == texto_norm or cidade_norm in texto_norm:
            return True

    return False


print(eh_cidade("Feira de Santana", cidades_bahia))


def eh_autor(t):
    # Nome de pessoa: 2 a 5 palavras com iniciais maiúsculas
    return re.fullmatch(r"([A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+)(\s+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+){1,4}", t) is not None

def classificar_linhas(linhas):
    identificados = {
        "ano": None,
        "instituicao": None,
        "titulo": None,
        "subtitulo": None,
        "autor": None,
        "cidade": None
    }

    resto = linhas[:]

    for _ in range(6):  # repete até estabilizar
        novos_resto = []
        for t in resto:
            s = t.strip()

            if identificados["ano"] is None and eh_ano(s):
                identificados["ano"] = s
                continue

            if identificados["instituicao"] is None and eh_instituicao(s):
                identificados["instituicao"] = s
                continue

            if identificados["titulo"] is None and eh_titulo(s):
                identificados["titulo"] = s
                continue

            if identificados["titulo"] and identificados["subtitulo"] is None and eh_subtitulo(s):
                identificados["subtitulo"] = s
                continue

            if identificados["autor"] is None and eh_autor(s):
                identificados["autor"] = s
                continue

            if identificados["cidade"] is None and eh_cidade(s,cidades_bahia):
                identificados["cidade"] = s
                continue

            novos_resto.append(t)

        if len(novos_resto) == len(resto):
            break

        resto = novos_resto

    return identificados

from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def formatar_capa(doc):
    """
    Formata a capa do TCC com base nos tipos de texto identificados.
    Inclui filtro de cidades da Bahia (em maiúsculas nas iniciais).
    """
    def formatar_cidade(texto):
        # transforma "feira de santana" → "Feira De Santana"
        return " ".join([palavra.capitalize() for palavra in texto.lower().split()])

    # encontra índice do primeiro parágrafo que começa com "resumo"
    index_resumo = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text and p.text.strip().lower().startswith("resumo"):
            index_resumo = idx
            break

    # se não achar resumo, define como comprimento total (processa tudo)
    limite = index_resumo if index_resumo is not None else len(doc.paragraphs)

    encontrou_instituicao = False
    encontrou_autor = False
    encontrou_curso = False
    encontrou_titulo = False
    encontrou_subtitulo = False
    encontrou_cidade = False
    encontrou_ano = False
    titulo_identificado = False

    for i, p in enumerate(doc.paragraphs):
        if i >= limite:
            break

        texto = p.text.strip()
        if not texto:
            continue

        tipo, titulo_identificado = classificar_texto(texto, titulo_identificado)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER  # Centraliza tudo na capa

        # --- Instituição ---
        if tipo == "instituicao" and not encontrou_instituicao:
            encontrou_instituicao = True
            for run in p.runs:
                run.font.size = Pt(12)
            continue

        # --- Autor ---
        if tipo == "autor" and not encontrou_autor:
            encontrou_autor = True
            for run in p.runs:
                run.font.size = Pt(12)
            continue

        # --- Curso ---
        if tipo == "curso" and not encontrou_curso:
            encontrou_curso = True
            for run in p.runs:
                run.font.size = Pt(12)
            continue

        # --- Título ---
        if tipo == "titulo_capa" and not encontrou_titulo:
            encontrou_titulo = True
            p.text = p.text.upper()
            for run in p.runs:
                run.font.size = Pt(14)
                run.bold = True
            continue

        # --- Subtítulo ---
        if tipo == "subtitulo" and encontrou_titulo and not encontrou_subtitulo:
            encontrou_subtitulo = True
            for run in p.runs:
                run.font.size = Pt(12)
            continue

        # --- Ano ---
        if eh_ano(texto) and not encontrou_ano:
            encontrou_ano = True
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(128, 128, 128)  # cinza
            continue

# -----------------------------------------------------------
# 🔹 Estado global para garantir que o título só apareça 1x
# -----------------------------------------------------------



from docx import Document

def eh_titulo(texto, titulo_identificado):
    """
    Detecta se o texto é o título principal da capa.
    Retorna uma tupla: (eh_titulo, novo_estado)
    """
    texto_limpo = texto.strip()
    if not texto_limpo or titulo_identificado:
        return False, titulo_identificado  # já identificou antes → não é título

    # Critérios mais fortes:
    # 1. Todo em maiúsculo (ou quase)
    # 2. Comprimento mínimo de 4 palavras
    maiusculas = sum(1 for c in texto_limpo if c.isupper())
    minusculas = sum(1 for c in texto_limpo if c.islower())
    proporcao_maiusculas = maiusculas / max(1, (maiusculas + minusculas))

    if (proporcao_maiusculas >= 0.85 or texto_limpo.isupper()) and len(texto_limpo.split()) >= 4:
        return True, True  # marcou o título e bloqueia o próximo
    return False, titulo_identificado


def eh_subtitulo(texto, titulo_identificado):
    """
    Detecta se o texto é um subtítulo da capa.
    Só é chamado após o título principal ter sido identificado.
    """
    texto_limpo = texto.strip()
    if not texto_limpo:
        return False

    # Só pode existir subtítulo depois que o título já foi encontrado
    if not titulo_identificado:
        return False

    # Subtítulo geralmente não é todo em maiúsculas
    if texto_limpo.isupper():
        return False

    # Pode adicionar outros critérios se quiser (ex: tamanho do texto, etc)
    return True


# 🔹 Exemplo de uso (funciona em qualquer parte do app)
def identificar_titulos(doc):
    titulo_identificado = False

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue

        eh_tit, titulo_identificado = eh_titulo(texto, titulo_identificado)
        if eh_tit:
            print("Título principal encontrado:", texto)
            continue

        if eh_subtitulo(texto, titulo_identificado):
            print("Subtítulo encontrado:", texto)

# -------- CLASSIFICAR PARÁGRAFOS --------
def classificar_texto(texto, titulo_identificado):
    texto = texto.strip()
    if not texto:
        return "vazio", titulo_identificado

    # --- CAPA ---
    if eh_instituicao(texto):
        return "instituicao", titulo_identificado

    if eh_autor(texto):
        return "autor", titulo_identificado
    
    if any(p in texto.lower() for p in palavras_orientador):
        return "orientador", titulo_identificado


    PALAVRAS_CURSO = [
        "curso", "bacharelado", "licenciatura", "engenharia",
        "tecnologia", "ciência", "sistemas", "informação",
        "direito", "medicina", "administração"
    ]
    if any(p in texto.lower() for p in PALAVRAS_CURSO):
        return "curso", titulo_identificado

    eh_tit, novo_estado = eh_titulo(texto, titulo_identificado)
    if eh_tit:
        return "titulo_capa", novo_estado

    if eh_cidade(texto, cidades_bahia):
        return "cidade", titulo_identificado

    if eh_ano(texto):
        return "ano", titulo_identificado

    # --- CORPO DO TEXTO ---
    if re.match(r"^\d+\s+[A-ZÁÉÍÓÚÂÊÔÃÕ]", texto):
        return "titulo_principal", titulo_identificado

    if re.match(r"^\d+\.\d+(\.\d+)*\s+[A-ZÁÉÍÓÚÂÊÔÃÕ]", texto):
        return "subtitulo", titulo_identificado

    if eh_subtitulo(texto, titulo_identificado):
        return "subtitulo", titulo_identificado
    # --- FALLBACK: cidade ---

    # Caso não se encaixe em nada:
    return "paragrafo", titulo_identificado


# -------- APLICAR FORMATAÇÃO --------
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import re

def aplicar_formatacao(doc, fonte_principal):
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    import re

    # ---------- 1. Define a fonte padrão ----------
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = fonte_principal
            try:
                r = run._element.rPr.rFonts
                r.set(qn("w:ascii"), fonte_principal)
                r.set(qn("w:hAnsi"), fonte_principal)
            except:
                pass

    antes_do_resumo = True
    titulo_identificado = False

    # ---------- 2. Flags de controle (só permite 1 de cada) ----------
    encontrado = {
        "instituicao": False,
        "autor": False,
        "orientador": False,
        "curso": False,
        "titulo_capa": False,
        "subtitulo": False,
        "cidade": False,
        "ano": False,
    }

    # ---------- 3. Loop de formatação ----------
    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue

        # Interrompe antes do resumo
        if re.match(r"^\s*resumo\b", texto, re.IGNORECASE):
            antes_do_resumo = False
            continue

        if not antes_do_resumo:
            continue

        tipo, titulo_identificado = classificar_texto(texto, titulo_identificado)
        print(f"[DEBUG] Linha analisada: '{texto}'")
        print(f"[DEBUG] Tipo detectado: {tipo}")

        # ---------- INSTITUIÇÃO ----------
        if tipo == "instituicao" and not encontrado["instituicao"]:
            encontrado["instituicao"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0, 128, 0)
            print("[DEBUG] Instituição formatada (verde)")
            continue

        # ---------- AUTOR ----------
        if tipo == "autor" and not encontrado["autor"]:
            encontrado["autor"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0, 0, 255)
            print("[DEBUG] Autor formatado (azul)")
            continue

        # ---------- ORIENTADOR -----------

        # ---------- ORIENTADOR ----------
        if tipo == "orientador" and not encontrado.get("orientador", False):
            encontrado["orientador"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 0, 255)  # rosa / magenta
                run.text = run.text.replace(":", "").strip().capitalize()
            print("[DEBUG] Orientador formatado (magenta)")
            continue


        # ---------- CURSO ----------
        if tipo == "curso" and not encontrado["curso"]:
            encontrado["curso"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(255, 128, 0)
            print("[DEBUG] Curso formatado (laranja)")
            continue

        # ---------- TÍTULO ----------
        if tipo == "titulo_capa" and not encontrado["titulo_capa"]:
            encontrado["titulo_capa"] = True
            p.text = p.text.upper()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(14)
                run.bold = True
                run.font.color.rgb = RGBColor(0, 153, 0)
            print("[DEBUG] Título formatado (verde escuro)")
            continue

        # ---------- SUBTÍTULO ----------
        if tipo == "subtitulo" and not encontrado["subtitulo"]:
            encontrado["subtitulo"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.bold = False
                run.font.color.rgb = RGBColor(102, 0, 204)
            print("[DEBUG] Subtítulo formatado (roxo)")
            continue

        # ---------- CIDADE ----------
        if tipo == "cidade" and not encontrado["cidade"]:
            encontrado["cidade"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(0, 255, 255)
                run.text = " ".join([palavra.capitalize() for palavra in run.text.lower().split()])
            print("[DEBUG] Cidade formatada (ciano)")
            continue

        # ---------- ANO ----------
        if tipo == "ano" and not encontrado["ano"]:
            encontrado["ano"] = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.color.rgb = RGBColor(128, 128, 128)
            print("[DEBUG] Ano formatado (cinza)")
            continue

  
# -------- VERIFICAR FORMATAÇÃO --------
def verificar_formatacao(doc):
    erros = []
    dentro_corpo = False
    titulo_capa_verificado = False
    titulo_identificado = False
    subtitulo_capa_verificado = False

    for p in doc.paragraphs:
        texto = p.text.strip()
        if not texto:
            continue

        tipo, titulo_identificado = classificar_texto(texto, titulo_identificado)


        

        # --- TÍTULO DA CAPA ---
        if tipo == "titulo_capa" and not titulo_capa_verificado:
            titulo_capa_verificado = True

            tamanhos = {run.font.size.pt for run in p.runs if run.font.size}
            if tamanhos and any(t != 14 for t in tamanhos):
                erros.append(f"❌ O título da capa deve estar no tamanho 14 → \"{texto}\"")

            if not any(run.bold for run in p.runs):
                erros.append(f"⚠️ O título da capa deve estar em NEGRITO → \"{texto}\"")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                erros.append(f"⚠️ O título da capa deve estar CENTRALIZADO → \"{texto}\"")

            if texto != texto.upper():
                erros.append(f"⚠️ O título da capa deve estar TODO EM MAIÚSCULO → \"{texto}\"")

            continue

        # --- SUBTÍTULO DA CAPA ---
        if tipo == "subtitulo" and not subtitulo_capa_verificado and not dentro_corpo:
            subtitulo_capa_verificado = True

            tamanhos = {run.font.size.pt for run in p.runs if run.font.size}
            if tamanhos and any(t != 12 for t in tamanhos):
                erros.append(f"❌ O subtítulo da capa deve estar no tamanho 12 → \"{texto}\"")

            if any(run.bold for run in p.runs):
                erros.append(f"⚠️ O subtítulo da capa não deve estar em negrito → \"{texto}\"")

            if p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                erros.append(f"⚠️ O subtítulo da capa deve estar CENTRALIZADO → \"{texto}\"")

            if texto == texto.upper():
                erros.append(f"⚠️ O subtítulo não deve estar TODO EM MAIÚSCULO → \"{texto}\"")

            continue

        # --- DETECTA INÍCIO DO CORPO ---
        if re.match(r'^\d+(\.\d+)*\s', texto):
            dentro_corpo = True

        # --- IGNORA O RESTO DA CAPA ---
        if not dentro_corpo:
            continue

        # --- TÍTULOS DO CORPO ---
        if tipo in ("titulo_principal", "subtitulo"):
            tamanhos = {run.font.size.pt for run in p.runs if run.font.size}
            if tamanhos and any(t != 12 for t in tamanhos):
                erros.append(f"❌ Tamanho incorreto no título → \"{texto}\" (deve ser 12)")

            if p.alignment != WD_ALIGN_PARAGRAPH.LEFT:
                erros.append(f"⚠️ O título deve estar alinhado à esquerda → \"{texto}\"")
            continue

        # --- PARÁGRAFOS NORMAIS ---
        tamanhos = {run.font.size.pt for run in p.runs if run.font.size}
        if tamanhos and any(t != 12 for t in tamanhos):
            erros.append(f"❌ Tamanho incorreto no parágrafo → \"{texto}\" (deve ser 12)")

        if p.alignment not in (WD_ALIGN_PARAGRAPH.JUSTIFY, None):
            erros.append(f"⚠️ O parágrafo deve ser justificado → \"{texto}\"")

    return erros if erros else ["✅ Nenhum problema de formatação encontrado."]


def verificar_margens(doc):
    padrao = [3, 2, 3, 2]
    s = doc.sections[0]
    margens = [round(s.top_margin.cm), round(s.bottom_margin.cm),
               round(s.left_margin.cm), round(s.right_margin.cm)]
    return margens == padrao


@app.route("/formatar", methods=["POST"])
def formatar():
    if "arquivo" not in request.files:
        return {"erro": "Envie um arquivo .docx"}, 400

    arquivo = request.files["arquivo"]
    caminho = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    arquivo.save(caminho)

    doc = Document(caminho)
    fonte = detectar_fonte_principal(doc)

    # -------- AJUSTA ESTILO NORMAL --------
    try:
        estilo_normal = doc.styles["Normal"].font
        estilo_normal.name = fonte
        for r in doc.styles["Normal"].element.xpath(".//w:rFonts"):
            r.set(qn("w:ascii"), fonte)
            r.set(qn("w:hAnsi"), fonte)
    except:
        pass

    formatar_capa(doc)
    aplicar_formatacao(doc, fonte)

    saida = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    doc.save(saida)
    return send_file(saida, as_attachment=True, download_name="arquivo_formatado_ABNT.docx")


@app.route("/verificar", methods=["POST"])
def verificar():
    if "arquivo" not in request.files:
        return {"erro": "Envie um arquivo .docx"}, 400

    arquivo = request.files["arquivo"]
    caminho = tempfile.NamedTemporaryFile(delete=False, suffix=".docx").name
    arquivo.save(caminho)
    doc = Document(caminho)

    return {
        "margens_corretas": verificar_margens(doc),
        "formatacao": verificar_formatacao(doc)
    }


if __name__ == "__main__":
    app.run(debug=True)
