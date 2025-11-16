from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml.ns import qn
import tempfile
from docx.shared import RGBColor
import re
from Condicoes import cidades_bahia, palavras_orientador, cursos, instituicao
import unicodedata


app = Flask(__name__)
CORS(app)


ja_tem_autor = False
ja_tem_cidade = False
ja_tem_instituicao = False
ja_tem_titulo_capa = False
ja_tem_curso = False
ja_tem_ano = False


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
    texto_original = linha.strip()
    texto_lower = texto_original.lower()

    # ============================================
    # 1) ANO (regra mais simples e 100% segura)
    # ============================================
    if texto_original.isdigit() and len(texto_original) == 4:
        return "ano"

    # ============================================
    # 2) CIDADE (com lista prévia e UF opcional)
    # ============================================
    if identificados["cidade"] is None:
        cidade = eh_cidade(texto_original, cidades_bahia)
        if cidade:
            return "cidade"

    # 3) INSTITUIÇÃO (função dedicada)
    if identificados["instituicao"] is None:
        if eh_instituicao(texto_original):
            return "instituicao"


    # ============================================
    # 4) CURSO / TIPO DE TRABALHO
    # ============================================

    if identificados["curso"] is None:

        texto_norm = remover_acentos(texto_lower)
        for c in cursos:
            if texto_norm == remover_acentos(c.lower()):
                return "curso"

    if "curso de " in texto_norm or "bacharelado em" in texto_norm:
        if len(texto_original.split()) >= 2:
            return "curso"

    # ============================================
    # 5) AUTOR (regras MUITO fortes)
    # 2 a 4 nomes, iniciando com maiúscula, sem números
    # ============================================
    if identificados["autor"] is None:
        palavras = texto_original.split()

        # 2 a 4 palavras
        if 2 <= len(palavras) <= 4:
            # Todas começam com Maiúscula
            if all(p[0].isupper() and p[1:].islower() for p in palavras if len(p) > 1):
                # Não contém palavras de instituição
                if not any(k in texto_lower for k in instituicao_keywords):
                    # Não contém curso
                    if not any(k in texto_lower for k in curso_keywords):
                        # Não contém cidade
                        if not eh_cidade(texto_original, cidades_bahia):
                            return "autor"

    # ============================================
    # 6) TÍTULO (≥ 40 caracteres e ≥ 6 palavras)
    # ============================================
    if identificados["titulo"] is None:
        if len(texto_original) >= 40 and len(texto_original.split()) >= 6:
            return "titulo"

    # ============================================
    # 7) SUBTÍTULO (só pode vir após título)
    # ============================================
    if identificados["titulo"] is not None and identificados["subtitulo"] is None:
        if 10 <= len(texto_original) <= 80:
            return "subtitulo"

    # ============================================
    # Nada identificado
    # ============================================
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

    for _ in range(8):  # repete para refinar a identificação
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

def eh_cidade(texto):
    texto_norm = remover_acentos(texto.strip().lower())

    # CIDADES EXATAS
    for c in cidades_bahia:
        if remover_acentos(c.lower()) == texto_norm:
            return True

    # CIDADE + UF (ex: Salvador - BA, Feira de Santana BA)
    if texto_norm.endswith(" ba") or texto_norm.endswith("-ba"):
        return True

    return False

print(eh_cidade("Feira de Santana"))

def eh_instituicao(texto):
    texto_lower = texto.lower()

    for inst in instituicao:
        inst_lower = inst.lower().strip()

        # Evita falso positivo parcial (ex: "centro" dentro de "concentração")
        if f" {inst_lower} " in f" {texto_lower} ":
            return True

        if texto_lower.startswith(inst_lower):
            return True

    return False


def eh_curso(txt): 
    t = remover_acentos(txt.strip().lower()) 
    for c in cursos: 
        if remover_acentos(c.lower()) in t: 
            return True
    return False

def eh_autor(t):
    # Nome de pessoa: 2 a 5 palavras com iniciais maiúsculas
    return re.fullmatch(r"([A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+)(\s+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+){1,4}", t) is not None

def classificar_linhas(linhas):
    """
    Recebe lista de linhas (texto bruto) e retorna dict com os primeiros
    valores detectados para: ano, instituicao, titulo, subtitulo, autor, cidade, curso
    """
    identificados = {
        "ano": None,
        "instituicao": None,
        "titulo": None,
        "subtitulo": None,
        "autor": None,
        "cidade": None,
        "curso": None
    }

    titulo_identificado = False
    resto = linhas[:]

    for _ in range(7):  # repete para refinar
        novos_resto = []
        for t in resto:
            s = t.strip()
            if not s:
                continue

            # 1 — ANO
            if identificados["ano"] is None and eh_ano(s):
                identificados["ano"] = s
                continue

            # 2 — INSTITUIÇÃO
            if identificados["instituicao"] is None and eh_instituicao(s):
                identificados["instituicao"] = s
                continue

            # 3 — CURSO (antes do autor para evitar confusão)
            if identificados["curso"] is None and eh_curso(s):
                identificados["curso"] = s
                continue

            # 4 — TÍTULO (usa estado)
            if identificados["titulo"] is None:
                eh, novo_estado = eh_titulo(s, titulo_identificado)
                if eh:
                    identificados["titulo"] = s
                    titulo_identificado = novo_estado
                    continue

            # 5 — SUBTÍTULO (só após título)
            if identificados["titulo"] and identificados["subtitulo"] is None:
                if eh_subtitulo(s, titulo_identificado):
                    identificados["subtitulo"] = s
                    continue

            # 6 — AUTOR
            if identificados["autor"] is None and eh_autor(s):
                identificados["autor"] = s
                continue

            # 7 — CIDADE
            if identificados["cidade"] is None and eh_cidade(s):
                identificados["cidade"] = s
                continue

            novos_resto.append(t)

        if len(novos_resto) == len(resto):
            break
        resto = novos_resto

    return identificados

def normalizar(texto):
    if not texto:
        return ""
    texto = texto.strip().lower()
    texto = unicodedata.normalize('NFD', texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = " ".join(texto.split())  # remove espaços duplicados
    return texto


from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def formatar_capa(doc):
    """
    Reconstrói e formata a área da capa na ordem:
    Instituição (opcional), Curso (opcional), Autor, Título,
    Subtítulo (opcional), Local, Ano.
    """

    # encontra índice do primeiro parágrafo que começa com "resumo"
    index_resumo = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text and p.text.strip().lower().startswith("resumo"):
            index_resumo = idx
            break
    limite = index_resumo if index_resumo is not None else len(doc.paragraphs)

    # monta lista de textos originais (apenas da região da capa)
    linhas = [p.text for p in doc.paragraphs[:limite]]

 

    # usa a função existente que retorna índices identificados
    try:
        identificados_indices = identificar_linhas_da_capa(doc)
    except Exception:
        identificados_indices = {}
        classificados = classificar_linhas(linhas)
        for chave, valor in classificados.items():
            if valor:
                try:
                    identificados_indices[chave] = linhas.index(valor)
                except ValueError:
                    identificados_indices[chave] = None
            else:
                identificados_indices[chave] = None

    # extrai texto conforme índices
    def txt(idx):
        return linhas[idx].strip() if (idx is not None and 0 <= idx < len(linhas) and linhas[idx].strip()) else None

    inst_txt = txt(identificados_indices.get("instituicao"))
    curso_txt = txt(identificados_indices.get("curso"))
    autor_txt = txt(identificados_indices.get("autor"))
    titulo_txt = txt(identificados_indices.get("titulo"))
    subt_txt = txt(identificados_indices.get("subtitulo"))
    cidade_txt = txt(identificados_indices.get("cidade"))
    ano_txt = txt(identificados_indices.get("ano"))

    # fallback para título
    if not titulo_txt:
        for i, l in enumerate(linhas):
            if l and eh_titulo(l, False)[0]:
                titulo_txt = l.strip()
                break

    # fallback autor
    if not autor_txt:
        for l in linhas:
            if l and eh_autor(l):
                autor_txt = l.strip()
                break

    # fallback ano
    if not ano_txt:
        for l in linhas:
            if l and eh_ano(l):
                ano_txt = l.strip()
                break

    # fallback cidade
    if not cidade_txt:
        for l in linhas:
            if l and eh_cidade(l, cidades_bahia):
                cidade_txt = l.strip()
                break

    # monta lista final ordenada
    capa_ordem = []
    if inst_txt:
        capa_ordem.append(("instituicao", inst_txt))
    if curso_txt:
        capa_ordem.append(("curso", curso_txt))
    if autor_txt:
        capa_ordem.append(("autor", autor_txt))
    if titulo_txt:
        capa_ordem.append(("titulo_capa", titulo_txt))
    if subt_txt and subt_txt != titulo_txt:
        capa_ordem.append(("subtitulo", subt_txt))
    if cidade_txt:
        capa_ordem.append(("cidade", cidade_txt))
    if ano_txt:
        capa_ordem.append(("ano", ano_txt))


    novos_pares = []
    usados = set()

    for tipo, texto in capa_ordem:
        if texto not in usados:
            novos_pares.append((tipo, texto))
            usados.add(texto)

    capa_ordem = novos_pares
    # ---------------------------------------------------------

    # encontra primeiro parágrafo não vazio
    first_idx = 0
    for idx in range(limite):
        if doc.paragraphs[idx].text.strip():
            first_idx = idx
            break

    # aplica formatação
    for offset, (tipo, texto) in enumerate(capa_ordem):
        target_idx = first_idx + offset
        p = doc.paragraphs[target_idx]
        p.text = texto.strip()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for run in p.runs:
            run.font.size = Pt(12)
            run.bold = False

        if tipo == "titulo_capa":
            texto_titulo = p.text  # mantém o texto ORIGINAL
            p.text = ""            # limpa

            run = p.add_run(texto_titulo.upper())  # exibe em MAIÚSCULO só no DOCX
            run.font.size = Pt(14)
            run.bold = True


    # limpa o restante
    start_clean = first_idx + len(capa_ordem)
    for idx in range(start_clean, limite):
        doc.paragraphs[idx].text = ""



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
    texto_limpo = texto.strip()
    if not texto_limpo:
        return False

    # Só existe subtítulo se o título já tiver sido identificado
    if not titulo_identificado:
        return False

    # Subtítulo NÃO pode estar totalmente em caixa alta
    if texto_limpo.isupper():
        return False

    # Subtítulo deve ter estrutura de frase (>= 3 palavras)
    if len(texto_limpo.split()) < 3:
        return False

    # Subtítulo deve conter letras minúsculas
    if texto_limpo.upper() == texto_limpo:
        return False

    # Não pode ser nome de autor
    if re.match(r'^[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+(\s+[A-ZÁÉÍÓÚÂÊÔÃÕ][a-záéíóúâêôãõç]+)+$', texto_limpo):
        return False

    # Não pode ser cidade simples (ex: "Feira de Santana")
    if texto_limpo.istitle() and len(texto_limpo.split()) <= 3:
        return False

    # Não pode ser ano
    if re.fullmatch(r'\d{4}', texto_limpo):
        return False

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
    global ja_tem_autor, ja_tem_cidade, ja_tem_instituicao
    global ja_tem_titulo_capa, ja_tem_curso, ja_tem_ano

    ja_tem_autor = False
    ja_tem_cidade = False
    ja_tem_instituicao = False
    ja_tem_titulo_capa = False
    ja_tem_curso = False
    ja_tem_ano = False
    texto = texto.strip()
    if not texto:
        return "vazio", titulo_identificado

    texto_lower = texto.lower()
    texto_norm = remover_acentos(texto_lower)

    # ============================================================
    # 1 — ANO  (vem antes de tudo)
    # ============================================================
    if not ja_tem_ano and eh_ano(texto):
        ja_tem_ano = True
        return "ano", titulo_identificado

    # ============================================================
    # 2 — CIDADE  (prioridade máxima)
    # ============================================================
    if not ja_tem_cidade and eh_cidade(texto):
        linha_curta = len(texto) <= 40
        termina_com_uf = texto_norm.endswith(" ba") or texto_norm.endswith(" bahia")
        somente_cidade = any(remover_acentos(c.lower()) == texto_norm for c in cidades_bahia)

        if linha_curta or termina_com_uf or somente_cidade:
            ja_tem_cidade = True
            return "cidade", titulo_identificado

    # ============================================================
    # 3 — AUTOR
    # ============================================================
    if not ja_tem_autor and eh_autor(texto):
        ja_tem_autor = True
        return "autor", titulo_identificado

    # ============================================================
    # 4 — ORIENTADOR
    # ============================================================
    if any(p in texto_lower for p in palavras_orientador):
        return "orientador", titulo_identificado

    # ============================================================
    # 5 — INSTITUIÇÃO
    # ============================================================
    if not ja_tem_instituicao and eh_instituicao(texto):
        ja_tem_instituicao = True
        return "instituicao", titulo_identificado

    # ============================================================
    # 6 — CURSO
    # ============================================================
    if not ja_tem_curso and eh_curso(texto):
        ja_tem_curso = True
        return "curso", titulo_identificado

    # ============================================================
    # 7 — TÍTULO DA CAPA
    # ============================================================
    eh_tit, novo_estado = eh_titulo(texto, titulo_identificado)
    if not ja_tem_titulo_capa and eh_tit:
        ja_tem_titulo_capa = True
        return "titulo_capa", novo_estado

    # 8 — SUBTÍTULO DA CAPA (linha logo após o título da capa)
    texto_sem_acentos = remover_acentos(texto)

    if ja_tem_titulo_capa and not ja_tem_autor:
    # Aceita frase maiúscula SEM acentos
        if texto_sem_acentos.isupper() and len(texto.split()) > 1:
          return "subtitulo_capa", titulo_identificado


    # ============================================================
    # 9 — TÍTULOS DO CORPO
    # ============================================================
    if re.match(r"^\d+\s+[A-ZÁÉÍÓÚÂÊÔÃÕ]", texto):
        return "titulo_principal", titulo_identificado

    if re.match(r"^\d+\.\d+(\.\d+)*\s+[A-ZÁÉÍÓÚÂÊÔÃÕ]", texto):
        return "subtitulo", titulo_identificado

    if eh_subtitulo(texto, titulo_identificado):
        return "subtitulo", titulo_identificado

    # ============================================================
    # 10 — PARÁGRAFO NORMAL
    # ============================================================
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
    global ja_tem_autor, ja_tem_cidade, ja_tem_instituicao
    global ja_tem_titulo_capa, ja_tem_curso, ja_tem_ano

    ja_tem_autor = False
    ja_tem_cidade = False
    ja_tem_instituicao = False
    ja_tem_titulo_capa = False
    ja_tem_curso = False
    ja_tem_ano = False

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
