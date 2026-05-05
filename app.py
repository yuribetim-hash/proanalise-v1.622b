import streamlit as st
import os
import json
from io import BytesIO
from datetime import datetime, timedelta
from docxtpl import DocxTemplate, RichText
import hashlib
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
import time
import re
from datetime import timezone, timedelta

st.set_page_config(
    page_title="Proanalise v1.622",
    page_icon="📐",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================
# CONFIGURAÇÃO DE FUSO HORÁRIO (BRASÍLIA)
# ============================================
# Definir fuso horário de Brasília (UTC-3)
BRASILIA_TZ = timezone(timedelta(hours=-3))

def get_horario_brasilia():
    """Retorna a data/hora atual no fuso horário de Brasília"""
    return datetime.now(BRASILIA_TZ)

def formatar_horario_brasilia(dt=None):
    """Formata a data/hora no padrão brasileiro"""
    if dt is None:
        dt = get_horario_brasilia()
    return dt.strftime("%d/%m/%Y %H:%M:%S")

# ============================================
# INICIALIZAÇÃO DE ESTADOS (PRIMEIRO)
# ============================================
def inicializar_estados():
    if "dados_antigos" not in st.session_state:
        st.session_state["dados_antigos"] = None
    if "etapa" not in st.session_state:
        st.session_state["etapa"] = "1. Protocolo"
    if "protocolo" not in st.session_state:
        st.session_state["protocolo"] = ""
    if "tipo" not in st.session_state:
        st.session_state["tipo"] = "Loteamento"
    if "tipo_analise" not in st.session_state:
        st.session_state["tipo_analise"] = "Aceite urbanístico"
    if "interessado" not in st.session_state:
        st.session_state["interessado"] = ""
    if "n_lotes" not in st.session_state:
        st.session_state["n_lotes"] = 1
    if "matriculas" not in st.session_state:
        st.session_state["matriculas"] = ""
    if "analista" not in st.session_state:
        st.session_state["analista"] = ""
    if "matricula_analista" not in st.session_state:
        st.session_state["matricula_analista"] = ""
    if "setor" not in st.session_state:
        st.session_state["setor"] = ""
    if "n_analise" not in st.session_state:
        st.session_state["n_analise"] = ""
    if "pendencias_manuais" not in st.session_state:
        st.session_state["pendencias_manuais"] = {}
    if "respostas_temp" not in st.session_state:
        st.session_state["respostas_temp"] = {}
    if "observacoes_temp" not in st.session_state:
        st.session_state["observacoes_temp"] = {}
    if "respostas_analise" not in st.session_state:
        st.session_state["respostas_analise"] = {}
    if "observacoes_analise" not in st.session_state:
        st.session_state["observacoes_analise"] = {}
    if "pendencias_analise" not in st.session_state:
        st.session_state["pendencias_analise"] = {}
    if "analise_ativa" not in st.session_state:
        st.session_state["analise_ativa"] = False
    if "analise_concluida" not in st.session_state:
        st.session_state["analise_concluida"] = False
    if "tempo_inicio" not in st.session_state:
        st.session_state["tempo_inicio"] = None
    if "tempo_fim" not in st.session_state:
        st.session_state["tempo_fim"] = None
    if "marcadas_revisao" not in st.session_state:
        st.session_state["marcadas_revisao"] = set()
    if "anotacoes_pessoais" not in st.session_state:
        st.session_state["anotacoes_pessoais"] = {}
    if "ultimo_backup" not in st.session_state:
        st.session_state["ultimo_backup"] = get_horario_brasilia()
    if "tema_mode" not in st.session_state:
        st.session_state["tema_mode"] = "claro"
    if "botao_flutuante" not in st.session_state:
        st.session_state["botao_flutuante"] = False
    if "logado" not in st.session_state:
        st.session_state["logado"] = False
    if "usuario" not in st.session_state:
        st.session_state["usuario"] = ""
    if "usuario_info" not in st.session_state:
        st.session_state["usuario_info"] = None
    if "papel" not in st.session_state:
        st.session_state["papel"] = ""
    if "nivel" not in st.session_state:
        st.session_state["nivel"] = 1
    if "scroll_to" not in st.session_state:
        st.session_state["scroll_to"] = None
    if "mostrar_painel_comando" not in st.session_state:
        st.session_state["mostrar_painel_comando"] = False

# Chamar inicialização imediatamente
inicializar_estados()

HASH_SALT = "proanalise_salt_2024"

# ============================================
# FUNÇÕES DE BACKUP
# ============================================
def fazer_backup_automatico():
    if not st.session_state.get("analise_ativa", False):
        return False
    
    agora = get_horario_brasilia()
    diff = (agora - st.session_state["ultimo_backup"]).total_seconds()
    
    if diff >= 300:
        pasta_backup = os.path.join("dados", "backups")
        os.makedirs(pasta_backup, exist_ok=True)
        
        timestamp = agora.strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(pasta_backup, f"backup_{timestamp}.json")
        
        dados_backup = {
            "protocolo": st.session_state.get("protocolo", ""),
            "respostas_analise": st.session_state.get("respostas_analise", {}),
            "observacoes_analise": st.session_state.get("observacoes_analise", {}),
            "pendencias_analise": st.session_state.get("pendencias_analise", {}),
            "etapa": st.session_state.get("etapa", ""),
            "marcadas_revisao": list(st.session_state.get("marcadas_revisao", set())),
            "anotacoes_pessoais": st.session_state.get("anotacoes_pessoais", {}),
            "data_backup": timestamp
        }
        
        with open(backup_file, "w", encoding="utf-8") as f:
            json.dump(dados_backup, f, indent=4, ensure_ascii=False)
        
        backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")])
        if len(backups) > 10:
            for old_backup in backups[:-10]:
                os.remove(os.path.join(pasta_backup, old_backup))
        
        st.session_state["ultimo_backup"] = agora
        return True
    return False

def salvar_backup_manual():
    pasta_backup = os.path.join("dados", "backups")
    os.makedirs(pasta_backup, exist_ok=True)
    
    timestamp = get_horario_brasilia().strftime("%Y%m%d_%H%M%S")
    backup_file = os.path.join(pasta_backup, f"backup_{timestamp}.json")
    
    dados_backup = {
        "protocolo": st.session_state.get("protocolo", ""),
        "respostas_analise": st.session_state.get("respostas_analise", {}),
        "observacoes_analise": st.session_state.get("observacoes_analise", {}),
        "pendencias_analise": st.session_state.get("pendencias_analise", {}),
        "etapa": st.session_state.get("etapa", ""),
        "marcadas_revisao": list(st.session_state.get("marcadas_revisao", set())),
        "anotacoes_pessoais": st.session_state.get("anotacoes_pessoais", {}),
        "data_backup": timestamp
    }
    
    with open(backup_file, "w", encoding="utf-8") as f:
        json.dump(dados_backup, f, indent=4, ensure_ascii=False)
    
    backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")])
    if len(backups) > 10:
        for old_backup in backups[:-10]:
            os.remove(os.path.join(pasta_backup, old_backup))
    
    return backup_file

def restaurar_analise():
    pasta_backup = os.path.join("dados", "backups")
    if not os.path.exists(pasta_backup):
        return None
    
    backups = sorted([f for f in os.listdir(pasta_backup) if f.startswith("backup_")], reverse=True)
    if not backups:
        return None
    
    opcoes = {}
    for b in backups[:10]:
        data_str = b.replace("backup_", "").replace(".json", "")
        try:
            data_obj = datetime.strptime(data_str, "%Y%m%d_%H%M%S")
            opcoes[b] = data_obj.strftime("%d/%m/%Y %H:%M:%S")
        except:
            opcoes[b] = data_str
    
    return opcoes

# ============================================
# FUNÇÕES DE PERMISSÕES E USUÁRIOS
# ============================================
def carregar_usuarios(caminho="usuarios.txt"):
    if not os.path.exists(caminho):
        st.error("Arquivo usuarios.txt não encontrado.")
        st.stop()
    
    usuarios = {}
    with open(caminho, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            partes = linha.split(";")
            if len(partes) >= 4:
                usuario, senha, papel, nivel = partes[0], partes[1], partes[2], partes[3]
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": papel.strip(),
                    "nivel": int(nivel.strip())
                }
            elif len(partes) == 3:
                usuario, senha, papel = partes
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": papel.strip(),
                    "nivel": 2 if "Sênior" in papel else 1
                }
            else:
                usuario, senha = partes[0], partes[1]
                usuarios[usuario.strip()] = {
                    "senha": senha.strip(),
                    "papel": "Analista",
                    "nivel": 2
                }
    return usuarios

def tem_permissao(nivel_necessario):
    if "usuario_info" not in st.session_state or st.session_state["usuario_info"] is None:
        return False
    return st.session_state["usuario_info"]["nivel"] >= nivel_necessario

def pode_ver_menu(menu_item):
    niveis_necessarios = {
        "1. Protocolo": 1,
        "2. Analista": 1,
        "3. Análise": 1,
        "4. Revisão": 1,
        "5. Gerar parecer": 1,
        "6. Dashboard": 2,
        "7. Comparador": 3
    }
    return tem_permissao(niveis_necessarios.get(menu_item, 1))

# ============================================
# FUNÇÕES DE CARREGAMENTO DE PERGUNTAS DINÂMICAS
# ============================================
def carregar_perguntas_por_tipo(tipo_empreendimento, tipo_analise):
    """Carrega o arquivo de perguntas correto baseado no tipo de empreendimento e análise"""
    
    # Mapeamento para o nome do arquivo
    if tipo_empreendimento == "Loteamento":
        if tipo_analise == "Aceite urbanístico":
            arquivo = "perguntas/loteamento_aceite.txt"
        else:
            arquivo = "perguntas/loteamento_alvara.txt"
    else:  # Condomínio fechado de lotes
        if tipo_analise == "Aceite urbanístico":
            arquivo = "perguntas/condominio_aceite.txt"
        else:
            arquivo = "perguntas/condominio_alvara.txt"
    
    # Tentar carregar o arquivo específico
    if os.path.exists(arquivo):
        return carregar_perguntas_txt(arquivo)
    
    # Fallback para o arquivo padrão
    if os.path.exists("perguntas.txt"):
        return carregar_perguntas_txt("perguntas.txt")
    
    st.warning(f"Arquivo de perguntas não encontrado: {arquivo}")
    return []

def carregar_perguntas_txt(caminho):
    perguntas = []
    bloco = {}
    
    if not os.path.exists(caminho):
        return []
    
    with open(caminho, "r", encoding="utf-8") as f:
        linhas = f.readlines()
    
    for linha in linhas:
        linha = linha.strip()
        if not linha:
            if bloco:
                perguntas.append(bloco)
                bloco = {}
            continue
        if linha.startswith("GRUPO:"):
            bloco["grupo"] = linha.replace("GRUPO:", "").strip()
        elif linha.startswith("ID:"):
            bloco["id"] = linha.replace("ID:", "").strip()
        elif linha.startswith("PERGUNTA:"):
            bloco["pergunta"] = linha.replace("PERGUNTA:", "").strip()
        elif linha.startswith("OPCOES:"):
            bloco["opcoes"] = [op.strip() for op in linha.replace("OPCOES:", "").strip().split(";")]
        elif linha.startswith("CONFORMES:"):
            bloco["conformes"] = [op.strip() for op in linha.replace("CONFORMES:", "").strip().split(";")]
        elif linha.startswith("REGRA_"):
            chave, valor = linha.split(":", 1)
            resposta = chave.replace("REGRA_", "").strip()
            bloco.setdefault("regras", {})[resposta] = {"texto": valor.strip()}
    
    if bloco:
        perguntas.append(bloco)
    
    return perguntas

# ============================================
# FUNÇÕES DE MÚLTIPLOS ANALISTAS
# ============================================
def get_pasta_protocolo(protocolo):
    protocolo_limpo = protocolo.replace("/", "-").strip()
    return os.path.join("dados", protocolo_limpo)

def salvar_analise_analista(protocolo, analista, papel, respostas, observacoes, pendencias):
    pasta = get_pasta_protocolo(protocolo)
    os.makedirs(pasta, exist_ok=True)
    
    analista_hash = hashlib.md5(f"{analista}_{papel}_{HASH_SALT}".encode()).hexdigest()[:8]
    
    arquivo = os.path.join(pasta, f"analise_{analista_hash}_{papel}.json")
    registro = {
        "protocolo": protocolo,
        "analista": analista,
        "papel": papel,
        "data": formatar_horario_brasilia(),
        "respostas": respostas,
        "observacoes": observacoes,
        "pendencias": pendencias,
        "hash": analista_hash
    }
    
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(registro, f, indent=4, ensure_ascii=False)
    
    return analista_hash

def carregar_analises_analistas(protocolo):
    pasta = get_pasta_protocolo(protocolo)
    if not os.path.exists(pasta):
        return []
    
    analises = []
    for arquivo in os.listdir(pasta):
        if arquivo.startswith("analise_") and arquivo.endswith(".json"):
            with open(os.path.join(pasta, arquivo), "r", encoding="utf-8") as f:
                analises.append(json.load(f))
    
    return analises

def comparar_analises(analises):
    if len(analises) < 2:
        return None
    
    resultado = {
        "total_diferencas": 0,
        "diferencas_por_pergunta": {},
        "analistas": [a["analista"] for a in analises],
        "papeis": [a["papel"] for a in analises],
        "data_analises": [a["data"] for a in analises]
    }
    
    todas_perguntas = set()
    for analise in analises:
        todas_perguntas.update(analise["respostas"].keys())
    
    for pergunta in todas_perguntas:
        respostas_analistas = {}
        for analise in analises:
            resp = analise["respostas"].get(pergunta, "Não respondida")
            respostas_analistas[analise["analista"]] = resp
        
        valores_unicos = set(respostas_analistas.values())
        if len(valores_unicos) > 1:
            resultado["diferencas_por_pergunta"][pergunta] = respostas_analistas
            resultado["total_diferencas"] += 1
    
    return resultado

def render_comparador_analises(protocolo_atual):
    if not tem_permissao(3):
        st.error("❌ Acesso negado! Apenas Analistas Responsáveis podem acessar o Comparador.")
        return
    
    st.subheader("🔍 Comparador de Análises")
    
    analises = carregar_analises_analistas(protocolo_atual)
    
    if not analises:
        st.info("Nenhuma análise de outro analista encontrada para este protocolo.")
        return
    
    st.write(f"**Total de análises encontradas:** {len(analises)}")
    
    for a in analises:
        st.write(f"- {a['analista']} ({a['papel']}) - {a['data']}")
    
    if st.button("Comparar Análises", use_container_width=True):
        comparacao = comparar_analises(analises)
        
        if comparacao and comparacao["total_diferencas"] > 0:
            st.warning(f"⚠️ **{comparacao['total_diferencas']} divergências encontradas**")
            
            for pergunta, respostas in comparacao["diferencas_por_pergunta"].items():
                with st.expander(f"📌 Pergunta ID: {pergunta}"):
                    for analista, resposta in respostas.items():
                        st.write(f"**{analista}:** {resposta}")
        else:
            st.success("✅ Todas as análises estão consistentes!")

# ============================================
# FUNÇÕES DE MÉTRICAS
# ============================================
def calcular_tempo_medio_analise():
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return None
    
    tempos = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
            if len(analises) >= 2:
                analises.sort()
                primeira = analises[0]
                ultima = analises[-1]
                
                with open(os.path.join(protocolo_path, primeira), "r", encoding="utf-8") as f:
                    data_primeira = datetime.strptime(json.load(f)["data"], "%d/%m/%Y")
                with open(os.path.join(protocolo_path, ultima), "r", encoding="utf-8") as f:
                    data_ultima = datetime.strptime(json.load(f)["data"], "%d/%m/%Y")
                
                tempo = (data_ultima - data_primeira).days
                tempos.append(tempo)
    
    if tempos:
        return sum(tempos) / len(tempos)
    return None

def gerar_grafico_inconformidades(respostas, grupos_inconformes):
    if not grupos_inconformes:
        return None
    
    dados = []
    for grupo, itens in grupos_inconformes.items():
        dados.append({"Grupo": grupo, "Inconformidades": len(itens)})
    
    df = pd.DataFrame(dados)
    fig = px.bar(df, x="Grupo", y="Inconformidades", 
                 title="Inconformidades por Grupo",
                 color="Inconformidades",
                 color_continuous_scale="Reds")
    
    fig.update_layout(
        xaxis_title="Grupo",
        yaxis_title="Número de Inconformidades",
        showlegend=False,
        height=400
    )
    
    return fig

def gerar_grafico_tempo_analises():
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return None
    
    dados_tempo = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
            if analises:
                with open(os.path.join(protocolo_path, analises[0]), "r", encoding="utf-8") as f:
                    primeira = json.load(f)
                with open(os.path.join(protocolo_path, analises[-1]), "r", encoding="utf-8") as f:
                    ultima = json.load(f)
                
                data_inicio = datetime.strptime(primeira["data"], "%d/%m/%Y")
                data_fim = datetime.strptime(ultima["data"], "%d/%m/%Y")
                dias = (data_fim - data_inicio).days
                
                dados_tempo.append({
                    "Protocolo": protocolo_dir.replace("-", "/"),
                    "Dias de Análise": dias,
                    "Conclusão": ultima.get("conclusao", "Em análise")
                })
    
    if dados_tempo:
        df = pd.DataFrame(dados_tempo)
        fig = px.bar(df, x="Protocolo", y="Dias de Análise", 
                     title="Tempo de Análise por Protocolo",
                     color="Conclusão",
                     color_discrete_map={"FAVORÁVEL": "green", "DESFAVORÁVEL": "red"})
        fig.update_layout(xaxis_tickangle=-45, height=400)
        return fig
    return None

# ============================================
# FUNÇÕES DE BUSCA E FILTROS
# ============================================
def buscar_protocolos(termo_busca, filtro_status=None, filtro_analista=None, data_inicio=None, data_fim=None):
    pasta_dados = "dados"
    if not os.path.exists(pasta_dados):
        return []
    
    resultados = []
    for protocolo_dir in os.listdir(pasta_dados):
        protocolo_path = os.path.join(pasta_dados, protocolo_dir)
        if os.path.isdir(protocolo_path):
            if termo_busca.lower() in protocolo_dir.lower():
                status = "Em análise"
                analista = "Não informado"
                data_analise = datetime.fromtimestamp(os.path.getmtime(protocolo_path))
                
                analises = [f for f in os.listdir(protocolo_path) if f.startswith("AN") and f.endswith(".json")]
                if analises:
                    ultima = sorted(analises)[-1]
                    with open(os.path.join(protocolo_path, ultima), "r", encoding="utf-8") as f:
                        dados = json.load(f)
                        status = dados.get("conclusao", "Em análise")
                        analista = dados.get("analista", "Não informado")
                
                if filtro_status and filtro_status != "Todos":
                    if status != filtro_status:
                        continue
                if filtro_analista and filtro_analista != "Todos":
                    if analista != filtro_analista:
                        continue
                if data_inicio:
                    if data_analise.date() < data_inicio:
                        continue
                if data_fim:
                    if data_analise.date() > data_fim:
                        continue
                
                resultados.append({
                    "protocolo": protocolo_dir.replace("-", "/"),
                    "status": status,
                    "analista": analista,
                    "data_ultima": data_analise.strftime("%d/%m/%Y"),
                    "qtd_analises": len(analises)
                })
    
    return resultados

# ============================================
# FUNÇÕES DE PRODUTIVIDADE
# ============================================
def proxima_pergunta_nao_respondida(respostas, perguntas):
    for idx, p in enumerate(perguntas):
        resposta = respostas.get(p["id"])
        if resposta in ("", None, "Selecione..."):
            return p["id"], idx
    return None, None

# ============================================
# FUNÇÕES AUXILIARES
# ============================================
def resposta_preenchida(valor):
    return valor not in ("", None, "Selecione...")

# ============================================
# CARREGAR TEMA (CLARO/ESCURO)
# ============================================
def carregar_tema():
    if "tema_mode" not in st.session_state:
        st.session_state["tema_mode"] = "claro"
    
    tema_claro = {
        "cores": {
            "primaria": "#0a2a3a",
            "primaria_clara": "#1a5276",
            "primaria_muito_escura": "#051a24",
            "secundaria": "#2c6b96",
            "sucesso": "#0d6e2e",
            "erro": "#b42318",
            "alerta": "#b54708",
            "info": "#175cd3",
            "fundo_claro": "#f8fafd",
            "fundo_branco": "#ffffff",
            "texto_principal": "#1a1a1a",
            "texto_secundario": "#2c3e50",
            "texto_caption": "#666666",
            "texto_titulo": "#051a24",
            "fundo_app": "linear-gradient(135deg, #e8f0fe 0%, #d4e4fc 100%)",
            "sidebar_fundo": "linear-gradient(180deg, #0a2a3a 0%, #051a24 100%)"
        },
        "botoes": {
            "primario_fundo": "#0a5c2a",
            "primario_fundo_hover": "#0d6e2e",
            "secundario_fundo": "#0a2a3a",
            "secundario_fundo_hover": "#1a5276",
            "texto": "#ffffff"
        },
        "status": {
            "conforme": {"fundo": "#ecfdf3", "borda": "#067647", "texto": "#067647", "icone": "✅"},
            "inconforme": {"fundo": "#fef3f2", "borda": "#b42318", "texto": "#7a271a", "icone": "⛔"},
            "pendente": {"fundo": "#fffaeb", "borda": "#b54708", "texto": "#7a4a0a", "icone": "⏳"},
            "nao_se_enquadra": {"fundo": "#eff6ff", "borda": "#175cd3", "texto": "#0e4a8a", "icone": "ℹ️"}
        }
    }
    
    tema_escuro = {
        "cores": {
            "primaria": "#e8f0fe",
            "primaria_clara": "#d4e4fc",
            "primaria_muito_escura": "#c5d5e6",
            "secundaria": "#2c6b96",
            "sucesso": "#0f8a3a",
            "erro": "#f5c2c7",
            "alerta": "#fedf89",
            "info": "#c7d7fe",
            "fundo_claro": "#1e1e2e",
            "fundo_branco": "#2a2a3e",
            "texto_principal": "#e0e0e0",
            "texto_secundario": "#b0b0b0",
            "texto_caption": "#888888",
            "texto_titulo": "#ffffff",
            "fundo_app": "linear-gradient(135deg, #1a1a2e 0%, #0a0a15 100%)",
            "sidebar_fundo": "linear-gradient(180deg, #0a0a15 0%, #05050a 100%)"
        },
        "botoes": {
            "primario_fundo": "#0f8a3a",
            "primario_fundo_hover": "#0d6e2e",
            "secundario_fundo": "#2c6b96",
            "secundario_fundo_hover": "#1a5276",
            "texto": "#ffffff"
        },
        "status": {
            "conforme": {"fundo": "#1a3a2a", "borda": "#0f8a3a", "texto": "#90ee90", "icone": "✅"},
            "inconforme": {"fundo": "#3a1a1a", "borda": "#f5c2c7", "texto": "#f5c2c7", "icone": "⛔"},
            "pendente": {"fundo": "#3a2a1a", "borda": "#fedf89", "texto": "#fedf89", "icone": "⏳"},
            "nao_se_enquadra": {"fundo": "#1a2a3a", "borda": "#c7d7fe", "texto": "#c7d7fe", "icone": "ℹ️"}
        }
    }
    
    if st.session_state["tema_mode"] == "escuro":
        return tema_escuro
    return tema_claro

# ============================================
# RENDERIZAÇÃO DO TEMA CSS
# ============================================
tema = carregar_tema()

css_tema = f"""
<style>
    .stApp {{ background: {tema["cores"]["fundo_app"]}; }}
    .main > div {{ background-color: {tema["cores"]["fundo_branco"]}; border-radius: 12px; padding: 1rem; }}
    
    [data-testid="stSidebar"] {{ background: {tema["cores"]["sidebar_fundo"]}; }}
    [data-testid="stSidebar"] * {{ color: {tema["botoes"]["texto"]} !important; }}
    
    h1 {{ color: {tema["cores"]["texto_titulo"]} !important; font-weight: 700 !important; font-size: 24px !important; }}
    h2, h3, h4 {{ color: {tema["cores"]["primaria"]} !important; font-weight: 600 !important; }}
    
    .stCaption {{
        color: {tema["cores"]["texto_caption"]} !important;
        font-size: 14px !important;
    }}
    
    /* CORREÇÃO DO CURSOR PISCANTE */
    .stTextInput input, .stTextArea textarea, .stNumberInput input {{
        background-color: {tema["cores"]["fundo_branco"]} !important;
        color: {tema["cores"]["texto_principal"]} !important;
        border: 1px solid {tema["cores"]["primaria_clara"]} !important;
        border-radius: 6px !important;
        caret-color: {tema["cores"]["primaria"]} !important;
    }}
    
    .stTextInput input:focus, .stTextArea textarea:focus, .stNumberInput input:focus {{
        border-color: {tema["cores"]["primaria"]} !important;
        box-shadow: 0 0 0 2px rgba(26, 82, 118, 0.2) !important;
        outline: none !important;
    }}
    
    .stTextInput label, .stSelectbox label, .stTextArea label, .stNumberInput label {{
        color: {tema["cores"]["primaria_clara"]} !important;
        font-weight: 600 !important;
    }}
    
    .stSelectbox select {{
        background-color: {tema["cores"]["fundo_branco"]} !important;
        color: {tema["cores"]["texto_principal"]} !important;
        border: 1px solid {tema["cores"]["primaria_clara"]} !important;
        border-radius: 6px !important;
        font-size: 14px !important;
    }}
    
    div[data-baseweb="menu"] {{
        background-color: #1a1a2e !important;
        border: 1px solid #2c6b96 !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.3
