from flask import Flask, render_template, request, redirect, url_for, flash
# =================================================================
# 🔑 NOVAS IMPORTAÇÕES DE SEGURANÇA E LOGIN
# =================================================================
import os
from dotenv import load_dotenv # <-- ADICIONADO: Para ler o .env
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from werkzeug.security import generate_password_hash, check_password_hash

import gspread
import json
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread.exceptions import WorksheetNotFound, APIError


# =================================================================
# 🚀 INICIALIZAÇÃO E CONFIGURAÇÃO
# =================================================================

# --- Carregar Variáveis de Ambiente (Chave Mestra) ---
# ISSO É CRUCIAL para ler FLASK_SECRET e GOOGLE_CREDENTIALS do .env
load_dotenv() 

# --- Configuração do Flask ---
app = Flask(__name__)
# O os.getenv AGORA LERÁ o FLASK_SECRET do seu arquivo .env
app.secret_key = os.getenv("FLASK_SECRET", "um-segredo-muito-forte-e-aleatorio-para-dev")


# =================================================================
# 🔑 CONFIGURAÇÃO DE LOGIN
# =================================================================

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login' 
login_manager.login_message_category = 'info'


SENHA_HASH = 'pbkdf2:sha256:600000$pXOI19hbtiyGAltI$60a3eb2c1e13ae1b3448cf06269a6713587b61f1aa57e792dde09d3ea599dfe4'
USUARIOS = {
    "smith": {
        "id": "smith",
        "password_hash": SENHA_HASH, 
        "name": "Smith"
    }
}

class User(UserMixin):
    def __init__(self, user_id):
        self.id = user_id
        self.name = USUARIOS.get(user_id, {}).get('name', 'Usuário Desconhecido')

@login_manager.user_loader
def load_user(user_id):
    if user_id in USUARIOS:
        return User(user_id)
    return None

# =================================================================
# ⚙️ CONFIGURAÇÃO DO GOOGLE SHEETS API
# =================================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# =========================================================
# 🔒 AQUI É A ÚNICA MUDANÇA: usar SOMENTE .env
# =========================================================

# Carrega as credenciais usando a variável de ambiente (do .env)
google_json = os.getenv("GOOGLE_CREDENTIALS")

if not google_json:
    # Sem fallback, falha se não tiver .env configurado
    raise RuntimeError("ERRO: Variável de ambiente GOOGLE_CREDENTIALS não encontrada. Defina-a no arquivo .env.")

try:
    INFO = json.loads(google_json)
    CREDS = Credentials.from_service_account_info(INFO, scopes=SCOPES)
    CLIENT = gspread.authorize(CREDS)
    SHEET = CLIENT.open("Controle de Expedição")
except Exception as e:
    print(f"ERRO ao inicializar o Google Sheets API: {e}")
    exit(1)

# Abertura das Worksheets principais
try:
    PED_WS = SHEET.worksheet("Pedidos")
    RET_WS = SHEET.worksheet("Retiradas")
except Exception as e:
    print(f"ERRO ao abrir worksheets principais: {e}")
    exit(1)


# --- Estrutura das Planilhas ---
try:
    PED_HEADERS = PED_WS.row_values(1)
except Exception:
    PED_HEADERS = ["Pedido", "Solicitante", "Setor", "Status", "Data Pedido", "Retirado por", "Retirado em", "Obs-geral"]

# --- Criação/Abertura de abas auxiliares com verificação ---
def inicializa_abas_auxiliares():
    global ITENS_WS, LOG_WS, VALIDACAO_WS, ITENS_HEADERS, LOG_HEADERS, VALIDACAO_HEADERS
    
    # 1. Aba Itens
    try:
        ITENS_WS = SHEET.worksheet("Itens")
        ITENS_HEADERS = ITENS_WS.row_values(1)
    except WorksheetNotFound:
        print("Criando aba 'Itens'...")
        ITENS_WS = SHEET.add_worksheet(title="Itens", rows=1000, cols=10)
        ITENS_HEADERS = ["Pedido", "Item", "Material", "Qtde", "Obs-item", "Solicitante", "Setor", "Data"]
        ITENS_WS.append_row(ITENS_HEADERS)

    # 2. Aba Histórico
    try:
        LOG_WS = SHEET.worksheet("Histórico")
        LOG_HEADERS = LOG_WS.row_values(1)
    except WorksheetNotFound:
        print("Criando aba 'Histórico'...")
        LOG_WS = SHEET.add_worksheet(title="Histórico", rows=1000, cols=10)
        LOG_HEADERS = ["Data/Hora", "Pedido", "Tipo", "Usuário", "Motivo", "Dados Anteriores"]
        LOG_WS.append_row(LOG_HEADERS)
        
    # 3. Aba Validação (Material e Setor)
    try:
        VALIDACAO_WS = SHEET.worksheet("Validacao")
        VALIDACAO_HEADERS = VALIDACAO_WS.row_values(1)
    except WorksheetNotFound:
        print("Criando aba 'Validacao'...")
        VALIDACAO_WS = SHEET.add_worksheet(title="Validacao", rows=1000, cols=5)
        VALIDACAO_HEADERS = ["Material", "Setor", "Solicitante", "Unidade", "Obs-Geral-Padrao"]
        VALIDACAO_WS.append_row(VALIDACAO_HEADERS)
        
    VALIDACAO_HEADERS = VALIDACAO_WS.row_values(1)

inicializa_abas_auxiliares()

# Função auxiliar para obter o índice de uma coluna
def get_col_index(headers, col_name):
    """Retorna o índice da coluna (1-base) para uso em gspread.update_cell."""
    try:
        return headers.index(col_name) + 1
    except ValueError:
        raise ValueError(f"Coluna '{col_name}' não encontrada nos cabeçalhos: {headers}")

# --- Funções Auxiliares ---

def carregar_listas_validacao():
    """Busca listas de dados únicos para preenchimento de campos (Setor e Material)."""
    try:
        dados = VALIDACAO_WS.get_all_records()
        materiais = sorted(list(set(d.get("Material") for d in dados if d.get("Material"))))
        setores = sorted(list(set(d.get("Setor") for d in dados if d.get("Setor"))))
        return {"setores": setores, "materiais": materiais}
    except Exception as e:
        print(f"Erro ao carregar listas de validação: {e}")
        return {"setores": [], "materiais": []}

def proximo_id():
    """Gera o próximo ID do pedido no formato EXP-AAAA-XXXXXX."""
    try:
        ids = PED_WS.col_values(get_col_index(PED_HEADERS, "Pedido"))[1:]
    except Exception:
        ids = []

    nums = [int(i.split("-")[-1]) for i in ids if i and i.startswith("EXP-")]
    
    ano_atual = datetime.now().year
    return f"EXP-{ano_atual}-{str(max(nums, default=0) + 1).zfill(6)}"

def append_atomico(ws, linha):
    """Tenta adicionar uma linha na planilha, com retries em caso de APIError."""
    for _ in range(3):
        try:
            ws.append_row(linha)
            return
        except APIError as e:
            print(f"APIError ao gravar, tentando novamente... Erro: {e}")
            continue
    raise RuntimeError("Falha persistente ao gravar na planilha após 3 tentativas.")

def append_itens(num_pedido, itens, solicitante, setor, data_pedido):
    """Adiciona os itens de um pedido na aba 'Itens'."""
    for idx, it in enumerate(itens, start=1):
        try:
            qtde = int(it["qtde"])
        except (ValueError, TypeError):
            qtde = 0 
            
        append_atomico(ITENS_WS, [
            num_pedido,
            idx,
            it["material"].strip() if it.get("material") else "N/A",
            qtde,
            it["obs"].strip() if it.get("obs") else "",
            solicitante,
            setor,
            data_pedido
        ])

def grava_log(tipo, pedido, usuario, motivo="", dados_anteriores=None):
    """Grava um registro de ação (log) na aba 'Histórico'."""
    append_atomico(LOG_WS, [
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        pedido,
        tipo,
        usuario,
        motivo,
        str(dados_anteriores) if dados_anteriores else ""
    ])


# =================================================================
# 🔑 ROTAS DE AUTENTICAÇÃO
# =================================================================

@app.route('/login', methods=['GET', 'POST'])
def login():
    """Rota para login do usuário."""
    if current_user.is_authenticated:
        return redirect(url_for('painel'))

    if request.method == 'POST':
        
        # ⚠️ CORREÇÃO DE SEGURANÇA E ROBUSTEZ:
        # Normaliza o username removendo espaços e convertendo para minúsculas
        # Garante que ' Admin ' e 'ADMIN' sejam tratados como 'admin'
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password')
        
        # Omitimos o if/else de validação de campo vazio (já que .get() não falha)
        
        if username in USUARIOS:
            user_data = USUARIOS[username]
            
            # 1. Verificar a senha (check_password_hash)
            if check_password_hash(user_data['password_hash'], password):
                user = User(username)
                login_user(user)
                flash(f'Bem-vindo(a), {user.name}!', 'success')
                
                # 2. Redireciona para a página de onde o usuário veio
                next_page = request.args.get('next')
                return redirect(next_page or url_for('painel'))
            else:
                # Falha: Senha errada
                flash('Senha incorreta.', 'danger')
        else:
            # Falha: Usuário não encontrado no dicionário
            flash('Usuário não encontrado.', 'danger')

    # A rota de login deve renderizar o template 'login.html'
    return render_template('login.html', title='Login')

@app.route('/logout')
@login_required 
def logout():
    """Rota para logout do usuário."""
    logout_user()
    return redirect(url_for('login'))

# =================================================================
# 🗺️ ROTAS PROTEGIDAS
# =================================================================

@app.route("/")
@login_required 
def painel():
    """Exibe o painel principal com a lista de todos os pedidos."""
    pedidos = PED_WS.get_all_records()
    
    for p in pedidos:
        p["RetiradoPor"] = p.get("Retirado por", "-")
        p["RetiradoEm"]  = p.get("Retirado em", "-")
        
    return render_template("painel.html", pedidos=pedidos)

@app.route("/novo", methods=["GET", "POST"])
@login_required 
def criar_pedido():
    """Rota para criar um novo pedido e seus itens associados."""
    listas = carregar_listas_validacao() 
    
    if request.method == "POST":
        try:
            solicitante = request.form["solicitante"].strip()
            setor = request.form["setor"].strip()
            try:
                total_itens = int(request.form["total_itens"])
            except (KeyError, ValueError):
                total_itens = 0
            
            if not solicitante or not setor:
                flash("Solicitante e Setor são obrigatórios.", "erro")
                return render_template("criar_pedido.html", listas=listas)
                
            data_sep = datetime.now().strftime("%d/%m/%Y %H:%M")
            num_pedido = proximo_id()

            # 2. Gravar cabeçalho do pedido
            pedido_data = [
                num_pedido,
                solicitante,
                setor,
                "EM EXPEDIÇÃO",
                data_sep,
                "",  # Retirado por
                "",  # Retirado em
                request.form.get("obs_geral", "").strip()
            ]
            
            if len(pedido_data) > len(PED_HEADERS):
                pedido_data = pedido_data[:len(PED_HEADERS)]
            elif len(pedido_data) < len(PED_HEADERS):
                pedido_data.extend([""] * (len(PED_HEADERS) - len(pedido_data)))

            append_atomico(PED_WS, pedido_data)

            # 3. Preparar e gravar itens
            itens = []
            for i in range(total_itens):
                material = request.form.get(f"material_{i}")
                qtde = request.form.get(f"qtde_{i}")
                
                if material and qtde:
                    try:
                        int(qtde)
                        itens.append({
                            "material": material,
                            "qtde": qtde,
                            "obs": request.form.get(f"obs_item_{i}", "")
                        })
                    except ValueError:
                        flash(f"Quantidade inválida para o item {i+1}. O item foi ignorado.", "aviso")
                            
            if not itens:
                flash("Nenhum item válido foi adicionado ao pedido.", "aviso")
                
            append_itens(num_pedido, itens, solicitante, setor, data_sep)
            
            grava_log("CRIAÇÃO", num_pedido, current_user.id) 
            
            flash("Pedido + itens criados com sucesso!", "sucesso")
            return redirect(url_for("painel"))
            
        except RuntimeError as e:
            flash(f"Erro ao gravar dados na planilha: {e}", "erro")
            return render_template("criar_pedido.html", listas=listas)
        except Exception as e:
            flash(f"Erro inesperado ao criar pedido: {e}", "erro")
            return render_template("criar_pedido.html", listas=listas)
            
    return render_template("criar_pedido.html", listas=listas)

# --- EDITAR CABEÇALHO ---
@app.route("/editar/<numero>", methods=["GET", "POST"])
@login_required 
def editar_pedido(numero):
    """Rota para editar os dados de cabeçalho (Solicitante, Setor, Obs-geral) de um pedido."""
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
        
    try:
        reg = PED_WS.row_values(cel.row)
        reg_dict = dict(zip(PED_HEADERS, reg))
    except Exception:
        flash("Erro ao carregar dados do pedido.", "erro")
        return redirect(url_for("painel"))
        
    listas = carregar_listas_validacao()

    itens_data = ITENS_WS.get_all_records()
    itens = [i for i in itens_data if i.get("Pedido") == numero]
    
    if request.method == "POST":
        try:
            usuario_log = current_user.id 
            
            novo_solicitante = request.form["solicitante"].strip()
            novo_setor = request.form["setor"].strip()
            novo_obs = request.form.get("obs_geral", "").strip()

            grava_log("EDIÇÃO-CABEÇALHO", numero, usuario_log, dados_anteriores=reg_dict)
            
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Solicitante"), novo_solicitante)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Setor"), novo_setor)
            
            if "Obs-geral" in PED_HEADERS:
                PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Obs-geral"), novo_obs)
            
            flash("Cabeçalho do Pedido alterado.", "sucesso")
            return redirect(url_for("painel"))
        except Exception as e:
            flash(f"Erro ao salvar alteração: {e}", "erro")
            return render_template("editar.html", pedido=numero, info=reg_dict, itens=itens, listas=listas)

    return render_template("editar.html", pedido=numero, info=reg_dict, itens=itens, listas=listas)

# --- EDITAR ITENS (NOVO) ---
@app.route("/editar_itens/<numero>")
@login_required 
def editar_itens(numero):
    """Exibe o formulário para edição dos itens de um pedido."""
    its = ITENS_WS.get_all_records()
    its = [i for i in its if i.get("Pedido") == numero]
    
    listas = carregar_listas_validacao()
    
    return render_template("editar_itens.html", pedido=numero, itens=its, listas=listas)


@app.route("/salvar_itens/<numero>", methods=["POST"])
@login_required 
def salvar_itens(numero):
    """Salva os itens editados, apagando e recriando os registros na aba 'Itens'."""
    usuario_log = current_user.id

    # 1. Busca itens antigos para log e para manter dados fixos
    antigos = ITENS_WS.get_all_records()
    antigos = [i for i in antigos if i.get("Pedido") == numero]
    
    if not antigos:
        flash("Não foi possível encontrar dados fixos para este pedido.", "erro")
        return redirect(url_for("painel"))

    dados_fixos = {
        "Solicitante": antigos[0].get("Solicitante", "N/A"),
        "Setor": antigos[0].get("Setor", "N/A"),
        "Data": antigos[0].get("Data", datetime.now().strftime("%d/%m/%Y %H:%M"))
    }

    try:
        # 2. Apaga todos os itens do pedido
        all_records = ITENS_WS.get_all_records()
        # Encontra a linha (1-base) para deletar, pulando o cabeçalho (start=2)
        rows_to_delete = [i for i, r in enumerate(all_records, start=2) if r.get("Pedido") == numero]
        
        for r in reversed(rows_to_delete):
            ITENS_WS.delete_rows(r)

        # 3. Grava novos itens
        total = int(request.form.get("total_itens", 0))
        novos = []
        for idx in range(total):
            material = request.form.get(f"material_{idx}")
            qtde_str = request.form.get(f"qtde_{idx}")
            obs_item = request.form.get(f"obs_item_{idx}", "")

            if material and qtde_str:
                try:
                    qtde = int(qtde_str)
                    novos.append([
                        numero,
                        idx + 1,
                        material.strip(),
                        qtde,
                        obs_item.strip(),
                        dados_fixos["Solicitante"],
                        dados_fixos["Setor"],
                        dados_fixos["Data"]
                    ])
                except ValueError:
                    flash(f"Quantidade inválida para o item {idx+1}. Item ignorado.", "aviso")
        
        for it in novos:
            append_atomico(ITENS_WS, it)

        # 4. Grava Log
        grava_log("EDIÇÃO-ITENS", numero, usuario_log, dados_anteriores=antigos)
        flash("Itens alterados.", "sucesso")
        
    except RuntimeError as e:
        flash(f"Falha ao salvar itens na planilha: {e}", "erro")
    except Exception as e:
        flash(f"Erro inesperado ao editar itens: {e}", "erro")

    return redirect(url_for("painel"))

# --- CANCELAR ---
@app.route("/cancelar/<numero>", methods=["GET", "POST"])
@login_required 
def cancelar_pedido(numero):
    """Rota para alterar o status de um pedido para CANCELADO."""
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
        
    try:
        reg = PED_WS.row_values(cel.row)
        reg_dict = dict(zip(PED_HEADERS, reg))
    except Exception:
        reg_dict = {"Pedido": numero, "Status": "N/A"}

    if request.method == "POST":
        motivo = request.form.get("motivo", "Sem motivo informado.").strip()
        usuario_log = current_user.id 
        
        try:
            grava_log("CANCELAMENTO", numero, usuario_log, motivo, dados_anteriores=reg_dict)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Status"), "CANCELADO")
            flash("Pedido cancelado.", "sucesso")
        except Exception as e:
            flash(f"Erro ao cancelar pedido: {e}", "erro")
            
        return redirect(url_for("painel"))
        
    return render_template("cancelar.html", pedido=numero)

# --- DAR BAIXA ---
@app.route("/baixa/<numero>", methods=["GET", "POST"])
@login_required 
def dar_baixa(numero):
    """Rota para dar baixa (mudar status para RETIRADO) em um pedido."""
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
        
    try:
        reg = PED_WS.row_values(cel.row)
        reg_dict = dict(zip(PED_HEADERS, reg))
    except Exception:
        reg_dict = {"Pedido": numero, "Status": "N/A"}

    if request.method == "POST":
        data_ret = datetime.now().strftime("%d/%m/%Y %H:%M")
        retirado_por = request.form["retirado_por"].strip()
        # O campo conferido_por tem um valor padrão mais seguro: o usuário logado
        conferido_por = request.form.get("conferido_por", current_user.id).strip() 

        if not retirado_por:
            flash("O campo 'Retirado por' é obrigatório.", "erro")
            return redirect(url_for("dar_baixa", numero=numero))
            
        try:
            # 1. Atualiza a aba "Pedidos"
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Status"), "RETIRADO")
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Retirado por"), retirado_por)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Retirado em"), data_ret)
            
            # 2. Grava na aba "Retiradas"
            append_atomico(RET_WS, [numero, data_ret, retirado_por, conferido_por])
            
            # 3. Grava Log
            grava_log("BAIXA/RETIRADA", numero, conferido_por, 
                      motivo=f"Retirado por: {retirado_por}", dados_anteriores=reg_dict)

            flash("Retirada confirmada.", "sucesso")
        except Exception as e:
            flash(f"Erro ao dar baixa no pedido: {e}", "erro")
            
        return redirect(url_for("painel"))
        
    return render_template("baixa.html", pedido=numero)

# --- ITENS (VISUALIZAÇÃO) ---
@app.route("/itens/<numero>")
@login_required 
def itens_do_pedido(numero):
    """Exibe a lista de itens de um pedido em modo de visualização."""
    its = ITENS_WS.get_all_records()
    its = [i for i in its if i.get("Pedido") == numero]
    return render_template("itens.html", pedido=numero, itens=its)

# --- Inicialização ---
if __name__ == "__main__":
    # Garante que o modo de depuração não esteja ativo em produção
    debug_mode = os.getenv("FLASK_ENV") != "production"
    app.run(debug=debug_mode, host='0.0.0.0', port=os.getenv("PORT", 5000))
