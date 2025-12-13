from flask import Flask, render_template, request, redirect, url_for, flash
# =================================================================
# 🔑 NOVAS IMPORTAÇÕES DE SEGURANÇA E LOGIN
# =================================================================
import os
from dotenv import load_dotenv
from flask_login import LoginManager, UserMixin, login_user, logout_user, current_user, login_required
from werkzeug.security import generate_password_hash, check_password_hash

# ---------- MONGO ----------
from mongo_users import autenticar, usuario_por_id

import gspread
import json
from google.oauth2.service_account import Credentials
from datetime import datetime
from gspread.exceptions import WorksheetNotFound, APIError

# =================================================================
# 🚀 INICIALIZAÇÃO E CONFIGURAÇÃO
# =================================================================
load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET", "um-segredo-muito-forte-e-aleatorio-para-dev")

# =================================================================
# 🔑 CONFIGURAÇÃO DE LOGIN  (MongoDB)
# =================================================================
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message_category = 'info'


class User(UserMixin):
    def __init__(self, user_id):
        self.id = user_id
        doc = usuario_por_id(user_id)
        self.name = doc.get("name", user_id) if doc else "Usuário Desconhecido"


@login_manager.user_loader
def load_user(user_id):
    if usuario_por_id(user_id):
        return User(user_id)
    return None


# =================================================================
# ⚙️ CONFIGURAÇÃO DO GOOGLE SHEETS API
# =================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

google_json = os.getenv("GOOGLE_CREDENTIALS")
if not google_json:
    raise RuntimeError("ERRO: GOOGLE_CREDENTIALS não encontrada no .env")

try:
    INFO = json.loads(google_json)
    CREDS = Credentials.from_service_account_info(INFO, scopes=SCOPES)
    CLIENT = gspread.authorize(CREDS)
    SHEET = CLIENT.open("Controle de Expedição")
except Exception as e:
    print(f"ERRO ao inicializar Google Sheets: {e}")
    exit(1)

PED_WS = SHEET.worksheet("Pedidos")
RET_WS = SHEET.worksheet("Retiradas")

try:
    PED_HEADERS = PED_WS.row_values(1)
except Exception:
    PED_HEADERS = ["Pedido", "Solicitante", "Setor", "Status", "Data Pedido",
                   "Retirado por", "Retirado em", "Obs-geral"]


def inicializa_abas_auxiliares():
    global ITENS_WS, LOG_WS, VALIDACAO_WS, ITENS_HEADERS, LOG_HEADERS, VALIDACAO_HEADERS
    try:
        ITENS_WS = SHEET.worksheet("Itens")
        ITENS_HEADERS = ITENS_WS.row_values(1)
    except WorksheetNotFound:
        ITENS_WS = SHEET.add_worksheet(title="Itens", rows=1000, cols=10)
        ITENS_HEADERS = ["Pedido", "Item", "Material", "Qtde",
                         "Obs-item", "Solicitante", "Setor", "Data"]
        ITENS_WS.append_row(ITENS_HEADERS)

    try:
        LOG_WS = SHEET.worksheet("Histórico")
        LOG_HEADERS = LOG_WS.row_values(1)
    except WorksheetNotFound:
        LOG_WS = SHEET.add_worksheet(title="Histórico", rows=1000, cols=10)
        LOG_HEADERS = ["Data/Hora", "Pedido", "Tipo", "Usuário",
                       "Motivo", "Dados Anteriores"]
        LOG_WS.append_row(LOG_HEADERS)

    try:
        VALIDACAO_WS = SHEET.worksheet("Validacao")
        VALIDACAO_HEADERS = VALIDACAO_WS.row_values(1)
    except WorksheetNotFound:
        VALIDACAO_WS = SHEET.add_worksheet(title="Validacao", rows=1000, cols=5)
        VALIDACAO_HEADERS = ["Material", "Setor", "Solicitante",
                             "Unidade", "Obs-Geral-Padrao"]
        VALIDACAO_WS.append_row(VALIDACAO_HEADERS)

inicializa_abas_auxiliares()


def get_col_index(headers, col_name):
    try:
        return headers.index(col_name) + 1
    except ValueError:
        raise ValueError(f"Coluna '{col_name}' não encontrada nos cabeçalhos: {headers}")


def carregar_listas_validacao():
    try:
        dados = VALIDACAO_WS.get_all_records()
        materiais = sorted({d.get("Material") for d in dados if d.get("Material")})
        setores = sorted({d.get("Setor") for d in dados if d.get("Setor")})
        return {"setores": setores, "materiais": materiais}
    except Exception as e:
        print("Erro listas validação:", e)
        return {"setores": [], "materiais": []}


def proximo_id():
    try:
        ids = PED_WS.col_values(get_col_index(PED_HEADERS, "Pedido"))[1:]
    except Exception:
        ids = []
    nums = [int(i.split("-")[-1]) for i in ids if i and i.startswith("EXP-")]
    ano = datetime.now().year
    return f"EXP-{ano}-{str(max(nums, default=0) + 1).zfill(6)}"


def append_atomico(ws, linha):
    for _ in range(3):
        try:
            ws.append_row(linha)
            return
        except APIError as e:
            print("APIError, retry...", e)
            continue
    raise RuntimeError("Falha ao gravar após 3 tentativas")


def append_itens(num_pedido, itens, solicitante, setor, data_pedido):
    for idx, it in enumerate(itens, start=1):
        try:
            qtde = int(it["qtde"])
        except (ValueError, TypeError):
            qtde = 0
        append_atomico(ITENS_WS, [
            num_pedido, idx,
            it["material"].strip() if it.get("material") else "N/A",
            qtde,
            it["obs"].strip() if it.get("obs") else "",
            solicitante, setor, data_pedido
        ])


def grava_log(tipo, pedido, usuario, motivo="", dados_anteriores=None):
    append_atomico(LOG_WS, [
        datetime.now().strftime("%d/%m/%Y %H:%M"),
        pedido, tipo, usuario, motivo,
        str(dados_anteriores) if dados_anteriores else ""
    ])


# =================================================================
# 🔑 ROTAS DE AUTENTICAÇÃO  (MongoDB)
# =================================================================
@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        return redirect(url_for('painel'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password')
        user_doc = autenticar(username, password)
        if user_doc:
            login_user(User(username))
            flash(f'Bem-vindo(a), {user_doc.get("name", username)}!', 'success')
            next_page = request.args.get('next')
            return redirect(next_page or url_for('painel'))
        flash('Usuário ou senha incorretos.', 'danger')
    return render_template('login.html', title='Login')


@app.route('/logout')
@login_required
def logout():
    logout_user()
    return redirect(url_for('login'))


# =================================================================
# 🗺️ ROTAS PROTEGIDAS (com ordenação e setores únicos)
# =================================================================
@app.route("/")
@login_required
def painel():
    pedidos = PED_WS.get_all_records()

    # 1. Converte a data para objeto datetime (campo "Data Pedido")
    for p in pedidos:
        p["DataPedidoObj"] = datetime.strptime(
            p.get("Data Pedido", "01/01/2000 00:00"), "%d/%m/%Y %H:%M"
        )

    # 2. Ordena: mais recente no topo
    pedidos = sorted(pedidos, key=lambda x: x["DataPedidoObj"], reverse=True)

    # 3. Remove a chave auxiliar
    for p in pedidos:
        p.pop("DataPedidoObj", None)

    # 4. Dados auxiliares
    for p in pedidos:
        p["RetiradoPor"] = p.get("Retirado por", "-")
        p["RetiradoEm"] = p.get("Retirado em", "-")

    # 5. Lista de setores únicos para o filtro
    setores_unicos = sorted({p.get("Setor") for p in pedidos if p.get("Setor")})

    return render_template("painel.html", pedidos=pedidos, setores_unicos=setores_unicos)


@app.route("/novo", methods=["GET", "POST"])
@login_required
def criar_pedido():
    listas = carregar_listas_validacao()
    if request.method == "POST":
        try:
            solicitante = request.form["solicitante"].strip()
            setor = request.form["setor"].strip()
            total_itens = int(request.form.get("total_itens", 0))
            if not solicitante or not setor:
                flash("Solicitante e Setor são obrigatórios.", "erro")
                return render_template("criar_pedido.html", listas=listas)
            data_sep = datetime.now().strftime("%d/%m/%Y %H:%M")
            num_pedido = proximo_id()
            pedido_data = [
                num_pedido, solicitante, setor, "EM EXPEDIÇÃO",
                data_sep, "", "", request.form.get("obs_geral", "").strip()
            ]
            if len(pedido_data) > len(PED_HEADERS):
                pedido_data = pedido_data[:len(PED_HEADERS)]
            elif len(pedido_data) < len(PED_HEADERS):
                pedido_data.extend([""] * (len(PED_HEADERS) - len(pedido_data)))
            append_atomico(PED_WS, pedido_data)

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
                        flash(f"Qtde inválida item {i+1}", "aviso")
            if not itens:
                flash("Nenhum item válido adicionado.", "aviso")
            append_itens(num_pedido, itens, solicitante, setor, data_sep)
            grava_log("CRIAÇÃO", num_pedido, current_user.id)
            flash("Pedido + itens criados!", "sucesso")
            return redirect(url_for("painel"))
        except RuntimeError as e:
            flash(f"Erro ao gravar: {e}", "erro")
            return render_template("criar_pedido.html", listas=listas)
        except Exception as e:
            flash(f"Erro inesperado: {e}", "erro")
            return render_template("criar_pedido.html", listas=listas)
    return render_template("criar_pedido.html", listas=listas)


@app.route("/editar/<numero>", methods=["GET", "POST"])
@login_required
def editar_pedido(numero):
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
    reg = PED_WS.row_values(cel.row)
    reg_dict = dict(zip(PED_HEADERS, reg))
    listas = carregar_listas_validacao()
    itens_data = ITENS_WS.get_all_records()
    itens = [i for i in itens_data if i.get("Pedido") == numero]
    if request.method == "POST":
        try:
            novo_solicitante = request.form["solicitante"].strip()
            novo_setor = request.form["setor"].strip()
            novo_obs = request.form.get("obs_geral", "").strip()
            grava_log("EDIÇÃO-CABEÇALHO", numero, current_user.id, dados_anteriores=reg_dict)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Solicitante"), novo_solicitante)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Setor"), novo_setor)
            if "Obs-geral" in PED_HEADERS:
                PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Obs-geral"), novo_obs)
            flash("Cabeçalho alterado.", "sucesso")
            return redirect(url_for("painel"))
        except Exception as e:
            flash(f"Erro ao salvar: {e}", "erro")
            return render_template("editar.html", pedido=numero, info=reg_dict, itens=itens, listas=listas)
    return render_template("editar.html", pedido=numero, info=reg_dict, itens=itens, listas=listas)


@app.route("/editar_itens/<numero>")
@login_required
def editar_itens(numero):
    its = ITENS_WS.get_all_records()
    its = [i for i in its if i.get("Pedido") == numero]
    listas = carregar_listas_validacao()
    return render_template("editar_itens.html", pedido=numero, itens=its, listas=listas)


@app.route("/salvar_itens/<numero>", methods=["POST"])
@login_required
def salvar_itens(numero):
    antigos = ITENS_WS.get_all_records()
    antigos = [i for i in antigos if i.get("Pedido") == numero]
    if not antigos:
        flash("Nenhum item encontrado.", "erro")
        return redirect(url_for("painel"))
    fixos = {
        "Solicitante": antigos[0].get("Solicitante", "N/A"),
        "Setor": antigos[0].get("Setor", "N/A"),
        "Data": antigos[0].get("Data", datetime.now().strftime("%d/%m/%Y %H:%M"))
    }
    try:
        all_records = ITENS_WS.get_all_records()
        rows_del = [i for i, r in enumerate(all_records, start=2) if r.get("Pedido") == numero]
        for r in reversed(rows_del):
            ITENS_WS.delete_rows(r)
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
                        numero, idx + 1, material.strip(), qtde, obs_item.strip(),
                        fixos["Solicitante"], fixos["Setor"], fixos["Data"]
                    ])
                except ValueError:
                    flash(f"Qtde inválida item {idx+1}", "aviso")
        for it in novos:
            append_atomico(ITENS_WS, it)
        grava_log("EDIÇÃO-ITENS", numero, current_user.id, dados_anteriores=antigos)
        flash("Itens alterados.", "sucesso")
    except RuntimeError as e:
        flash(f"Falha ao salvar itens: {e}", "erro")
    except Exception as e:
        flash(f"Erro inesperado: {e}", "erro")
    return redirect(url_for("painel"))


@app.route("/cancelar/<numero>", methods=["GET", "POST"])
@login_required
def cancelar_pedido(numero):
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
    reg = PED_WS.row_values(cel.row)
    reg_dict = dict(zip(PED_HEADERS, reg))
    if request.method == "POST":
        motivo = request.form.get("motivo", "Sem motivo informado.").strip()
        try:
            grava_log("CANCELAMENTO", numero, current_user.id, motivo, dados_anteriores=reg_dict)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Status"), "CANCELADO")
            flash("Pedido cancelado.", "sucesso")
        except Exception as e:
            flash(f"Erro ao cancelar: {e}", "erro")
        return redirect(url_for("painel"))
    return render_template("cancelar.html", pedido=numero)


@app.route("/baixa/<numero>", methods=["GET", "POST"])
@login_required
def dar_baixa(numero):
    try:
        cel = PED_WS.find(numero, in_column=get_col_index(PED_HEADERS, "Pedido"))
    except gspread.CellNotFound:
        flash("Pedido não localizado.", "erro")
        return redirect(url_for("painel"))
    reg = PED_WS.row_values(cel.row)
    reg_dict = dict(zip(PED_HEADERS, reg))
    if request.method == "POST":
        data_ret = datetime.now().strftime("%d/%m/%Y %H:%M")
        retirado_por = request.form["retirado_por"].strip()
        conferido_por = request.form.get("conferido_por", current_user.id).strip()
        if not retirado_por:
            flash("Campo 'Retirado por' é obrigatório.", "erro")
            return redirect(url_for("dar_baixa", numero=numero))
        try:
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Status"), "RETIRADO")
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Retirado por"), retirado_por)
            PED_WS.update_cell(cel.row, get_col_index(PED_HEADERS, "Retirado em"), data_ret)
            append_atomico(RET_WS, [numero, data_ret, retirado_por, conferido_por])
            grava_log("BAIXA/RETIRADA", numero, conferido_por,
                      motivo=f"Retirado por: {retirado_por}", dados_anteriores=reg_dict)
            flash("Retirada confirmada.", "sucesso")
        except Exception as e:
            flash(f"Erro ao dar baixa: {e}", "erro")
        return redirect(url_for("painel"))
    return render_template("baixa.html", pedido=numero)


@app.route("/itens/<numero>")
@login_required
def itens_do_pedido(numero):
    its = ITENS_WS.get_all_records()
    its = [i for i in its if i.get("Pedido") == numero]
    return render_template("itens.html", pedido=numero, itens=its)


# =================================================================
# 🔥 INICIALIZAÇÃO
# =================================================================
if __name__ == "__main__":
    debug_mode = os.getenv("FLASK_ENV") != "production"
    app.run(debug=debug_mode, host='0.0.0.0', port=int(os.getenv("PORT", 5000)))