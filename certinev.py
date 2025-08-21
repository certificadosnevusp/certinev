from flask import Flask, render_template, request, redirect, url_for, session, jsonify, send_file, make_response, flash
import pandas as pd
import os
import qrcode
from PIL import Image
import base64
import shutil
from io import BytesIO
from functools import wraps
from datetime import datetime
from docxtpl import DocxTemplate
import zipfile
import tempfile
import json
from werkzeug.security import generate_password_hash, check_password_hash
import locale
import re
from docx2pdf import convert
import uuid
import logging
from dotenv import load_dotenv

# --- Configuração de Logging para melhor rastreabilidade de erros ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Carregar variáveis de ambiente ---
# A função load_dotenv() carrega as variáveis de um arquivo .env, se existir.
# Isso garante que segredos como a chave secreta do Flask não estejam hardcoded no código-fonte.
load_dotenv()

# Definir o locale para português para formatar a data corretamente
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    # Em alguns sistemas, 'pt_BR' pode ser o nome do locale
    locale.setlocale(locale.LC_TIME, 'pt_BR')

# --- Configuração inicial do Flask ---
app = Flask(__name__)
# A secret_key agora é lida de uma variável de ambiente para maior segurança.
# Se a variável não estiver definida, usa-se uma chave temporária para desenvolvimento.
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'certinev@123_temp_dev_key')
pd.options.mode.chained_assignment = None

# Define o diretório base do projeto para garantir que os caminhos estejam corretos
BASE_DIR = os.path.dirname(__file__)
USERS_FILE = os.path.join(BASE_DIR, 'users.json')
PLANILHAS_DIR = os.path.join(BASE_DIR, 'planilhas')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')

# --- Funções de Ajuda de Arquivos ---
def load_users():
    """Carrega os usuários do arquivo JSON de forma segura."""
    if os.path.exists(USERS_FILE):
        try:
            with open(USERS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logging.error(f"Erro ao carregar o arquivo de usuários: {e}")
            return {}
    return {}

def save_users(users):
    """Salva os usuários no arquivo JSON."""
    try:
        with open(USERS_FILE, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=4)
        return True
    except IOError as e:
        logging.error(f"Erro ao salvar o arquivo de usuários: {e}")
        return False

def read_excel_data(file_path, sheet_name):
    """Lê dados de uma planilha Excel de forma segura."""
    try:
        if os.path.exists(file_path):
            return pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            return None
    except (FileNotFoundError, ValueError) as e:
        logging.error(f"Erro ao ler o ficheiro Excel '{file_path}' (aba '{sheet_name}'): {e}")
        return None
    except Exception as e:
        logging.error(f"Erro inesperado ao ler o ficheiro Excel: {e}")
        return None

def save_excel_data(file_path, sheets):
    """Salva dados em um arquivo Excel com múltiplas abas."""
    try:
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, df in sheets.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        logging.error(f"Erro ao salvar o ficheiro Excel '{file_path}': {e}")
        return False

# Carregar os utilizadores ao iniciar a aplicação
USERS = load_users()

# --- Decoradores e Validações ---
def login_required(f):
    """Decorador para proteger rotas que exigem autenticação."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'username' not in session:
            flash('Por favor, faça login para aceder a esta página.', 'info')
            return redirect(url_for('login'))
        
        global USERS
        USERS = load_users()
        user = USERS.get(session.get('username'))

        if not user:
            session.pop('username', None)
            session.pop('user_type', None)
            return redirect(url_for('login'))
        
        # O usuário agora pode ser 'alfa' ou 'admin'
        if user.get('requires_password_change') and request.endpoint not in ['mudar_senha', 'logout']:
            flash('Por favor, mude sua senha para continuar.', 'warning')
            return redirect(url_for('mudar_senha'))
        
        return f(*args, **kwargs)
    return decorated_function

def alfa_required(f):
    """Decorador para proteger rotas exclusivas do usuário 'alfa'."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if session.get('user_type') != 'alfa':
            flash('Permissão negada.', 'danger')
            return redirect(url_for('cria_emite'))
        return f(*args, **kwargs)
    return decorated_function

def validate_cpf(cpf):
    """Função para validar o formato do CPF."""
    return re.match(r'^\d{11}$', cpf)

# --- Rotas da Aplicação: Autenticação ---
@app.route('/')
def index():
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        global USERS
        USERS = load_users()
        user = USERS.get(username)
        
        # Verifica se o usuário existe e se a senha corresponde ao hash armazenado
        if user and check_password_hash(user['password_hash'], password):
            session['username'] = username
            session['user_type'] = user['type']
            if user.get('requires_password_change'):
                return redirect(url_for('mudar_senha'))
            return redirect(url_for('cria_emite'))
        else:
            flash('Usuário ou senha incorretos!', 'danger')
            return render_template('index_login.html')
            
    return render_template('index_login.html')

@app.route('/logout')
@login_required
def logout():
    session.pop('username', None)
    session.pop('user_type', None)
    flash('Você saiu da sua conta com sucesso.', 'success')
    return redirect(url_for('login'))

@app.route('/mudar_senha', methods=['GET', 'POST'])
@login_required
def mudar_senha():
    global USERS
    USERS = load_users()
    user = USERS.get(session.get('username'))

    # Se o usuário não existe ou não precisa mudar a senha, redireciona
    if not user or not user.get('requires_password_change'):
        return redirect(url_for('cria_emite'))
    
    if request.method == 'POST':
        nova_senha = request.form.get('nova_senha')
        confirmar_senha = request.form.get('confirmar_senha')

        if not nova_senha or nova_senha != confirmar_senha:
            flash('As senhas não coincidem ou estão vazias.', 'danger')
            return render_template('mudar_senha.html')
        
        # Salva a nova senha como um hash seguro
        USERS[session.get('username')]['password_hash'] = generate_password_hash(nova_senha)
        USERS[session.get('username')]['requires_password_change'] = False
        save_users(USERS)
        
        flash('Senha alterada com sucesso!', 'success')
        return redirect(url_for('cria_emite'))

    return render_template('mudar_senha.html')

# --- Rotas da Aplicação: Gerenciamento de Usuários ---
@app.route('/gerenciar_usuarios')
@login_required
@alfa_required
def gerenciar_usuarios():
    global USERS
    USERS = load_users()
    return render_template('gerenciar_usuarios.html', users=USERS)

@app.route('/excluir_usuario/<string:username>')
@login_required
@alfa_required
def excluir_usuario(username):
    if username == 'alfa':
        flash('Erro: O usuário "alfa" não pode ser excluído.', 'danger')
        return redirect(url_for('gerenciar_usuarios'))
        
    global USERS
    USERS = load_users()
    if username in USERS:
        del USERS[username]
        save_users(USERS)
        flash(f'Usuário {username} excluído com sucesso.', 'success')
        return redirect(url_for('gerenciar_usuarios'))
    
    flash(f'Erro: Usuário {username} não encontrado.', 'danger')
    return redirect(url_for('gerenciar_usuarios'))

@app.route('/criar_usuario')
@login_required
@alfa_required
def criar_usuario():
    return render_template('criar_usuario.html')

@app.route('/processa_usuario', methods=['POST'])
@login_required
@alfa_required
def processa_usuario():
    username = request.form.get('username')
    user_type = request.form.get('user_type')
    # Senha temporária agora é hashada
    temp_password = 'mudar@123'
    temp_password_hash = generate_password_hash(temp_password)

    if not username or not user_type:
        flash('O nome de usuário e o tipo de usuário são obrigatórios.', 'danger')
        return redirect(url_for('criar_usuario'))
    
    global USERS
    USERS = load_users()
    if username in USERS:
        flash('Erro: Usuário já existe!', 'danger')
        return redirect(url_for('criar_usuario'))

    USERS[username] = {
        'password_hash': temp_password_hash,
        'type': user_type,
        'requires_password_change': True
    }
    
    if save_users(USERS):
        flash(f'Usuário "{username}" criado com sucesso. A senha temporária é "{temp_password}".', 'success')
        return redirect(url_for('gerenciar_usuarios'))
    else:
        flash('Erro ao criar usuário.', 'danger')
        return redirect(url_for('criar_usuario'))

# --- Rotas da Aplicação: Eventos e Certificados ---
@app.route('/cria_emite')
@login_required
def cria_emite():
    return render_template('cria_emite.html')

@app.route('/criar_evento')
@login_required
def criar_evento():
    return render_template('criar_evento.html', tipos_evento=["do curso", "da aula", "do seminário", "do treinamento", "do evento", "do webinário", "da roda de conversa"],
                           cargas_horarias={1: "1 (uma) hora", 2: "2 (duas) horas", 3: "3 (três) horas", 4: "4 (quatro) horas", 5: "5 (cinco) horas",
                                             6: "6 (seis) horas", 7: "7 (sete) horas", 8: "8 (oito) horas", 9: "9 (nove) horas", 10: "10 (dez) horas",
                                             11: "11 (onze) horas", 12: "12 (doze) horas"})

@app.route('/processa_evento', methods=['POST'])
@login_required
def processa_evento():
    tipo_evento = request.form.get('tipo_evento')
    titulo_descricao = request.form.get('titulo_descricao')
    horario_inicio = request.form.get('horario_inicio')
    horario_fim = request.form.get('horario_fim')
    modalidade_tipo = request.form.get('modalidade_tipo')
    data_evento_str = request.form.get('data_evento')
    carga_horaria_valor = request.form.get('carga_horaria', type=int)

    if not all([tipo_evento, titulo_descricao, horario_inicio, horario_fim, modalidade_tipo, data_evento_str, carga_horaria_valor]):
        flash('Todos os campos do formulário são obrigatórios.', 'danger')
        return redirect(url_for('criar_evento'))

    TIPO_EVENTO = ["do curso", "da aula", "do seminário", "do treinamento", "do evento", "do webinário", "da roda de conversa"]
    CARGA_HORARIA_OPCOES = {
        1: "1 (uma) hora", 2: "2 (duas) horas", 3: "3 (três) horas", 4: "4 (quatro) horas", 5: "5 (cinco) horas",
        6: "6 (seis) horas", 7: "7 (sete) horas", 8: "8 (oito) horas", 9: "9 (nove) horas", 10: "10 (dez) horas",
        11: "11 (onze) horas", 12: "12 (doze) horas"
    }

    titulo_completo = f"{tipo_evento.capitalize()} {titulo_descricao}"
    carga_horaria = CARGA_HORARIA_OPCOES.get(carga_horaria_valor)
    
    if not carga_horaria:
        flash('Carga horária inválida.', 'danger')
        return redirect(url_for('criar_evento'))

    try:
        data_evento_obj = datetime.strptime(data_evento_str, '%Y-%m-%d')
        data_evento_formatada = data_evento_obj.strftime('%d de %B de %Y')
    except ValueError:
        flash('Formato de data inválido.', 'danger')
        return redirect(url_for('criar_evento'))

    nome_arquivo = re.sub(r'[^a-zA-Z0-9_.]', '', titulo_completo.replace(' ', '_')) + '.xlsx'
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_arquivo)

    if not os.path.exists(PLANILHAS_DIR):
        os.makedirs(PLANILHAS_DIR)

    dados_evento_df = pd.DataFrame({
        'Campo': ['Título/Descrição', 'Data', 'Horário de Início', 'Horário de Fim', 'Modalidade', 'Carga Horária', 'Criado por'],
        'Valor': [titulo_completo, data_evento_formatada, horario_inicio, horario_fim, modalidade_tipo, carga_horaria, session.get('username')]
    })
    lista_presenca_df = pd.DataFrame(columns=['Nome Completo', 'CPF', 'Email', 'Newsletter Opt-in'])

    if save_excel_data(caminho_arquivo, {'Dados do Evento': dados_evento_df, 'Lista de Presença': lista_presenca_df}):
        flash(f'Evento "{titulo_completo}" criado com sucesso!', 'success')
        return redirect(url_for('confirmado', titulo_descricao=titulo_completo, modalidade_tipo=modalidade_tipo))
    else:
        flash('Ocorreu um erro ao criar o evento.', 'danger')
        return redirect(url_for('criar_evento'))

@app.route('/confirmado')
@login_required
def confirmado():
    titulo_descricao = request.args.get('titulo_descricao')
    modalidade_tipo = request.args.get('modalidade_tipo')

    if not titulo_descricao or not modalidade_tipo:
        flash('Erro: Detalhes do evento não foram fornecidos.', 'danger')
        return redirect(url_for('cria_emite'))
    
    qr_code_base64 = None
    link_presenca = None
    
    if modalidade_tipo in ['presencial', 'hibrida']:
        link_presenca = url_for('lista_presencial', titulo_descricao=titulo_descricao, _external=True)
        try:
            qr_img = qrcode.make(link_presenca)
            buf = BytesIO()
            qr_img.save(buf, format='PNG')
            qr_code_base64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        except Exception as e:
            logging.error(f"Erro ao gerar QR Code para evento '{titulo_descricao}': {e}")
    elif modalidade_tipo == 'online':
        link_presenca = url_for('lista_online', titulo_descricao=titulo_descricao, _external=True)

    return render_template('confirmado.html', link_presenca=link_presenca, qr_code_base64=qr_code_base64, titulo_descricao=titulo_descricao)

# --- Rotas da Aplicação: Lista de Presença ---
@app.route('/lista_presencial/<string:titulo_descricao>')
def lista_presencial(titulo_descricao):
    if not titulo_descricao:
        return "Erro: Parâmetros inválidos.", 400
    
    nome_arquivo = re.sub(r'[^a-zA-Z0-9_.]', '', titulo_descricao.replace(' ', '_')) + '.xlsx'
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_arquivo)
    
    if not os.path.exists(caminho_arquivo):
        return "Erro: Evento não encontrado.", 404
        
    return render_template('lista_presencial.html', titulo_descricao=titulo_descricao)

@app.route('/lista_online/<string:titulo_descricao>')
def lista_online(titulo_descricao):
    if not titulo_descricao:
        return "Erro: Parâmetros inválidos.", 400

    nome_arquivo = re.sub(r'[^a-zA-Z0-9_.]', '', titulo_descricao.replace(' ', '_')) + '.xlsx'
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_arquivo)

    if not os.path.exists(caminho_arquivo):
        return "Erro: Evento não encontrado.", 404

    return render_template('lista_presencial.html', titulo_descricao=titulo_descricao)

@app.route('/verifica_cpf', methods=['POST'])
def verifica_cpf():
    cpf = request.form.get('cpf', '').strip()
    titulo_descricao = request.form.get('titulo_descricao', '').strip()
    
    if not validate_cpf(cpf) or not titulo_descricao:
        return jsonify({'exists': False, 'message': 'CPF ou evento inválido.'}), 400

    nome_arquivo = re.sub(r'[^a-zA-Z0-9_.]', '', titulo_descricao.replace(' ', '_')) + '.xlsx'
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_arquivo)

    if not os.path.exists(caminho_arquivo):
        return jsonify({'exists': False, 'message': 'Evento não encontrado.'}), 404

    df = read_excel_data(caminho_arquivo, 'Lista de Presença')
    
    if df is not None and not df[df['CPF'].astype(str).str.strip() == cpf].empty:
        return jsonify({'exists': True})
        
    return jsonify({'exists': False})

@app.route('/registra_presenca', methods=['POST'])
def registra_presenca():
    titulo_descricao = request.form.get('titulo_descricao')
    nome_completo = request.form.get('nome_completo')
    cpf = request.form.get('cpf')
    email = request.form.get('email')
    newsletter_opt_in = request.form.get('newsletter_opt_in') == 'on'

    if not all([titulo_descricao, nome_completo, cpf, email]):
        flash('Por favor, preencha todos os campos.', 'danger')
        return redirect(url_for('lista_presencial', titulo_descricao=titulo_descricao))
    
    nome_arquivo = re.sub(r'[^a-zA-Z0-9_.]', '', titulo_descricao.replace(' ', '_')) + '.xlsx'
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_arquivo)

    try:
        df_existente = read_excel_data(caminho_arquivo, 'Lista de Presença')
        if df_existente is None:
            flash('Erro: Arquivo do evento não encontrado.', 'danger')
            return redirect(url_for('lista_presencial', titulo_descricao=titulo_descricao))

        if not df_existente[df_existente['CPF'].astype(str).str.strip() == cpf].empty:
            flash('Presença já registrada para este CPF!', 'info')
            return redirect(url_for('lista_presencial', titulo_descricao=titulo_descricao))
        
        nova_linha_df = pd.DataFrame([{
            'Nome Completo': nome_completo,
            'CPF': cpf,
            'Email': email,
            'Newsletter Opt-in': 'Sim' if newsletter_opt_in else 'Não'
        }])
        
        df_atualizado = pd.concat([df_existente, nova_linha_df], ignore_index=True)
        df_dados_evento = read_excel_data(caminho_arquivo, 'Dados do Evento')
        
        if save_excel_data(caminho_arquivo, {'Dados do Evento': df_dados_evento, 'Lista de Presença': df_atualizado}):
            return redirect(url_for('presenca_registrada'))
        else:
            flash('Ocorreu um erro ao salvar os dados.', 'danger')
            return redirect(url_for('lista_presencial', titulo_descricao=titulo_descricao))
    
    except Exception as e:
        flash(f'Ocorreu um erro ao registrar a presença: {e}', 'danger')
        return redirect(url_for('lista_presencial', titulo_descricao=titulo_descricao))

@app.route('/presenca_registrada')
def presenca_registrada():
    return render_template('presenca_registrada.html')

# --- Rotas da Aplicação: Visualização e Download ---
@app.route('/visualizar_listas')
@login_required
def visualizar_listas():
    eventos = []
    if os.path.exists(PLANILHAS_DIR):
        eventos = [f for f in os.listdir(PLANILHAS_DIR) if f.endswith('.xlsx')]
        eventos = sorted([os.path.splitext(f)[0] for f in eventos])
    return render_template('visualizar_listas.html', eventos=eventos)

@app.route('/visualizar_evento/<string:nome_evento>')
@login_required
def visualizar_evento(nome_evento):
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_evento + '.xlsx')
    
    if not os.path.exists(caminho_arquivo):
        flash(f'Erro: Evento "{nome_evento}" não encontrado.', 'danger')
        return redirect(url_for('visualizar_listas'))
    
    df_dados_evento = read_excel_data(caminho_arquivo, 'Dados do Evento')
    df_presenca = read_excel_data(caminho_arquivo, 'Lista de Presença')
    
    if df_dados_evento is None or df_presenca is None:
        flash('Erro ao ler os dados do evento.', 'danger')
        return redirect(url_for('visualizar_listas'))

    evento_dict = df_dados_evento.set_index('Campo')['Valor'].to_dict()
    lista_presenca = df_presenca.to_dict('records')
    modalidade = evento_dict.get('Modalidade', '').lower()
    titulo_descricao = evento_dict.get('Título/Descrição', '')
    link_presenca = None
    qr_code_base64 = None

    if modalidade in ['presencial', 'hibrida']:
        link_presenca = url_for('lista_presencial', titulo_descricao=titulo_descricao, _external=True)
        try:
            qr_img = qrcode.make(link_presenca)
            buf = BytesIO()
            qr_img.save(buf, format='PNG')
            qr_code_base64 = base64.b64encode(buf.getvalue()).decode('utf-8')
        except Exception as e:
            logging.error(f"Erro ao gerar QR Code na visualização do evento '{nome_evento}': {e}")
    elif modalidade == 'online':
        link_presenca = url_for('lista_online', titulo_descricao=titulo_descricao, _external=True)

    return render_template('visualizar_evento.html',
                           nome_evento=nome_evento,
                           evento=evento_dict,
                           lista_presenca=lista_presenca,
                           link_presenca=link_presenca,
                           qr_code_base64=qr_code_base64
                          )

@app.route('/download_lista/<string:nome_evento>')
@login_required
def download_lista(nome_evento):
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_evento + '.xlsx')
    
    if not os.path.exists(caminho_arquivo):
        flash(f'Erro: Evento "{nome_evento}" não encontrado.', 'danger')
        return redirect(url_for('visualizar_listas'))
    
    return send_file(caminho_arquivo, as_attachment=True, download_name=nome_evento + '.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# --- Rotas da Aplicação: Geração de Certificados ---
@app.route('/gerar_certificados')
@login_required
def gerar_certificados():
    eventos = []
    if os.path.exists(PLANILHAS_DIR):
        eventos = [f for f in os.listdir(PLANILHAS_DIR) if f.endswith('.xlsx')]
        eventos = sorted([os.path.splitext(f)[0] for f in eventos])
    return render_template('gerar_certificados.html', eventos=eventos)

@app.route('/lista_participantes/<string:nome_evento>')
@login_required
def lista_participantes(nome_evento):
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_evento + '.xlsx')
    
    if not os.path.exists(caminho_arquivo):
        flash(f'Erro: Evento "{nome_evento}" não encontrado.', 'danger')
        return redirect(url_for('gerar_certificados'))

    df_presenca = read_excel_data(caminho_arquivo, 'Lista de Presença')
    
    if df_presenca is None:
        flash('Erro ao ler os dados do evento.', 'danger')
        return redirect(url_for('gerar_certificados'))

    participantes = df_presenca.to_dict('records')
    
    return render_template('lista_participantes.html',
                           nome_evento=nome_evento,
                           evento_nome_display=nome_evento.replace('_', ' ').replace('-', ' '),
                           participantes=participantes
                          )

@app.route('/download_certificado/<nome_evento>/<cpf>', methods=['GET'])
@login_required
def download_certificado(nome_evento, cpf):
    temp_dir_path = None
    try:
        planilha_path = os.path.join(PLANILHAS_DIR, f"{nome_evento}.xlsx")
        if not os.path.exists(planilha_path):
            flash('Planilha do evento não encontrada.', 'danger')
            return redirect(url_for('cria_emite'))

        df_presenca = read_excel_data(planilha_path, sheet_name='Lista de Presença')
        if df_presenca is None:
            flash('Erro ao ler a lista de presença.', 'danger')
            return redirect(url_for('lista_participantes', nome_evento=nome_evento))

        participante = df_presenca[df_presenca['CPF'].astype(str) == str(cpf)]
        if participante.empty:
            flash('Participante não encontrado na lista de presença.', 'danger')
            return redirect(url_for('lista_participantes', nome_evento=nome_evento))

        df_evento = read_excel_data(planilha_path, sheet_name='Dados do Evento')
        if df_evento is None:
            flash('Erro ao ler os dados do evento.', 'danger')
            return redirect(url_for('lista_participantes', nome_evento=nome_evento))

        evento_dict = df_evento.set_index('Campo')['Valor'].to_dict()
        context = {
            'nome': participante['Nome Completo'].values[0],
            'documento_tipo': 'CPF',
            'documento_numero': str(cpf),
            'titulo_descricao': evento_dict.get('Título/Descrição', ''),
            'data_evento': evento_dict.get('Data', ''),
            'carga_horaria': evento_dict.get('Carga Horária', '')
        }

        caminho_template = os.path.join(TEMPLATES_DIR, 'Certificado_Base.docx')
        if not os.path.exists(caminho_template):
            flash('Template de certificado não encontrado.', 'danger')
            return redirect(url_for('lista_participantes', nome_evento=nome_evento))

        tpl = DocxTemplate(caminho_template)
        tpl.render(context)
        
        temp_dir_path = tempfile.mkdtemp()
        nome_docx = f"Certificado_{re.sub(r'[^a-zA-Z0-9_]', '', participante['Nome Completo'].values[0].replace(' ', '_'))}_{nome_evento}.docx"
        caminho_certificado_docx = os.path.join(temp_dir_path, nome_docx)
        tpl.save(caminho_certificado_docx)
        
        caminho_certificado_pdf = caminho_certificado_docx.replace('.docx', '.pdf')
        nome_certificado_download = nome_docx.replace('.docx', '.pdf')
        
        # Oferecendo alternativas para a conversão de DOCX para PDF
        try:
            convert(caminho_certificado_docx, caminho_certificado_pdf)
            response = send_file(caminho_certificado_pdf, as_attachment=True, download_name=nome_certificado_download)
        except Exception as e:
            logging.warning(f"Erro ao converter o certificado para PDF ({e}). Retornando o arquivo DOCX.")
            flash("Erro ao converter para PDF. Baixando o arquivo DOCX. Por favor, converta-o manualmente.", 'warning')
            response = send_file(caminho_certificado_docx, as_attachment=True, download_name=nome_docx)
        
        # Cria uma função de callback para limpar os arquivos temporários após o envio
        def cleanup_files():
            try:
                os.remove(caminho_certificado_docx)
                if os.path.exists(caminho_certificado_pdf):
                    os.remove(caminho_certificado_pdf)
                shutil.rmtree(temp_dir_path)
            except Exception as cleanup_e:
                logging.error(f"Erro ao limpar arquivos temporários: {cleanup_e}")

        response.call_on_close(cleanup_files)
        return response

    except Exception as e:
        logging.error(f"Ocorreu um erro inesperado na geração de certificado: {e}")
        flash(f"Ocorreu um erro inesperado: {e}", 'danger')
        if temp_dir_path and os.path.exists(temp_dir_path):
            shutil.rmtree(temp_dir_path)
        return redirect(url_for('lista_participantes', nome_evento=nome_evento))

@app.route('/download_certificados_zip/<string:nome_evento>')
@login_required
def download_certificados_zip(nome_evento):
    caminho_arquivo = os.path.join(PLANILHAS_DIR, nome_evento + '.xlsx')
    if not os.path.exists(caminho_arquivo):
        flash(f'Erro: Evento "{nome_evento}" não encontrado.', 'danger')
        return redirect(url_for('gerar_certificados'))
        
    df_dados_evento = read_excel_data(caminho_arquivo, 'Dados do Evento')
    df_presenca = read_excel_data(caminho_arquivo, 'Lista de Presença')

    if df_dados_evento is None or df_presenca is None:
        flash('Erro ao ler os dados do evento.', 'danger')
        return redirect(url_for('gerar_certificados'))

    evento = df_dados_evento.set_index('Campo')['Valor'].to_dict()
    participantes = df_presenca.to_dict('records')

    if not participantes:
        flash('Nenhum participante encontrado para este evento.', 'info')
        return redirect(url_for('lista_participantes', nome_evento=nome_evento))

    caminho_template = os.path.join(TEMPLATES_DIR, 'Certificado_Base.docx')
    if not os.path.exists(caminho_template):
        flash('Erro: Template de certificado não encontrado.', 'danger')
        return redirect(url_for('gerar_certificados'))

    zip_buffer = BytesIO()
    temp_dir_path = tempfile.mkdtemp()
    
    try:
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for participante in participantes:
                context = {
                    'nome': participante.get('Nome Completo', ''),
                    'documento_tipo': 'CPF',
                    'documento_numero': participante.get('CPF', ''),
                    'titulo_descricao': evento.get('Título/Descrição', ''),
                    'data_evento': evento.get('Data', ''),
                    'carga_horaria': evento.get('Carga Horária', '')
                }
                
                doc = DocxTemplate(caminho_template)
                doc.render(context)
                
                unique_id = uuid.uuid4()
                nome_docx_temp = f"certificado_{unique_id}.docx"
                caminho_docx_temp = os.path.join(temp_dir_path, nome_docx_temp)
                doc.save(caminho_docx_temp)
                
                caminho_pdf_temp = caminho_docx_temp.replace('.docx', '.pdf')
                
                # Tentativa de conversão
                try:
                    convert(caminho_docx_temp, caminho_pdf_temp)
                    nome_pdf_final = f"certificado_{re.sub(r'[^a-zA-Z0-9_]', '', participante.get('CPF', 'sem_cpf'))}.pdf"
                    zip_file.write(caminho_pdf_temp, nome_pdf_final)
                except Exception as e:
                    logging.warning(f"Erro ao converter certificado de {participante.get('Nome Completo', '')} para PDF: {e}. Adicionando o DOCX no zip.")
                    nome_docx_final = f"certificado_{re.sub(r'[^a-zA-Z0-9_]', '', participante.get('CPF', 'sem_cpf'))}.docx"
                    zip_file.write(caminho_docx_temp, nome_docx_final)
                    continue

    finally:
        # A nova forma de limpar o diretório, garantindo que tudo seja removido
        shutil.rmtree(temp_dir_path)

    zip_buffer.seek(0)
    response = make_response(zip_buffer.getvalue())
    response.headers['Content-Type'] = 'application/zip'
    response.headers['Content-Disposition'] = f'attachment; filename=certificados_{nome_evento}.zip'
    return response

# --- Inicialização da Aplicação ---
if __name__ == '__main__':
    if not os.path.exists(PLANILHAS_DIR):
        os.makedirs(PLANILHAS_DIR)
    
    # Criar um usuário 'alfa' inicial se o arquivo não existir
    if not os.path.exists(USERS_FILE):
        logging.info("Arquivo 'users.json' não encontrado. Criando usuário 'alfa' inicial.")
        initial_users = {
            "alfa": {
                # A senha temporária padrão agora é hashada
                "password_hash": generate_password_hash("certinev@123"),
                "type": "alfa",
                "requires_password_change": True  # Força a alteração da senha no primeiro login
            }
        }
        save_users(initial_users)
        
    app.run(debug=True)