from flask import Flask, request, render_template_string, redirect, session, url_for
import sqlite3

app = Flask(__name__)
app.secret_key = "chave_secreta_super_segura"

# ==========================================
# Configuração do Banco de Dados
# ==========================================
def init_db():
    conn = sqlite3.connect('sistema_web.db')
    c = conn.cursor()
    # Tabela de Usuários
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
                 id INTEGER PRIMARY KEY, usuario TEXT UNIQUE, senha TEXT)''')
    # Tabela de Botões Personalizados
    c.execute('''CREATE TABLE IF NOT EXISTS botoes (
                 id INTEGER PRIMARY KEY, nome TEXT, cor TEXT, link TEXT)''')
    
    # Criar usuário admin padrão se não existir
    c.execute("SELECT * FROM usuarios WHERE usuario='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO usuarios (usuario, senha) VALUES ('admin', '1234')")
    
    conn.commit()
    conn.close()

init_db()

# ==========================================
# Telas em HTML (Front-end)
# ==========================================
HTML_LOGIN = '''
<!DOCTYPE html>
<html>
<head><title>Login - Robô</title><style>body {font-family: Arial; display: flex; justify-content: center; align-items: center; height: 100vh; background: #f0f2f5;} .card {background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.1); text-align: center;} input {width: 90%; padding: 10px; margin: 10px 0; border: 1px solid #ccc; border-radius: 4px;} button {background: #007bff; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; width: 100%;}</style></head>
<body>
    <div class="card">
        <h2>Entrar no Sistema</h2>
        <form method="POST">
            <input type="text" name="usuario" placeholder="Nome de Usuário" required>
            <input type="password" name="senha" placeholder="Senha" required>
            <button type="submit">Entrar</button>
        </form>
        <p style="color:red;">{{ erro }}</p>
    </div>
</body>
</html>
'''

HTML_DASHBOARD = '''
<!DOCTYPE html>
<html>
<head><title>Painel - Robô</title><style>body {font-family: Arial; background: #f0f2f5; margin: 0;} .header {background: #333; color: white; padding: 15px; display: flex; justify-content: space-between;} .container {padding: 20px;} .btn-custom {padding: 15px 25px; margin: 10px; border: none; border-radius: 5px; color: white; font-weight: bold; cursor: pointer; text-decoration: none; display: inline-block;} .card {background: white; padding: 20px; border-radius: 8px; margin-top: 20px;} input, select {padding: 8px; margin: 5px;}</style></head>
<body>
    <div class="header">
        <h2>🤖 Painel de Controle</h2>
        <div>Bem-vindo, <b>{{ session['usuario'] }}</b> | <a href="/logout" style="color: #ff4c4c;">Sair</a></div>
    </div>
    
    <div class="container">
        <h3>Seus Botões:</h3>
        {% for botao in botoes %}
            <a href="{{ botao[3] }}" class="btn-custom" style="background-color: {{ botao[2] }};">{{ botao[1] }}</a>
        {% endfor %}

        <div class="card">
            <h3>➕ Adicionar Novo Botão</h3>
            <form action="/add_botao" method="POST">
                <input type="text" name="nome" placeholder="Nome do Botão" required>
                <input type="text" name="link" placeholder="Link (ex: /pecas)" required>
                <select name="cor">
                    <option value="#4CAF50">Verde</option>
                    <option value="#2196F3">Azul</option>
                    <option value="#f44336">Vermelho</option>
                    <option value="#ff9800">Laranja</option>
                </select>
                <button type="submit" style="padding: 9px; background: #333; color: white; border: none; cursor:pointer;">Criar Botão</button>
            </form>
        </div>

        <div class="card">
            <h3>👤 Criar Novo Usuário</h3>
            <form action="/add_usuario" method="POST">
                <input type="text" name="novo_usuario" placeholder="Nome do Usuário" required>
                <input type="password" name="nova_senha" placeholder="Senha" required>
                <button type="submit" style="padding: 9px; background: #333; color: white; border: none; cursor:pointer;">Cadastrar Usuário</button>
            </form>
        </div>
    </div>
</body>
</html>
'''

# ==========================================
# Rotas do Site
# ==========================================
@app.route('/', methods=['GET', 'POST'])
def login():
    erro = ''
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']
        
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("SELECT * FROM usuarios WHERE usuario=? AND senha=?", (usuario, senha))
        user = c.fetchone()
        conn.close()
        
        if user:
            session['usuario'] = usuario
            return redirect(url_for('dashboard'))
        else:
            erro = 'Usuário ou senha incorretos!'
            
    return render_template_string(HTML_LOGIN, erro=erro)

@app.route('/dashboard')
def dashboard():
    if 'usuario' not in session:
        return redirect(url_for('login'))
        
    conn = sqlite3.connect('sistema_web.db')
    c = conn.cursor()
    c.execute("SELECT * FROM botoes")
    botoes = c.fetchall()
    conn.close()
    
    return render_template_string(HTML_DASHBOARD, botoes=botoes)

@app.route('/add_botao', methods=['POST'])
def add_botao():
    if 'usuario' in session:
        nome = request.form['nome']
        link = request.form['link']
        cor = request.form['cor']
        
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("INSERT INTO botoes (nome, cor, link) VALUES (?, ?, ?)", (nome, cor, link))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard'))

@app.route('/add_usuario', methods=['POST'])
def add_usuario():
    if 'usuario' in session:
        novo_usuario = request.form['novo_usuario']
        nova_senha = request.form['nova_senha']
        
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        try:
            c.execute("INSERT INTO usuarios (usuario, senha) VALUES (?, ?)", (novo_usuario, nova_senha))
            conn.commit()
        except sqlite3.IntegrityError:
            pass
        conn.close()
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, port=8080)