from flask import Flask, render_template, render_template_string, request, redirect, session, url_for, jsonify
import sqlite3
import json
import urllib.request
import urllib.error
import ssl

app = Flask(__name__)
app.secret_key = "chave_secreta_super_segura"

def init_db():
    conn = sqlite3.connect('sistema_web.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (id INTEGER PRIMARY KEY, usuario TEXT UNIQUE, senha TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS pecas (id INTEGER PRIMARY KEY, nome TEXT, estoque INTEGER, min_estoque INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS guepar_uso (id INTEGER PRIMARY KEY, nome TEXT, estoque INTEGER, min_estoque INTEGER)''')
    c.execute('''CREATE TABLE IF NOT EXISTS fornecedores (id INTEGER PRIMARY KEY, nome TEXT, contato TEXT, telefone TEXT)''')
    try:
        c.execute("ALTER TABLE fornecedores ADD COLUMN cnpj TEXT")
    except:
        pass 
    
    c.execute("SELECT * FROM usuarios WHERE usuario='admin'")
    if not c.fetchone():
        c.execute("INSERT INTO usuarios (usuario, senha) VALUES ('admin', '1234')")
    conn.commit()
    conn.close()

init_db()

@app.route('/', methods=['GET', 'POST'])
def login():
    erro = ''
    if 'usuario' in session: return redirect(url_for('dashboard'))
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
            erro = 'Usuário/senha incorretos!'
    HTML_LOGIN = '''<!DOCTYPE html><html><head><title>Login - Guepar vs RobÔ</title><style>body {background: #0f172a; color: white; font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0;} .card {background: #1e293b; padding: 40px; border-radius: 8px; text-align: center; border: 1px solid #38bdf8; box-shadow: 0 4px 15px rgba(0,0,0,0.5);} input {width: 100%; padding: 12px; margin: 10px 0; border-radius: 5px; border: none; background: #334155; color: white; box-sizing: border-box;} button {background: #0ea5e9; color: white; font-weight: bold; border: none; padding: 12px; border-radius: 5px; cursor: pointer; width: 100%; margin-top: 10px;} button:hover {background: #0284c7;}</style></head><body><div class="card"><h2 style="margin-top: 0; color: #38bdf8;">Guepar vs RobÔ</h2><p>Acesso Restrito</p><form method="POST"><input type="text" name="usuario" placeholder="Usuário" required><input type="password" name="senha" placeholder="Senha" required><button type="submit">ENTRAR</button></form><p style="color:#ef4444;">{{ erro }}</p></div></body></html>'''
    return render_template_string(HTML_LOGIN, erro=erro)

@app.route('/dashboard')
def dashboard():
    if 'usuario' not in session: return redirect(url_for('login'))
    conn = sqlite3.connect('sistema_web.db')
    c = conn.cursor()
    c.execute("SELECT * FROM pecas")
    pecas = c.fetchall()
    c.execute("SELECT * FROM guepar_uso")
    itens_guepar = c.fetchall()
    c.execute("SELECT * FROM fornecedores")
    fornecedores = c.fetchall()
    conn.close()
    
    total_pecas = len(pecas)
    pecas_criticas = sum(1 for p in pecas if p[2] <= p[3])
    total_guepar = len(itens_guepar)
    guepar_criticos = sum(1 for g in itens_guepar if g[2] <= g[3])
    total_fornecedores = len(fornecedores)
    
    return render_template('index.html', pecas=pecas, itens_guepar=itens_guepar, fornecedores=fornecedores, total_pecas=total_pecas, pecas_criticas=pecas_criticas, total_guepar=total_guepar, guepar_criticos=guepar_criticos, total_fornecedores=total_fornecedores)

@app.route('/api/chat', methods=['POST'])
def chat_ia():
    if 'usuario' not in session: return jsonify({'resposta': 'Acesso negado.'})
    
    dados = request.get_json()
    mensagem_usuario = dados.get('mensagem', '')
    
    conn = sqlite3.connect('sistema_web.db')
    c = conn.cursor()
    c.execute("SELECT nome, estoque, min_estoque FROM pecas")
    pecas = c.fetchall()
    c.execute("SELECT nome, estoque, min_estoque FROM guepar_uso")
    guepar = c.fetchall()
    conn.close()
    
    pecas_falta = [p[0] for p in pecas if p[1] <= p[2]]
    guepar_falta = [g[0] for g in guepar if g[1] <= g[2]]
    
    # ========================================================
    # COLOQUE SUA CHAVE AQUI (ENTRE AS ASPAS)
    API_KEY = "AQ.Ab8RN6L9R0vz4amwMdY5T1WL_feFriBosGG3To7r-Ftz9KHkhQ" 
    # ========================================================
    
    if API_KEY == "SUA_CHAVE_AQUI" or API_KEY == "":
        resposta = f"🤖 [Modo Simulação Ativo]\nEu sou a IA e analisei seu banco de dados:\nVocê tem **{len(pecas)} peças de robô** e **{len(guepar)} materiais Guepar** cadastrados.\n\n"
        if pecas_falta or guepar_falta: resposta += f"⚠️ ATENÇÃO: Os seguintes itens estão acabando e precisam de reposição: {', '.join(pecas_falta + guepar_falta)}.\n\n"
        else: resposta += "✅ Seu estoque está perfeitamente saudável no momento.\n\n"
        resposta += "*Para conversar livremente comigo, cole a chave real no arquivo app.py!*"
    else:
        try:
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={API_KEY}"
            prompt_contexto = f"Você é o assistente virtual do painel 'Guepar vs RobÔ'. O estoque atual de peças do robô é: {pecas}. Os materiais da equipe são: {guepar}. O usuário perguntou: {mensagem_usuario}. Responda de forma direta e amigável em português do Brasil."
            payload = json.dumps({"contents": [{"parts":[{"text": prompt_contexto}]}]}).encode('utf-8')
            req = urllib.request.Request(url, data=payload, headers={'Content-Type': 'application/json'})
            
            # ATUALIZAÇÃO IMPORTANTE: Ignora verificação do Windows (resolve o erro de conexão)
            contexto_ssl = ssl._create_unverified_context()
            
            with urllib.request.urlopen(req, context=contexto_ssl) as response:
                result = json.loads(response.read().decode())
                resposta = result['candidates'][0]['content']['parts'][0]['text']
        
        # Se o Google recusar a chave, agora ele vai nos dizer exatamente o porquê
        except urllib.error.HTTPError as e:
            erro_detalhe = e.read().decode('utf-8')
            resposta = f"❌ Erro do Google ({e.code}): A API foi recusada. Detalhe: {erro_detalhe}"
        except Exception as e:
            resposta = f"❌ Erro no sistema/conexão: {str(e)}"

    return jsonify({'resposta': resposta})

@app.route('/add_peca', methods=['POST'])
def add_peca():
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("INSERT INTO pecas (nome, estoque, min_estoque) VALUES (?, ?, ?)", (request.form['nome'], request.form['estoque'], request.form['min_estoque']))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='pecas-robo'))

@app.route('/edit_peca/<int:id>', methods=['POST'])
def edit_peca(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("UPDATE pecas SET nome=?, estoque=?, min_estoque=? WHERE id=?", (request.form['nome'], request.form['estoque'], request.form['min_estoque'], id))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='pecas-robo'))

@app.route('/delete_peca/<int:id>')
def delete_peca(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("DELETE FROM pecas WHERE id=?", (id,))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='pecas-robo'))

@app.route('/add_guepar', methods=['POST'])
def add_guepar():
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("INSERT INTO guepar_uso (nome, estoque, min_estoque) VALUES (?, ?, ?)", (request.form['nome'], request.form['estoque'], request.form['min_estoque']))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='guepar-uso'))

@app.route('/edit_guepar/<int:id>', methods=['POST'])
def edit_guepar(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("UPDATE guepar_uso SET nome=?, estoque=?, min_estoque=? WHERE id=?", (request.form['nome'], request.form['estoque'], request.form['min_estoque'], id))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='guepar-uso'))

@app.route('/delete_guepar/<int:id>')
def delete_guepar(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("DELETE FROM guepar_uso WHERE id=?", (id,))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='guepar-uso'))

@app.route('/add_fornecedor', methods=['POST'])
def add_fornecedor():
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("INSERT INTO fornecedores (nome, contato, telefone, cnpj) VALUES (?, ?, ?, ?)", (request.form['nome'], request.form['contato'], request.form['telefone'], request.form.get('cnpj', '')))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='fornecedores'))

@app.route('/edit_fornecedor/<int:id>', methods=['POST'])
def edit_fornecedor(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("UPDATE fornecedores SET nome=?, contato=?, telefone=?, cnpj=? WHERE id=?", (request.form['nome'], request.form['contato'], request.form['telefone'], request.form.get('cnpj', ''), id))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='fornecedores'))

@app.route('/delete_fornecedor/<int:id>')
def delete_fornecedor(id):
    if 'usuario' in session:
        conn = sqlite3.connect('sistema_web.db')
        c = conn.cursor()
        c.execute("DELETE FROM fornecedores WHERE id=?", (id,))
        conn.commit()
        conn.close()
    return redirect(url_for('dashboard', tab='fornecedores'))

@app.route('/logout')
def logout():
    session.pop('usuario', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)