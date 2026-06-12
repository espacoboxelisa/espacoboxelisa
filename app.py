import streamlit as st
import pandas as pd
import os
import unicodedata
import re
import calendar
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from fpdf import FPDF
from datetime import datetime

# ================= CONFIGURAÇÕES DE E-MAIL ================= #
# Substitua pelos seus dados reais
EMAIL_REMETENTE = "seu_email_que_vai_enviar@gmail.com"
SENHA_APP_EMAIL = "sua_senha_de_app_de_16_digitos"
EMAIL_ADMINISTRACAO = "email_da_administracao_que_vai_receber@gmail.com"

# ================= TEMA E CSS SOFISTICADO ================= #
st.set_page_config(page_title="Sistema de Locação", layout="wide")
st.markdown("""
    <style>
        [data-testid="stAppViewContainer"] { background-color: #1e1e24; color: #e0e0e0; }
        [data-testid="stHeader"] { background-color: #1e1e24; }
        [data-testid="stSidebar"] { background-color: #2b2b36; }
        label, label p, div[data-testid="stWidgetLabel"] p {
            color: #4fc3f7 !important;
            font-weight: 600 !important;
            font-size: 15px !important;
        }
        [data-testid="metric-container"] {
            background-color: #2b2b36;
            border-radius: 8px;
            padding: 15px;
            border: 1px solid #3d3d4d;
            box-shadow: 0 4px 6px rgba(0,0,0,0.3);
        }
        h1, h2, h3, h4, h5, h6 { color: #ffffff !important; }
    </style>
""", unsafe_allow_html=True)

# ================= CONFIGURAÇÕES INICIAIS ================= #
NOME_CSV = "recebimentos_aluguel.xlsx - Cadastro.csv"
MESES_LISTA = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
               'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
COLUNAS_BASE = ['Andar', 'Sala', 'Nome', 'CPF', 'Data de Nascimento', 'Endereço', 
                'E-mail', 'Locador', 'Valor (R$)'] + MESES_LISTA

if 'df_sistema' not in st.session_state:
    tipos_colunas = {col: str for col in COLUNAS_BASE}
    if os.path.exists(NOME_CSV):
        try:
            df_carregado = pd.read_csv(NOME_CSV, dtype=tipos_colunas, sep=';', encoding='utf-8-sig').reset_index(drop=True)
            if len(df_carregado.columns) < 5: df_carregado = pd.read_csv(NOME_CSV, dtype=tipos_colunas, sep=',').reset_index(drop=True)
        except:
            df_carregado = pd.read_csv(NOME_CSV, sep=',').reset_index(drop=True)
            
        df_carregado = df_carregado.astype(object).fillna('')
        for col in COLUNAS_BASE:
            if col not in df_carregado.columns: df_carregado[col] = ''
            else: df_carregado[col] = df_carregado[col].astype(str).replace('nan', '').replace('NaN', '')
        
        df_carregado = df_carregado[COLUNAS_BASE]
        st.session_state.df_sistema = df_carregado
    else:
        st.session_state.df_sistema = pd.DataFrame(columns=COLUNAS_BASE)

# ================= FUNÇÕES DO SISTEMA ================= #
def guardar_na_base_dados():
    st.session_state.df_sistema = st.session_state.df_sistema[COLUNAS_BASE]
    st.session_state.df_sistema.to_csv(NOME_CSV, index=False, sep=';', encoding='utf-8-sig')

def converter_para_float(valor):
    if pd.isna(valor) or str(valor).strip() == "": return 0.0
    texto_limpo = str(valor).replace('R$', '').replace(' ', '').strip()
    if ',' in texto_limpo and '.' in texto_limpo: texto_limpo = texto_limpo.replace('.', '').replace(',', '.')
    elif ',' in texto_limpo: texto_limpo = texto_limpo.replace(',', '.')
    try: return float(texto_limpo)
    except ValueError: return 0.0

def formatar_moeda(valor):
    try:
        valor_float = float(valor)
        return f"R$ {valor_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError): return "R$ 0,00"

def remover_acentos(texto):
    if not isinstance(texto, str): texto = str(texto)
    return "".join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def aplicar_mascara_cpf(cpf_str):
    cpf_numeros = re.sub(r'\D', '', str(cpf_str)) 
    if len(cpf_numeros) == 11: return f"{cpf_numeros[:3]}.{cpf_numeros[3:6]}.{cpf_numeros[6:9]}-{cpf_numeros[9:]}"
    return str(cpf_str).strip()

def disparar_email_admin(caminho_pdf, nome_locatario, sala, mes_ref, valor_total, recebedor):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_REMETENTE
    msg['To'] = EMAIL_ADMINISTRACAO
    msg['Subject'] = f"✅ Pagamento Recebido - {nome_locatario} (Sala {sala}) - {mes_ref}"

    corpo = f"""
    Olá Administração,
    
    Um novo pagamento de locação foi processado no sistema.
    
    DETALHES DA OPERAÇÃO:
    - Locatário: {nome_locatario}
    - Sala: {sala}
    - Mês de Referência: {mes_ref}
    - Valor Total Recebido: {formatar_moeda(valor_total)}
    - Recebedor: {recebedor}
    - Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}
    
    O recibo oficial em PDF segue em anexo para os registros financeiros.
    
    Att,
    Painel de Locação Automatizado
    """
    msg.attach(MIMEText(corpo, 'plain'))

    try:
        with open(caminho_pdf, "rb") as f:
            anexo = MIMEApplication(f.read(), _subtype="pdf")
            anexo.add_header('Content-Disposition', 'attachment', filename=os.path.basename(caminho_pdf))
            msg.attach(anexo)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP_EMAIL)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Erro ao enviar o e-mail: {e}")
        return False

def colorir_meses(val):
    if not isinstance(val, str) or val.strip() == "": return ""
    val_upper = val.upper()
    if val_upper.startswith("PAGO"): return "background-color: #2e7d32; color: #ffffff; font-weight: bold;" 
    if val_upper == "-" or "PRÉ-CONTRATO" in val_upper: return "background-color: #424242; color: #9e9e9e;" 
    if val_upper.startswith("VENCE EM"):
        try:
            data_venc = datetime.strptime(val_upper.replace("VENCE EM", "").strip(), "%d/%m/%Y").date()
            hoje = datetime.today().date()
            if data_venc < hoje: return "background-color: #c62828; color: #ffffff; font-weight: bold;" 
            elif data_venc == hoje: return "background-color: #ef6c00; color: #ffffff; font-weight: bold;" 
            else: return "background-color: #fbc02d; color: #212121; font-weight: bold;" 
        except: return ""
    return ""

def mapear_dados_mes(celula, valor_base, andar, sala, nome, cpf, mes_nome):
    dados = {
        'Andar': andar, 'Sala': sala, 'Nome': nome, 'CPF': cpf, 'Mês de Referencia': mes_nome,
        'Valor (R$)': valor_base, 'Vencimento': '', 'Data do Pagamento': '',
        'Juros': 0.0, 'Multa': 0.0, 'Valor Recebido': valor_base, 'Forma de Pagamento': '',
        'Observações': '', 'Status': 'Pré-contrato' if celula == "-" else 'Pendente'
    }
    if not isinstance(celula, str) or celula.strip() == "": return dados
    if celula.upper().startswith("VENCE EM"):
        dados['Vencimento'] = celula.upper().replace("VENCE EM", "").strip()
        dados['Status'] = 'Pendente'
    elif celula.upper().startswith("PAGO"):
        dados['Status'] = 'Pago'
        for parte in celula.split('|'):
            parte = parte.strip().upper()
            if parte.startswith("PAGO EM"): dados['Data do Pagamento'] = parte.replace("PAGO EM", "").strip()
            elif parte.startswith("TOTAL:"): dados['Valor Recebido'] = parte.split(':')[1].strip()
            elif parte.startswith("JUROS:"): dados['Juros'] = parte.split(':')[1].strip()
            elif parte.startswith("MULTA:"): dados['Multa'] = parte.split(':')[1].strip()
            elif parte.startswith("FORMA:"): dados['Forma de Pagamento'] = parte.split(':')[1].strip()
            elif parte.startswith("OBS:"): dados['Observações'] = parte.split(':')[1].strip()
            elif parte.startswith("VENC:"): dados['Vencimento'] = parte.split(':')[1].strip()
    return dados

def gerar_recibo_pdf(dados):
    pdf = FPDF()
    pdf.add_page()
    
    valor_aluguel = converter_para_float(dados.get('Valor (R$)', 0))
    juros = converter_para_float(dados.get('Juros', 0))
    multa = converter_para_float(dados.get('Multa', 0))
    total_recebido = converter_para_float(dados.get('Valor Recebido', 0))

    pdf.set_font("Helvetica", 'B', 16)
    pdf.cell(0, 15, "RECIBO DE PAGAMENTO DE ALUGUEL", ln=True, align='C')
    pdf.ln(4)
    pdf.line(10, 28, 200, 28)
    pdf.ln(5)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Mes de Referencia:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, remover_acentos(str(dados.get('Mês de Referencia', ''))).upper(), 0, 1)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Locatario:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    cpf = str(dados.get('CPF', '')).strip()
    pdf.cell(0, 8, f"{remover_acentos(str(dados.get('Nome', ''))).upper()}{f' (CPF: {cpf})' if cpf else ''}", 0, 1)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Localizacao / Sala:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, f"SALA {str(dados.get('Sala', '')).upper()} - ANDAR: {remover_acentos(str(dados.get('Andar', ''))).upper()}", 0, 1)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Vencimento Original:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, str(dados.get('Vencimento', '')), 0, 1)

    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Data do Pagamento:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, str(dados.get('Data do Pagamento', '')), 0, 1)
    
    dias_atraso = dados.get('Dias de Atraso', 0)
    if dias_atraso > 0:
        pdf.set_font("Helvetica", 'BI', 10)
        pdf.set_text_color(200, 0, 0)
        pdf.cell(0, 8, f"Pagamento efetuado com {dias_atraso} dia(s) de atraso. Sujeito a correcoes.", 0, 1)
        pdf.set_text_color(0, 0, 0)
    
    pdf.ln(2)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(5)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Valor do Aluguel:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, formatar_moeda(valor_aluguel), 0, 1)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Juros:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, formatar_moeda(juros), 0, 1)
    
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Multa:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, formatar_moeda(multa), 0, 1)
    
    pdf.ln(2)
    pdf.set_font("Helvetica", 'B', 13)
    pdf.cell(45, 10, "TOTAL RECEBIDO:", 0, 0)
    pdf.cell(0, 10, formatar_moeda(total_recebido), 0, 1)
    
    pdf.ln(2)
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(45, 8, "Forma de Pagamento:", 0, 0)
    pdf.set_font("Helvetica", size=11)
    pdf.cell(0, 8, remover_acentos(str(dados.get('Forma de Pagamento', ''))).upper(), 0, 1)
    
    obs = remover_acentos(str(dados.get('Observações', ''))).strip().upper()
    if obs and obs != "NAN":
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 8, "Observacoes:", 0, 1)
        pdf.set_font("Helvetica", size=11)
        pdf.multi_cell(0, 8, obs)
        
    pdf.ln(20)
    
    locador_nome = remover_acentos(str(dados.get('Locador', 'LOCADOR')).upper())
    recebedor_nome = remover_acentos(str(dados.get('Recebedor', 'RECEBEDOR')).upper())
    
    pdf.cell(90, 8, "___________________________________", 0, 0, 'C')
    pdf.cell(90, 8, "___________________________________", 0, 1, 'C')
    pdf.cell(90, 6, f"LOCADOR: {locador_nome}", 0, 0, 'C')
    pdf.cell(90, 6, f"RECEBEDOR: {recebedor_nome}", 0, 1, 'C')
    
    mes_ref_limpo = remover_acentos(str(dados.get('Mês de Referencia', 'Sem_Mes'))).replace("/", "-")
    data_pag_limpa = str(dados.get('Data do Pagamento', 'Sem_Data')).replace("/", "-")
    sala = remover_acentos(str(dados.get('Sala', 'Sem_Sala')))
    locatario = remover_acentos(str(dados.get('Nome', 'Sem_Nome')))
    
    pasta_mensal = f"Recibos_{mes_ref_limpo}"
    if not os.path.exists(pasta_mensal): os.makedirs(pasta_mensal)
        
    caminho_completo = os.path.join(pasta_mensal, f"{mes_ref_limpo} - Sala {sala} - {locatario} - {data_pag_limpa}.pdf")
    pdf.output(caminho_completo)
    return caminho_completo

# ================= LAYOUT DO APLICATIVO ================= #
st.title("🏢 Painel Integrado de Alugueres e Recibos")

aba_recibo, aba_cadastro, aba_base = st.tabs([
    "🔍 Pesquisar & Emitir Recibo", "➕ Novo Cadastro (Inquilino)", "📊 Dashboard & Base de Dados"
])

# ================= ABA 1: RECIBO ================= #
with aba_recibo:
    st.subheader("Cálculo Pró-rata e Emissão de Recibos")
    if 'termo_busca_anterior' not in st.session_state: st.session_state.termo_busca_anterior = ""
    termo_busca = st.text_input("Introduza o Nome do Locatário ou o Número da Sala para pesquisar:", key="busca_recibo")
    
    if termo_busca != st.session_state.termo_busca_anterior:
        st.session_state.termo_busca_anterior = termo_busca
        if 'caminho_pdf_gerado' in st.session_state: del st.session_state['caminho_pdf_gerado'] 
    
    if termo_busca:
        df_atual = st.session_state.df_sistema
        resultado = df_atual[df_atual['Nome'].astype(str).str.contains(termo_busca, case=False, na=False) | 
                             df_atual['Sala'].astype(str).str.contains(termo_busca, case=False, na=False)]
        
        if not resultado.empty:
            indices = resultado.index.tolist()
            opcoes_seletor = [f"Sala {resultado.loc[i, 'Sala']} | {resultado.loc[i, 'Nome']} | Base: {formatar_moeda(resultado.loc[i, 'Valor (R$)'])}" for i in indices]
            selecionado = st.selectbox("Selecione o locatário correspondente:", options=opcoes_seletor)
            idx_real = indices[opcoes_seletor.index(selecionado)]
            registo = df_atual.loc[idx_real]
            
            mes_selecionado = st.selectbox("Selecione o Mês para dar baixa / emitir recibo:", MESES_LISTA)
            
            chave_transacao = f"{idx_real}_{mes_selecionado}"
            if 'chave_anterior' not in st.session_state or st.session_state.chave_anterior != chave_transacao:
                st.session_state.chave_anterior = chave_transacao
                if 'caminho_pdf_gerado' in st.session_state: del st.session_state['caminho_pdf_gerado']

            celula_atual = str(registo.get(mes_selecionado, ''))
            dados_mes = mapear_dados_mes(celula_atual, registo['Valor (R$)'], registo['Andar'], registo['Sala'], registo['Nome'], registo['CPF'], mes_selecionado)
            
            st.write("---")
            if dados_mes['Status'] == 'Pago': st.warning(f"⚠️ O mês de **{mes_selecionado}** já consta como **PAGO** em {dados_mes['Data do Pagamento']}.")
            elif dados_mes['Status'] == 'Pré-contrato': st.error(f"⚠️ **{mes_selecionado}** é um mês anterior ao início do contrato deste inquilino.")
            else: st.info(f"📅 Status Atual de **{mes_selecionado}**: {celula_atual}")
                
            c1, c2 = st.columns(2)
            valor_base_seguro = converter_para_float(registo.get('Valor (R$)', 0))
            
            with c1:
                st.markdown(f"#### Contrato: **{registo['Nome']}**")
                st.write(f"**Locador Registrado:** {registo.get('Locador', 'Não informado')}")
                st.write(f"**Andar / Sala:** {registo['Andar']} / SALA {registo['Sala']}")
                venc_str = dados_mes['Vencimento']
                st.write(f"**Vencimento Original:** {venc_str if venc_str else 'Não definido'}")
                st.markdown("##### Dados da Emissão")
                recebedor_input = st.text_input("Nome do Recebedor (Assinatura)", value="Administração")
            
            with c2:
                data_pag_input = st.date_input("Data do Pagamento", value=datetime.today(), format="DD/MM/YYYY")
                dias_atraso = 0
                if venc_str:
                    try:
                        dias_atraso = max(0, (data_pag_input - datetime.strptime(venc_str, "%d/%m/%Y").date()).days)
                    except: pass
                
                if dias_atraso > 0: st.error(f"🚨 **ALERTA:** Pagamento com **{dias_atraso} dias de atraso**.")
                elif dados_mes['Status'] != 'Pré-contrato': st.success("✅ Pagamento dentro do prazo estipulado.")
                    
                taxa_multa_pct = st.number_input("Multa (%)", min_value=0.0, value=2.0 if dias_atraso > 0 else 0.0, step=0.50, format="%.2f")
                taxa_juros_pct = st.number_input("Juros ao Mês (%) - Pro-rata", min_value=0.0, value=1.0 if dias_atraso > 0 else 0.0, step=0.50, format="%.2f")
                forma_pagamento = st.selectbox("Forma de Pagamento", ["PIX", "DINHEIRO", "TRANSFERÊNCIA", "CARTÃO", "BOLETO"])
                observacoes_input = st.text_area("Observações do Recibo", value=dados_mes['Observações'])
            
            valor_multa_reais = valor_base_seguro * (taxa_multa_pct / 100) if dias_atraso > 0 else 0.0
            valor_juros_reais = (valor_base_seguro * (taxa_juros_pct / 100) / 30 * dias_atraso) if dias_atraso > 0 else 0.0
            valor_total_calculado = valor_base_seguro + valor_juros_reais + valor_multa_reais
            
            st.markdown(f"""
            <div style="background-color: #2b2b36; padding: 20px; border-radius: 8px; border: 1px solid #4caf50;">
                <h3 style="margin-top: 0; color: #4caf50;">💰 Valor Final a Pagar: {formatar_moeda(valor_total_calculado)}</h3>
                <p style="margin-bottom: 0;">Base: {formatar_moeda(valor_base_seguro)} | Juros Pró-rata: {formatar_moeda(valor_juros_reais)} | Multa: {formatar_moeda(valor_multa_reais)}</p>
            </div><br>
            """, unsafe_allow_html=True)
            
            if st.button("💾 Confirmar Pagamento e Gerar PDF", type="primary", use_container_width=True):
                obs_final = str(observacoes_input).upper() if str(observacoes_input).strip() else "NENHUMA"
                data_pag_str = data_pag_input.strftime("%d/%m/%Y")
                venc_salvo = venc_str if venc_str else "11/11/2222"
                
                celula_nova = f"PAGO em {data_pag_str} | Total: {valor_total_calculado:.2f} | Juros: {valor_juros_reais:.2f} | Multa: {valor_multa_reais:.2f} | Forma: {forma_pagamento.upper()} | Obs: {obs_final} | Venc: {venc_salvo}"
                st.session_state.df_sistema.loc[idx_real, mes_selecionado] = celula_nova
                guardar_na_base_dados()
                
                dados_para_pdf = {
                    'Andar': registo['Andar'], 'Sala': registo['Sala'], 'Nome': registo['Nome'], 'CPF': registo['CPF'],
                    'Mês de Referencia': mes_selecionado.upper(), 'Vencimento': venc_salvo, 'Valor (R$)': valor_base_seguro,
                    'Data do Pagamento': data_pag_str, 'Juros': valor_juros_reais, 'Multa': valor_multa_reais,
                    'Valor Recebido': valor_total_calculado, 'Forma de Pagamento': forma_pagamento, 'Observações': obs_final,
                    'Dias de Atraso': dias_atraso, 'Locador': registo.get('Locador', ''), 'Recebedor': recebedor_input
                }
                
                caminho_pdf = gerar_recibo_pdf(dados_para_pdf)
                st.session_state['caminho_pdf_gerado'] = caminho_pdf
                
                # ENVIO DE EMAIL AQUI
                with st.spinner("Enviando recibo para a Administração..."):
                    sucesso_email = disparar_email_admin(caminho_pdf, registo['Nome'], registo['Sala'], mes_selecionado, valor_total_calculado, recebedor_input)
                
                if sucesso_email: st.success(f"🎉 Baixa efetuada em {mes_selecionado} e E-mail enviado com sucesso!")
                else: st.warning(f"Baixa efetuada e PDF gerado, mas ocorreu um erro ao enviar o e-mail (Verifique as credenciais).")
                
            if 'caminho_pdf_gerado' in st.session_state and os.path.exists(st.session_state['caminho_pdf_gerado']):
                with open(st.session_state['caminho_pdf_gerado'], "rb") as arquivo:
                    st.download_button("📥 Descarregar PDF de " + mes_selecionado, data=arquivo.read(), file_name=os.path.basename(st.session_state['caminho_pdf_gerado']), mime="application/pdf", type="secondary", use_container_width=True)
        else:
            st.warning("⚠️ Nenhum registro encontrado.")

# ================= ABA 2: CADASTRO ================= #
with aba_cadastro:
    st.subheader("Formulário de Cadastro Completo (Novo Inquilino)")
    
    with st.form("novo_registro_completo", clear_on_submit=True):
        col_esq, col_mid, col_dir = st.columns(3)
        with col_esq:
            andar_selecionado = st.selectbox("Andar", ["TÉRREO", "1º", "2º", "3º", "4º", "5º", "6º", "OUTRO"])
            novo_andar_custom = st.text_input("Se escolheu 'OUTRO', digite o andar aqui:")
            nova_sala = st.text_input("Número da Sala (Ex: 05, 12)")
            novo_nome = st.text_input("Nome Completo do Locatário")
            novo_cpf = st.text_input("CPF (Opcional)", max_chars=14)
        with col_mid:
            novo_nascimento = st.text_input("Data de Nascimento (Ex: DD/MM/AAAA)")
            novo_endereco = st.text_area("Endereço Completo")
            novo_email = st.text_input("E-mail de Contato")
        with col_dir:
            novo_locador = st.text_input("Nome do Locador Responsável", value="Administração / Proprietário")
            novo_valor_aluguel = st.number_input("Valor Mensal do Aluguel (R$)", min_value=0.0, step=10.0, format="%.2f")
            novo_ano = st.number_input("Ano de Referência", min_value=2020, max_value=2100, value=datetime.today().year, step=1)
            mes_inicio = st.selectbox("Mês de Início do Contrato", MESES_LISTA)
            dia_vencimento = st.number_input("Dia Fixo de Vencimento (Ex: 11)", min_value=1, max_value=31, value=11, step=1)
            
        if st.form_submit_button("Gravar Cadastro na Base"):
            andar_final = novo_andar_custom.strip().upper() if andar_selecionado == "OUTRO" else andar_selecionado
            sala_upper = nova_sala.strip().upper()
            nome_upper = novo_nome.strip().upper()
            cpf_formatado = aplicar_mascara_cpf(novo_cpf.strip().upper())
            
            if nome_upper and sala_upper and andar_final:
                df_atual = st.session_state.df_sistema
                if not df_atual[(df_atual['Nome'].astype(str).str.upper() == nome_upper) & (df_atual['Sala'].astype(str).str.upper() == sala_upper)].empty:
                    st.error(f"❌ Erro: '{nome_upper}' já está cadastrado na Sala '{sala_upper}'!")
                else:
                    nova_linha_dados = {
                        'Andar': str(andar_final), 'Sala': str(sala_upper), 'Nome': str(nome_upper), 
                        'CPF': str(cpf_formatado), 'Data de Nascimento': str(novo_nascimento), 
                        'Endereço': str(novo_endereco).upper(), 'E-mail': str(novo_email),
                        'Locador': str(novo_locador).upper(), 'Valor (R$)': f"{novo_valor_aluguel:.2f}".replace('.', ',')
                    }
                    
                    idx_mes_inicio = MESES_LISTA.index(mes_inicio)
                    for idx, m_nome in enumerate(MESES_LISTA):
                        if idx < idx_mes_inicio: nova_linha_dados[m_nome] = "-"
                        else:
                            idx_mes_real = idx + 1
                            dia_real = min(int(dia_vencimento), calendar.monthrange(int(novo_ano), idx_mes_real)[1])
                            nova_linha_dados[m_nome] = f"Vence em {datetime(int(novo_ano), idx_mes_real, dia_real).strftime('%d/%m/%Y')}"
                    
                    st.session_state.df_sistema = pd.concat([st.session_state.df_sistema, pd.DataFrame([nova_linha_dados])], ignore_index=True)
                    guardar_na_base_dados()
                    st.success(f"✅ Cadastro anual criado com sucesso para {nome_upper}!")
            else:
                st.error("❌ Preencha os campos obrigatórios (Andar, Nome e Sala).")

# ================= ABA 3: DASHBOARD E GESTÃO ================= #
with aba_base:
    st.subheader("Dashboard Gerencial e Base de Dados")
    
    df_atual = st.session_state.df_sistema
    total_contratos = len(df_atual) if not df_atual.empty else 0
    total_inadimplentes, total_pagos = 0, 0
    
    if not df_atual.empty:
        hoje = datetime.today().date()
        for index, row in df_atual.iterrows():
            for m_nome in MESES_LISTA:
                celula = str(row.get(m_nome, '')).upper()
                if celula.startswith("PAGO"): total_pagos += 1
                elif celula.startswith("VENCE EM"):
                    try:
                        if datetime.strptime(celula.replace("VENCE EM", "").strip(), "%d/%m/%Y").date() < hoje: total_inadimplentes += 1
                    except: pass

    m1, m2, m3 = st.columns(3)
    m1.metric("Contratos Ativos", total_contratos)
    m2.metric("Mensalidades Adimplentes (Pagas)", total_pagos)
    m3.metric("Mensalidades Inadimplentes (Atrasadas)", total_inadimplentes, delta="- Cuidado Financeiro" if total_inadimplentes > 0 else "Tudo em Dia", delta_color="inverse")
    st.write("---")
    
    acao_db = st.radio("Selecione a ação na Base de Dados:", ["📊 Visualizar Tabela Completa", "✏️ Editar Cadastro", "🗑️ Excluir Cadastro"], horizontal=True)
    st.divider()
    
    if acao_db == "📊 Visualizar Tabela Completa":
        if not df_atual.empty:
            try: styled_df = df_atual.style.map(colorir_meses, subset=MESES_LISTA)
            except AttributeError: styled_df = df_atual.style.applymap(colorir_meses, subset=MESES_LISTA)
            st.dataframe(styled_df, use_container_width=True)
        else: st.dataframe(df_atual)
            
        st.download_button("📥 Exportar Base Atualizada para Excel", data=df_atual.to_csv(index=False, sep=';', encoding='utf-8-sig'), file_name="Planilha_Alugueis_Anual.csv", mime="text/csv")
        
    elif acao_db == "✏️ Editar Cadastro":
        if not df_atual.empty:
            busca_editar = st.text_input("Introduza o Nome do Locatário ou Sala para Editar:", key="busca_ed")
            if busca_editar:
                res_ed = df_atual[df_atual['Nome'].astype(str).str.contains(busca_editar, case=False, na=False) | 
                                  df_atual['Sala'].astype(str).str.contains(busca_editar, case=False, na=False)]
                if not res_ed.empty:
                    opcoes_ed = [f"Linha {i} | Sala {df_atual.loc[i, 'Sala']} - {df_atual.loc[i, 'Nome']}" for i in res_ed.index]
                    selecao_ed = st.selectbox("Selecione qual cadastro deseja editar:", opcoes_ed)
                    idx_ed = res_ed.index[opcoes_ed.index(selecao_ed)]
                    reg_ed = df_atual.loc[idx_ed]
                    
                    with st.form("form_editar_cadastro"):
                        c1, c2, c3 = st.columns(3)
                        edit_andar = c1.text_input("Andar", value=str(reg_ed['Andar']))
                        edit_sala = c1.text_input("Sala", value=str(reg_ed['Sala']))
                        edit_nome = c1.text_input("Nome", value=str(reg_ed['Nome']))
                        
                        edit_cpf = c2.text_input("CPF", value=str(reg_ed['CPF']), max_chars=14)
                        edit_nascimento = c2.text_input("Nascimento", value=str(reg_ed.get('Data de Nascimento', '')))
                        edit_email = c2.text_input("E-mail", value=str(reg_ed.get('E-mail', '')))
                        
                        edit_endereco = c3.text_input("Endereço", value=str(reg_ed.get('Endereço', '')))
                        edit_locador = c3.text_input("Locador", value=str(reg_ed.get('Locador', '')))
                        edit_valor = c3.text_input("Valor Base (R$)", value=str(reg_ed['Valor (R$)']))
                        
                        st.markdown("##### ✏️ Ajuste Fino dos Meses")
                        with st.expander("Expandir Histórico Financeiro"):
                            edit_meses = {}
                            m_cols = st.columns(4)
                            for idx, m_nome in enumerate(MESES_LISTA):
                                with m_cols[idx % 4]: edit_meses[m_nome] = st.text_input(m_nome, value=str(reg_ed.get(m_nome, '')))
                        
                        if st.form_submit_button("💾 Confirmar Edição"):
                            st.session_state.df_sistema.loc[idx_ed, 'Andar'] = str(edit_andar).strip().upper()
                            st.session_state.df_sistema.loc[idx_ed, 'Sala'] = str(edit_sala).strip().upper()
                            st.session_state.df_sistema.loc[idx_ed, 'Nome'] = str(edit_nome).strip().upper()
                            st.session_state.df_sistema.loc[idx_ed, 'CPF'] = aplicar_mascara_cpf(str(edit_cpf).strip().upper())
                            st.session_state.df_sistema.loc[idx_ed, 'Data de Nascimento'] = str(edit_nascimento).strip()
                            st.session_state.df_sistema.loc[idx_ed, 'E-mail'] = str(edit_email).strip()
                            st.session_state.df_sistema.loc[idx_ed, 'Endereço'] = str(edit_endereco).strip().upper()
                            st.session_state.df_sistema.loc[idx_ed, 'Locador'] = str(edit_locador).strip().upper()
                            st.session_state.df_sistema.loc[idx_ed, 'Valor (R$)'] = str(edit_valor).strip()
                            for m_nome in MESES_LISTA: st.session_state.df_sistema.loc[idx_ed, m_nome] = str(edit_meses[m_nome]).strip()
                            guardar_na_base_dados()
                            st.success("✅ O Cadastro foi atualizado!")
                            st.rerun()
                else: st.warning("Nenhum registro encontrado.")

    elif acao_db == "🗑️ Excluir Cadastro":
        if not df_atual.empty:
            busca_ex = st.text_input("Introduza o Nome do Locatário ou Sala para Excluir:", key="busca_ex")
            if busca_ex:
                res_ex = df_atual[df_atual['Nome'].astype(str).str.contains(busca_ex, case=False, na=False) | 
                                  df_atual['Sala'].astype(str).str.contains(busca_ex, case=False, na=False)]
                if not res_ex.empty:
                    opcoes_ex = [f"Linha {i} | Sala {df_atual.loc[i, 'Sala']} - {df_atual.loc[i, 'Nome']}" for i in res_ex.index]
                    selecao_ex = st.selectbox("Selecione qual cadastro EXCLUIR:", opcoes_ex)
                    idx_ex = res_ex.index[opcoes_ex.index(selecao_ex)]
                    st.warning(f"Você apagará **{df_atual.loc[idx_ex, 'Nome']}** permanentemente.")
                    if st.button("🚨 EXCLUIR DEFINITIVAMENTE") and st.checkbox("Tenho certeza"):
                        st.session_state.df_sistema = st.session_state.df_sistema.drop(idx_ex).reset_index(drop=True)
                        guardar_na_base_dados()
                        st.success("🗑️ Linha removida!")
                        st.rerun()
                else: st.warning("Nenhum registro encontrado.")