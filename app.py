import streamlit as st
import pandas as pd
from docx import Document
import io
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

# ==========================================
# FUNÇÕES AUXILIARES DE FORMATAÇÃO
# ==========================================
def encontrar_paragrafo(doc, texto_tag):
    for p in doc.paragraphs:
        if texto_tag in p.text: return p
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    if texto_tag in p.text: return p
    return None

def preencher_celula(celula, texto, fundo_cinza=False, negrito=False, alinhar_centro=True):
    """
    Injeta o texto com 3 espaços (1 linha acima e 1 abaixo),
    aplica alinhamento e fonte Calibri 12 (com opção de negrito).
    """
    # Adiciona os espaços em branco (linhas) antes e depois
    celula.text = f"\n{texto}\n"
    
    # Centralização Vertical
    celula.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    # Formatação do Parágrafo e Fonte (Calibri 12, Negrito e Alinhamento)
    for p in celula.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if alinhar_centro else WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.name = 'Calibri'
            run.font.size = Pt(12)
            run.font.bold = negrito
            
    # Aplica o fundo cinza se solicitado
    if fundo_cinza:
        tcPr = celula._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'D9D9D9') # Código Hex para Cinza-Claro
        tcPr.append(shd)

def definir_larguras(tabela, larguras):
    """
    Ajusta a largura de cada coluna para economizar espaço nos números.
    """
    for row in tabela.rows:
        for idx, celula in enumerate(row.cells):
            if idx < len(larguras):
                celula.width = larguras[idx]

# ==========================================
# MOTOR INTELIGENTE DE PREENCHIMENTO
# ==========================================
def preencher_documento(dados_pessoais, df_vinculos, df_remuneracoes):
    doc = Document("template.docx")
    
    # 1. Substituições de Texto Simples (Garante Calibri 12 nas tags)
    for paragrafo in doc.paragraphs:
        for chave, valor in dados_pessoais.items():
            if chave in paragrafo.text:
                paragrafo.text = paragrafo.text.replace(chave, str(valor))
                for run in paragrafo.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(12)
                
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                # CORREÇÃO: Lê parágrafo por parágrafo dentro da célula 
                # para não destruir as linhas de assinatura e formatações originais
                for p in celula.paragraphs:
                    for chave, valor in dados_pessoais.items():
                        if chave in p.text:
                            p.text = p.text.replace(chave, str(valor))
                            for run in p.runs:
                                run.font.name = 'Calibri'
                                run.font.size = Pt(12)

    # =========================================================
    # TABELA 1: DADOS FUNCIONAIS (1ª Página)
    # =========================================================
    p_func1 = encontrar_paragrafo(doc, '{{TAB_FUNC_1}}')
    if p_func1 is not None:
        tbl1 = doc.add_table(rows=0, cols=3)
        tbl1.style = 'Table Grid'
        
        for idx, row in df_vinculos.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
            v = str(row['Vínculo'])
            
            r_adm = tbl1.add_row().cells
            preencher_celula(r_adm[0], f"DATA DE ADMISSÃO NO VÍNCULO {v}:\n{row.get('Dt. Admissão', '-')}", alinhar_centro=False)
            preencher_celula(r_adm[1], f"Nº DE PORTARIA DE NOMEAÇÃO:\n{row.get('Port. Nomeação', 'NA')}", alinhar_centro=False)
            preencher_celula(r_adm[2], f"DATA DA PUBLICAÇÃO:\n{row.get('Pub. Nomeação', 'NA')}", alinhar_centro=False)
            
            r_des = tbl1.add_row().cells
            preencher_celula(r_des[0], f"DATA DE DESLIGAMENTO NO VÍNCULO {v}:\n{row.get('Dt. Desligamento', '-')}", alinhar_centro=False)
            preencher_celula(r_des[1], f"Nº DE PORTARIA DE EXONERAÇÃO/ DEMISSÃO:\n{row.get('Port. Exoneração', 'NA')}", alinhar_centro=False)
            preencher_celula(r_des[2], f"DATA DA PUBLICAÇÃO:\n{row.get('Pub. Exoneração', 'NA')}", alinhar_centro=False)
        
        # Ajusta larguras igualmente
        definir_larguras(tbl1, [Cm(5.3), Cm(5.3), Cm(5.3)])
        
        p_func1._p.addnext(tbl1._tbl)
        p_func1.text = p_func1.text.replace('{{TAB_FUNC_1}}', '')

    # =========================================================
    # TABELA 2: PERÍODOS DE CONTRIBUIÇÃO (1ª Página)
    # =========================================================
    p_per = encontrar_paragrafo(doc, '{{TAB_PER}}')
    if p_per is not None:
        tbl2 = doc.add_table(rows=1, cols=5)
        tbl2.style = 'Table Grid'
        headers = ["SEQ.:", "DATA INÍCIO:", "DATA FIM:", "CARGO/FUNÇÃO:", "CATEGORIA FUNCIONAL:"]
        for i, h in enumerate(headers):
            preencher_celula(tbl2.cell(0, i), h, negrito=True)
        
        for idx, row in df_vinculos.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
            seq = str(idx + 1).zfill(2)
            r = tbl2.add_row().cells
            
            cat = str(row.get('Categoria Funcional', ''))
            cat_efet = "X" if cat == "Efetivo/Estável" else " "
            cat_comis = "X" if cat == "Comissionado/Mandato Eletivo" else " "
            cat_contrat = "X" if cat == "Contratado" else " "
            cat_texto = f"({cat_efet}) Efetivo/Estável\n({cat_comis}) Comissionado/Mandato Eletivo\n({cat_contrat}) Contratado"

            preencher_celula(r[0], seq)
            preencher_celula(r[1], str(row.get('Dt. Admissão', '-')))
            preencher_celula(r[2], str(row.get('Dt. Desligamento', '-')))
            preencher_celula(r[3], str(row.get('Cargo / Função', '-')).upper())
            preencher_celula(r[4], cat_texto)

        # Mesclagem Vertical Inteligente
        if len(tbl2.rows) > 1:
            for col_idx in [3, 4]:
                start_row = 1
                while start_row < len(tbl2.rows):
                    end_row = start_row
                    while end_row + 1 < len(tbl2.rows) and tbl2.cell(end_row + 1, col_idx).text == tbl2.cell(start_row, col_idx).text:
                        end_row += 1
                    if end_row > start_row:
                        start_cell = tbl2.cell(start_row, col_idx)
                        raw_text = start_cell.text.strip() # Salva o texto sem as quebras de linha
                        for r_idx in range(start_row + 1, end_row + 1):
                            start_cell.merge(tbl2.cell(r_idx, col_idx))
                        preencher_celula(start_cell, raw_text) # Refaz a formatação na célula grande
                    start_row = end_row + 1

        # Deixa o SEQ. bem menor (1.5cm) e datas ajustadas
        definir_larguras(tbl2, [Cm(1.5), Cm(3.0), Cm(3.0), Cm(4.5), Cm(4.0)])
        
        p_per._p.addnext(tbl2._tbl)
        p_per.text = p_per.text.replace('{{TAB_PER}}', '')

    # =========================================================
    # TABELA 3: DADOS FUNCIONAIS 2 (2ª Página)
    # =========================================================
    p_func2 = encontrar_paragrafo(doc, '{{TAB_FUNC_2}}')
    if p_func2 is not None:
        tbl3 = doc.add_table(rows=0, cols=4)
        tbl3.style = 'Table Grid'
        
        for idx, row in df_vinculos.iterrows():
            if pd.isna(row.iloc[0]) or str(row.iloc[0]).strip() == "": continue
            v = str(row['Vínculo'])
            r = tbl3.add_row().cells
            
            preencher_celula(r[0], f"DATA DE ADMISSÃO\nNO VÍNCULO {v}:\n{row.get('Dt. Admissão', '-')}", alinhar_centro=False)
            preencher_celula(r[1], f"DATA DE EXONERAÇÃO\nNO VÍNCULO {v}:\n{row.get('Dt. Desligamento', '-')}", alinhar_centro=False)
            
            if idx == 0:
                preencher_celula(r[2], f"PIS/PASEP:\n{dados_pessoais.get('{{PIS}}', '')}", alinhar_centro=False)
                preencher_celula(r[3], f"CPF:\n{dados_pessoais.get('{{CPF}}', '')}", alinhar_centro=False)
            else:
                preencher_celula(r[2], "", alinhar_centro=False)
                preencher_celula(r[3], "", alinhar_centro=False)
        
        if len(tbl3.rows) > 1:
            for col_idx in [2, 3]:
                start_cell = tbl3.cell(0, col_idx)
                raw_text = start_cell.text.strip()
                for r_idx in range(1, len(tbl3.rows)):
                    start_cell.merge(tbl3.cell(r_idx, col_idx))
                preencher_celula(start_cell, raw_text, alinhar_centro=False)

        definir_larguras(tbl3, [Cm(4), Cm(4), Cm(4), Cm(4)])

        p_func2._p.addnext(tbl3._tbl)
        p_func2.text = p_func2.text.replace('{{TAB_FUNC_2}}', '')

    # =========================================================
    # TABELA 4: MATRIZ DE REMUNERAÇÕES PADRÃO INSS
    # =========================================================
    p_remun = encontrar_paragrafo(doc, '{{TABELAS_REMUNERACAO}}')
    if p_remun is not None:
        anos_encontrados = set()
        dados_matriz = {}
        for _, row in df_remuneracoes.iterrows():
            comp = str(row.get('Competência', '')).strip()
            val = str(row.get('Valor (R$)', '')).strip()
            if '/' in comp and val:
                try:
                    m, a = comp.split('/')
                    anos_encontrados.add(int(a))
                    dados_matriz[(int(m), int(a))] = val
                except:
                    pass

        anos_ordenados = sorted(list(anos_encontrados))
        blocos_de_anos = [anos_ordenados[i:i+5] for i in range(0, len(anos_ordenados), 5)]
        meses_nomes = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

        elemento_anterior = p_remun._p
        
        for bloco in blocos_de_anos:
            tbl = doc.add_table(rows=14, cols=len(bloco)+1)
            tbl.style = 'Table Grid'
            
            # Cabeçalho: Mês (Mesclado) com Fundo Cinza
            c_mes0 = tbl.cell(0, 0)
            c_mes0.merge(tbl.cell(1, 0))
            preencher_celula(c_mes0, "Mês", fundo_cinza=True, negrito=True)
            
            # Cabeçalho: Anos e "Valor($)" com Fundo Cinza
            for col_idx, ano in enumerate(bloco):
                c_ano = tbl.cell(0, col_idx+1)
                preencher_celula(c_ano, f"Ano: {ano}", fundo_cinza=True, negrito=True)
                
                c_val = tbl.cell(1, col_idx+1)
                preencher_celula(c_val, "Valor($)", fundo_cinza=True, negrito=True)
                
            # Preenchendo os meses e valores
            for mes_idx, mes_nome in enumerate(meses_nomes):
                c_m = tbl.cell(mes_idx+2, 0)
                preencher_celula(c_m, mes_nome)
                
                for col_idx, ano in enumerate(bloco):
                    val = dados_matriz.get((mes_idx+1, ano), "-")
                    c_v = tbl.cell(mes_idx+2, col_idx+1)
                    preencher_celula(c_v, val)
                    
            # Ajusta larguras: Mês ganha mais espaço (3.5cm), Valores ficam menores (2.5cm)
            larguras_t4 = [Cm(3.5)] + [Cm(2.5)] * len(bloco)
            definir_larguras(tbl, larguras_t4)
            
            elemento_anterior.addnext(tbl._tbl)
            elemento_anterior = tbl._tbl
            
            p_space = OxmlElement('w:p')
            elemento_anterior.addnext(p_space)
            elemento_anterior = p_space
            
        p_remun.text = p_remun.text.replace('{{TABELAS_REMUNERACAO}}', '')

    arquivo_gerado = io.BytesIO()
    doc.save(arquivo_gerado)
    arquivo_gerado.seek(0)
    return arquivo_gerado

# ==========================================
# INTERFACE DO STREAMLIT
# ==========================================
st.set_page_config(page_title="Preenchedor Automático DTC", layout="wide")

st.title("📄 Preenchedor Automático DTC")
st.write("Interface otimizada. O sistema desenhará magicamente as tabelas exatas do padrão INSS com fontes e cores corretas.")

# --- 1. Dados Pessoais ---
st.subheader("1. Dados Pessoais")
col1, col2, col3 = st.columns(3)
with col1:
    nome = st.text_input("Nome do Servidor")
    rg = st.text_input("RG / Órgão Emissor")
    mae = st.text_input("Nome da Mãe")
with col2:
    cpf = st.text_input("CPF")
    pis = st.text_input("PIS/PASEP")
with col3:
    matricula = st.text_input("Matrícula")
    data_nasc = st.text_input("Data de Nascimento")

dados_pessoais = {
    "{{NOME}}": nome, "{{CPF}}": cpf, "{{MATRICULA}}": matricula,
    "{{RG}}": rg, "{{PIS}}": pis, "{{MAE}}": mae, "{{DATA_NASC}}": data_nasc,
    "{{DATA}}": datetime.today().strftime('%d/%m/%Y')
}

st.markdown("---")

# --- 2. Vínculos Funcionais ---
st.subheader("2. Vínculos, Períodos e Cargos")

colunas_vinc = [
    "Vínculo", "Dt. Admissão", "Dt. Desligamento", 
    "Cargo / Função", "Categoria Funcional", 
    "Port. Nomeação", "Pub. Nomeação", "Port. Exoneração", "Pub. Exoneração"
]
df_vinc_vazio = pd.DataFrame(columns=colunas_vinc, index=range(3))

config_colunas = {
    "Categoria Funcional": st.column_config.SelectboxColumn(
        "Categoria Funcional",
        options=["Efetivo/Estável", "Comissionado/Mandato Eletivo", "Contratado"],
        required=True
    )
}

df_vinculos = st.data_editor(df_vinc_vazio, num_rows="dynamic", use_container_width=True, column_config=config_colunas)

st.markdown("---")

# --- 3. Remunerações ---
st.subheader("3. Remunerações")
st.success("⚡ **Acelerador:** Continue a digitar rapidamente um mês por linha. O sistema irá converter tudo em blocos de 5 anos!")

col_ano1, col_ano2, col_btn = st.columns([1, 1, 2])
with col_ano1:
    ano_inicio = st.number_input("Ano Inicial", min_value=1990, max_value=2050, value=2013, step=1)
with col_ano2:
    ano_fim = st.number_input("Ano Final", min_value=1990, max_value=2050, value=2024, step=1)

if 'df_remun_base' not in st.session_state:
    st.session_state.df_remun_base = pd.DataFrame([{"Competência": "", "Valor (R$)": ""}] * 5)

with col_btn:
    st.write("") 
    if st.button("⬇️ Criar Grade de Meses", type="secondary", use_container_width=True):
        meses = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
        linhas = []
        for ano in range(ano_inicio, ano_fim + 1):
            for mes in meses:
                linhas.append({"Competência": f"{mes}/{ano}", "Valor (R$)": ""})
        st.session_state.df_remun_base = pd.DataFrame(linhas)

df_remuneracoes = st.data_editor(st.session_state.df_remun_base, num_rows="dynamic", use_container_width=True, height=400)

# --- Botão de Geração Final ---
st.markdown("---")
if st.button("🚀 GERAR DOCUMENTO PADRÃO INSS", type="primary", use_container_width=True):
    if nome and cpf:
        with st.spinner('A desenhar as tabelas exatas do INSS com fontes e cores...'):
            try:
                arquivo_docx = preencher_documento(dados_pessoais, df_vinculos, df_remuneracoes)
                st.balloons()
                st.download_button(
                    label="📥 BAIXAR DTC (WORD)",
                    data=arquivo_docx,
                    file_name=f"DTC_{nome.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro ao processar as matrizes. Detalhe técnico: {e}")
    else:
        st.error("⚠️ Preencha pelo menos o Nome e o CPF do servidor.")

# --- Rodapé ---
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray; font-size: 12px;'><i>Desenvolvido por André Almeida Costa</i></p>", unsafe_allow_html=True)