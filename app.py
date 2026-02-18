import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from io import BytesIO
import zipfile
import os
import re
import tempfile
import traceback

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Medições - Eng. Civil", layout="centered")
st.title("🏗️ Gerador Automático de Medições")

# --- FUNÇÕES AUXILIARES ---
def limpar_codigo(valor):
    if valor is None: return ""
    return str(valor).strip()

def para_float(valor):
    if valor is None: return 0.0
    if isinstance(valor, (int, float)): return float(valor)
    try:
        return float(str(valor).replace(',', '.'))
    except:
        return 0.0

def ordenar_abas_ruas(sheet_names):
    ruas = [s for s in sheet_names if s.upper().startswith(("RUA", "PAV", "TRECHO"))]
    def chave_ordenacao(nome):
        if "INICIAL" in nome.upper(): return 0
        numeros = re.findall(r'\d+', nome)
        if numeros: return int(numeros[0])
        return 999 
    return sorted(ruas, key=chave_ordenacao)

# --- PROCESSAMENTO (CORE) ---
def processar_arquivo(arquivo_analise, modelo_bytes, nome_modelo_original):
    
    extensao = os.path.splitext(nome_modelo_original)[1].lower()
    eh_macro = (extensao == '.xlsm')
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=extensao) as tmp_modelo:
        tmp_modelo.write(modelo_bytes.getvalue())
        tmp_path = tmp_modelo.name

    nome_arquivo_final = "medicao_padrao" + extensao 

    try:
        wb_modelo = openpyxl.load_workbook(tmp_path, keep_vba=eh_macro) 
        wb_analise = openpyxl.load_workbook(arquivo_analise, data_only=True)
        
        # --- 0. EXTRAÇÃO DO NOME DO ARQUIVO (H5 e K5) ---
        abas_ruas = ordenar_abas_ruas(wb_analise.sheetnames)
        nome_municipio = "MUNICIPIO"
        num_sam = "00"
        
        if abas_ruas:
            ws_info = wb_analise[abas_ruas[0]]
            
            val_h5 = ws_info['H5'].value
            if val_h5: 
                nome_municipio = str(val_h5).strip().replace(" ", "_").upper()
            
            val_k5 = ws_info['K5'].value
            if val_k5:
                num_sam = str(val_k5).strip()

        nome_arquivo_final = f"planilha_medição_{nome_municipio}_sam{num_sam}_lt_1_med_01__Lei14133_v08_2025{extensao}"

        # --- A. DADOS DA OBRA ---
        if 'Dados_Obra' in wb_analise.sheetnames and 'Dados_Obra' in wb_modelo.sheetnames:
            ws_orig = wb_analise['Dados_Obra']
            ws_dest = wb_modelo['Dados_Obra']
            for row in range(4, 61):
                for col in range(2, 41): 
                    val = ws_orig.cell(row=row, column=col).value
                    if val is not None:
                        try: ws_dest.cell(row=row, column=col).value = val
                        except: pass

        # --- B. LISTA MESTRA ---
        aba_origem_csv = None
        if 'CSV_GLOBAL' in wb_analise.sheetnames: aba_origem_csv = 'CSV_GLOBAL'
        elif 'CSV - Cartilha' in wb_analise.sheetnames: aba_origem_csv = 'CSV - Cartilha'
        
        ws_global_dest = None
        if 'Dados_GLOBAL_Inicial' in wb_modelo.sheetnames:
            ws_global_dest = wb_modelo['Dados_GLOBAL_Inicial']
        else:
            raise Exception("Aba 'Dados_GLOBAL_Inicial' não encontrada.")

        mapa_codigo_linha = {} 

        if aba_origem_csv and ws_global_dest:
            ws_orig = wb_analise[aba_origem_csv]
            max_row = ws_orig.max_row
            linha_dest = 10
            seq_contador = 1
            
            for r_orig in range(11, max_row + 1):
                filtro = ws_orig.cell(row=r_orig, column=4).value 
                
                if filtro and str(filtro).strip().upper() == 'X':
                    raw_code = ws_orig.cell(row=r_orig, column=6).value
                    code_key = limpar_codigo(raw_code)
                    
                    if code_key:
                        mapa_codigo_linha[code_key] = linha_dest

                    for c_orig in range(5, 12): 
                        c_dest = c_orig + 1 
                        if c_orig == 5: 
                            v_orig = ws_orig.cell(row=r_orig, column=c_orig).value
                            if v_orig and (isinstance(v_orig, (int, float)) or str(v_orig).strip().isdigit()):
                                try: 
                                    ws_global_dest.cell(row=linha_dest, column=c_dest).value = seq_contador
                                    seq_contador += 1
                                except: pass
                            else:
                                try: ws_global_dest.cell(row=linha_dest, column=c_dest).value = v_orig
                                except: pass
                        else:
                            val = ws_orig.cell(row=r_orig, column=c_orig).value
                            if val is not None:
                                try: ws_global_dest.cell(row=linha_dest, column=c_dest).value = val
                                except: pass
                    
                    linha_dest += 1

        # --- C. DISTRIBUIÇÃO POR RUA ---
        coluna_atual = 19 # S
        
        if ws_global_dest:
            for nome_aba in abas_ruas:
                ws_rua = wb_analise[nome_aba]
                
                try: ws_global_dest.cell(row=8, column=coluna_atual).value = nome_aba
                except: pass
                
                for linha_item in mapa_codigo_linha.values():
                    ws_global_dest.cell(row=linha_item, column=coluna_atual).value = 0

                max_r_rua = min(ws_rua.max_row, 5000)
                
                for r in range(11, max_r_rua + 1):
                    filtro_rua = ws_rua.cell(row=r, column=3).value
                    if filtro_rua and str(filtro_rua).strip().upper() == 'X':
                        raw_code_rua = ws_rua.cell(row=r, column=7).value
                        key_rua = limpar_codigo(raw_code_rua)
                        
                        qtd_rua = ws_rua.cell(row=r, column=20).value 
                        valor_numerico = para_float(qtd_rua)
                        
                        if key_rua and key_rua in mapa_codigo_linha:
                            linha_alvo = mapa_codigo_linha[key_rua]
                            
                            valor_atual = ws_global_dest.cell(row=linha_alvo, column=coluna_atual).value
                            if valor_atual is None: valor_atual = 0
                            
                            novo_total = float(valor_atual) + valor_numerico
                            ws_global_dest.cell(row=linha_alvo, column=coluna_atual).value = novo_total

                coluna_atual += 1
                if coluna_atual > 55: break 

        # --- D. FÓRMULAS ---
        try: ws_global_dest['N8'].value = len(abas_ruas)
        except: pass
        
        ultima_coluna_ruas_idx = coluna_atual - 1
        if ultima_coluna_ruas_idx >= 19:
            letra_ultima = get_column_letter(ultima_coluna_ruas_idx)
            for linha_item in mapa_codigo_linha.values():
                formula = f'=IF(K{linha_item}=0,"-",IF(ROUND(K{linha_item},2)=ROUND(SUM(S{linha_item}:{letra_ultima}{linha_item}),2),"Ok","Verificar"))'
                ws_global_dest.cell(row=linha_item, column=18).value = formula

        wb_modelo.save(tmp_path)
        wb_modelo.close()
        
        with open(tmp_path, "rb") as f:
            dados_finais = BytesIO(f.read())
            
    except Exception as e:
        raise e
    finally:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
            
    return dados_finais, nome_arquivo_final

# --- INTERFACE ---
col1, col2 = st.columns(2)
with col1:
    st.subheader("1. Arquivos de Análise")
    arquivos_analise = st.file_uploader("Arquivos .xlsx", accept_multiple_files=True)

with col2:
    st.subheader("2. Modelo")
    arquivo_modelo = st.file_uploader("Modelo (.xlsx/.xlsm)", type=['xlsx', 'xlsm'])

st.markdown("---")

# Container para o botão de ação (Placeholder)
# Isso permite substituir o botão por outro elemento (texto/botão inativo)
action_container = st.empty()

# Verificação se arquivos existem
if arquivos_analise and arquivo_modelo:
    
    # Renderiza o botão no placeholder
    if action_container.button("🚀 Processar"):
        
        # 1. SUBSTITUI O BOTÃO IMEDIATAMENTE POR "PROCESSANDO" (INATIVO)
        # Isso impede o clique duplo visualmente e dá feedback instantâneo
        action_container.button("Processando... ⏳", disabled=True)
        
        # 2. INICIA O SPINNER (Bolinha rodando)
        with st.spinner("Lendo arquivos e calculando medições..."):
            
            try:
                total_arquivos = len(arquivos_analise)
                modelo_bytes = BytesIO(arquivo_modelo.getvalue())
                nome_modelo = arquivo_modelo.name
                
                # --- CASO 1: ARQUIVO ÚNICO ---
                if total_arquivos == 1:
                    arq = arquivos_analise[0]
                    res, nome_gerado = processar_arquivo(arq, modelo_bytes, nome_modelo)
                    
                    st.success(f"✅ Sucesso! Arquivo gerado: {nome_gerado}")
                    
                    ext = os.path.splitext(nome_gerado)[1].lower()
                    mime_type = "application/vnd.ms-excel.sheet.macroEnabled.12" if ext == '.xlsm' else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    
                    st.download_button(
                        label="📥 Baixar Planilha",
                        data=res.getvalue(),
                        file_name=nome_gerado,
                        mime=mime_type
                    )

                # --- CASO 2: MÚLTIPLOS ARQUIVOS (ZIP) ---
                else:
                    zip_buffer = BytesIO()
                    
                    # Sem barra de status textual, apenas o spinner visual
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for i, arq in enumerate(arquivos_analise):
                            res, nome_gerado = processar_arquivo(arq, modelo_bytes, nome_modelo)
                            zip_file.writestr(nome_gerado, res.getvalue())
                    
                    st.success(f"✅ {total_arquivos} arquivos processados com sucesso!")
                    
                    st.download_button(
                        label="📥 Baixar ZIP",
                        data=zip_buffer.getvalue(),
                        file_name="Lote_Medicoes.zip",
                        mime="application/zip"
                    )
                    
            except Exception as e:
                # Em caso de erro, limpamos o botão de "Processando" para permitir tentar de novo
                action_container.empty()
                st.error(f"Ocorreu um erro: {e}")
                st.error(traceback.format_exc())
                # Recria o botão original para tentar novamente
                if st.button("Tentar Novamente"):
                     st.rerun()

else:
    # Mostra botão desabilitado se não houver arquivos
    st.button("🚀 Processar", disabled=True, help="Faça upload dos arquivos primeiro")