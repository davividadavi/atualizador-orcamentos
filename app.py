import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Atualizador de Orçamento SINAPI", layout="centered")

st.title("🔄 Atualizador Automático de Orçamento")
st.write("Faça o upload da sua Planilha Orçamentária e da Referência do Orçafascio para atualizar os valores automaticamente.")

col1, col2 = st.columns(2)
with col1:
    budget_file = st.file_uploader("📥 1. Planilha Orçamentária Base (.xlsx)", type=["xlsx"])
with col2:
    ref_file = st.file_uploader("📥 2. Referência SINAPI Orçafascio (.xlsx)", type=["xlsx"])

if budget_file and ref_file:
    # Retorna o ponteiro do arquivo para o início antes de ler
    budget_file.seek(0)
    wb = openpyxl.load_workbook(budget_file)
    
    # Abre uma SEGUNDA cópia apenas para ler os resultados das fórmulas
    budget_file.seek(0)
    wb_data_only = openpyxl.load_workbook(budget_file, data_only=True)
    
    sheet_names = wb.sheetnames
    selected_sheet = st.selectbox("📌 Selecione qual aba contém o Orçamento a ser atualizado:", sheet_names)

    if st.button("🚀 Processar Atualização"):
        with st.spinner("Analisando e atualizando valores..."):
            try:
                ref_df_raw = pd.read_excel(ref_file)
                ref_header_idx = ref_df_raw[ref_df_raw.apply(lambda r: r.astype(str).str.contains("Código", case=False, na=False).any(), axis=1)].index[0]
                
                ref_df = pd.read_excel(ref_file, header=ref_header_idx + 1)
                
                col_codigo_ref = ref_df.columns[1]
                col_valor_ref = ref_df.columns[8]
                
                ref_df_clean = ref_df.dropna(subset=[col_codigo_ref, col_valor_ref])
                precos_dict = {}
                for k, v in zip(ref_df_clean[col_codigo_ref], ref_df_clean[col_valor_ref]):
                    chave = str(k).strip()
                    if chave.upper() not in ["CÓDIGO", "CODIGO", "NAN", ""]:
                        precos_dict[chave] = v

                # Planilha de edição (onde vamos salvar)
                ws = wb[selected_sheet]
                # Planilha de leitura (onde vamos pegar os valores das fórmulas)
                ws_data = wb_data_only[selected_sheet]

                header_row = None
                codigo_col = None
                custo_unit_col = None
                bdi_col = None
                preco_unit_col = None
                quantidade_col = None
                preco_total_col = None

                for row in range(1, 30): 
                    for col in range(1, 15): 
                        cell_val = str(ws.cell(row=row, column=col).value).strip().upper()
                        if cell_val in ["CÓDIGO", "CODIGO"]:
                            header_row = row
                            codigo_col = col
                            break
                    if header_row:
                        break

                if not header_row or not codigo_col:
                    st.error(f"Não foi possível encontrar a coluna 'CÓDIGO' na aba '{selected_sheet}'.")
                    st.stop()
                
                for col in range(1, 20):
                    cell_val = str(ws.cell(row=header_row, column=col).value).strip().upper()
                    if not cell_val or cell_val == "NONE": continue
                    
                    if cell_val in ["CUSTO UNITÁRIO (SEM BDI)", "CUSTO UNITÁRIO", "CUSTO UNIT.", "VALOR UNIT", "VALOR UNITÁRIO", "VALOR UNITÁRIO (R$)"]:
                        custo_unit_col = col
                    elif cell_val == "BDI":
                        bdi_col = col
                    elif cell_val in ["PREÇO UNITÁRIO (COM BDI)", "VALOR UNIT COM BDI", "VALOR COM BDI"]:
                        preco_unit_col = col
                    elif cell_val in ["QUANTIDADE", "QUANT.", "QTD"]:
                        quantidade_col = col
                    elif cell_val in ["PREÇO TOTAL", "TOTAL", "VALOR TOTAL"]:
                        preco_total_col = col

                if not custo_unit_col:
                    st.error(f"Não foi possível encontrar a coluna de Valor Unitário na aba '{selected_sheet}'.")
                    st.stop()

                itens_atualizados = 0
                log_alteracoes = [] 

                for row in range(header_row + 1, ws.max_row + 1):
                    cod_cell = ws.cell(row=row, column=codigo_col)
                    cod_val = str(cod_cell.value).strip() if cod_cell.value else None
                    
                    if not cod_val or cod_val.upper() in ["CÓDIGO", "CODIGO", "NONE"]:
                        continue

                    if cod_val in precos_dict:
                        novo_valor = precos_dict[cod_val]
                        valor_antigo = 0.0
                        
                        if custo_unit_col:
                            # AQUI ESTÁ O TRUQUE: Lemos o valor da planilha "data_only"
                            celula_antiga = ws_data.cell(row=row, column=custo_unit_col).value
                            if celula_antiga is not None:
                                try:
                                    valor_antigo = float(celula_antiga)
                                except:
                                    valor_antigo = celula_antiga 
                            
                            # E gravamos o valor novo na planilha oficial
                            ws.cell(row=row, column=custo_unit_col).value = novo_valor
                        
                        bdi_val = 0
                        if bdi_col:
                            # Lê o BDI da planilha data_only também, caso seja uma fórmula
                            bdi_raw = ws_data.cell(row=row, column=bdi_col).value
                            if bdi_raw is not None:
                                try:
                                    bdi_val = float(bdi_raw) 
                                except:
                                    pass
                        
                        try:
                            novo_preco_com_bdi = float(novo_valor) * (1 + bdi_val)
                        except:
                            novo_preco_com_bdi = novo_valor
                        
                        if preco_unit_col:
                            ws.cell(row=row, column=preco_unit_col).value = novo_preco_com_bdi
                            
                        if quantidade_col and preco_total_col:
                            # Lê a quantidade da planilha data_only, caso seja fórmula
                            qtd = ws_data.cell(row=row, column=quantidade_col).value
                            if qtd is not None:
                                try:
                                    qtd_float = float(qtd)
                                    ws.cell(row=row, column=preco_total_col).value = qtd_float * float(novo_preco_com_bdi)
                                except:
                                    pass
                        
                        log_alteracoes.append({
                            "Linha Excel": row,
                            "Código": cod_val,
                            "Valor Antigo (R$)": valor_antigo,
                            "Novo Valor (R$)": novo_valor
                        })
                        
                        itens_atualizados += 1

                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                st.success(f"✅ Atualização concluída! {itens_atualizados} itens foram atualizados na aba '{selected_sheet}'.")
                
                if log_alteracoes:
                    st.write("### 📊 Relatório de Conferência")
                    st.write("Verifique abaixo os valores antigos e os novos valores aplicados:")
                    df_log = pd.DataFrame(log_alteracoes)
                    
                    df_log["Valor Antigo (R$)"] = pd.to_numeric(df_log["Valor Antigo (R$)"], errors='coerce').fillna(0)
                    df_log["Novo Valor (R$)"] = pd.to_numeric(df_log["Novo Valor (R$)"], errors='coerce').fillna(0)
                    
                    st.dataframe(
                        df_log.style.format({
                            "Valor Antigo (R$)": "{:.2f}",
                            "Novo Valor (R$)": "{:.2f}"
                        }),
                        use_container_width=True
                    )
                
                st.download_button(
                    label="⬇️ Baixar Planilha Atualizada",
                    data=output,
                    file_name=f"Orcamento_Atualizado_{selected_sheet}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Ocorreu um erro durante o processamento: {e}")
