# ... (o resto do código acima continua igual)

    # --- 4. EXIBIÇÃO E CONFIRMAÇÃO ---
    if len(dados_para_enviar) > 0:
        df = pd.DataFrame(dados_para_enviar, columns=["Data", "Arquivo", "Técnico", "Horas"])
        st.success(f"Encontrei {len(dados_para_enviar)} registros!")
        st.dataframe(df)
        
        if st.button("Confirmar e Gravar TUDO no Sheets"):
            with st.spinner("Enviando dados para a aba 'Comissoes'..."):
                try:
                    client = conectar_sheets()
                    
                    # 1. Abre o ARQUIVO (Planilha) pelo nome
                    arquivo_sheet = client.open("Dados_HTML")
                    
                    # 2. Seleciona a ABA específica pelo nome "Comissoes"
                    # Se der erro aqui, é porque a aba não tem exatamente esse nome
                    aba = arquivo_sheet.worksheet("Comissoes") 
                    
                    # 3. Envia os dados
                    aba.append_rows(dados_para_enviar)
                    
                    st.balloons()
                    st.success("✅ Sucesso! Dados gravados na aba 'Comissoes'.")
                    
                except Exception as e:
                    # Tratamento para o falso erro 200
                    if "200" in str(e):
                         st.balloons()
                         st.success("✅ Sucesso! (O Google confirmou o recebimento).")
                    else:
                        st.error(f"Erro ao salvar: {e}")
                        st.warning("Verifique se a aba da planilha se chama exatamente 'Comissoes' e se o robô tem acesso de Editor.")
