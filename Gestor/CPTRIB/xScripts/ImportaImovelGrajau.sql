CREATE PROCEDURE sp_ImportaImovel as 
DELETE TAB_IMOVEL
INSERT INTO TAB_IMOVEL(tim_ic,TIM_TCI_IM,TIM_TTL_COD_TIP_LOGR,tim_ic_anterior,tim_unidade,TIM_TCI_INSCRICAO_ANTERIOR,tim_tlg_cod_logradouro,
TIM_TBA_COD_BAIRRO,tim_numero,tim_complemento,TIM_DATA_CADASTRO,TIM_CEP,TIM_SECAO,TIM_QUADRA,TIM_LOTE,TIM_VALOR_TERRENO,TIM_VALOR_EDIFIC,
TIM_TUS_USUARIO,TIM_TUS_COD_USUARIO)

select imovel_inscricao,TCI_IM,tlg_ttl_cod_tip_logr,imovel_inscricao_ant,substring(imovel_inscricao,13,2),imovel_prop_codigo,
imovel_loc_cod_logra,imovel_loc_cod_bairro,imovel_loc_num_logra,imovel_loc_compl_logra,
getdate(),'65940000',imovel_loc_secao_logra,imovel_loc_quadra,imovel_loc_lote,
imovel_vvt,imovel_vve,'MIGRACAO','MIGRACAO' from import..imoveis$,TAB_CONTRIBUINTE,TAB_LOGRADOURO
where imovel_prop_codigo  = TCI_INSCRICAO_ANTERIOR AND IMOVEL_LOC_COD_LOGRA = tlg_COD_LOGRADOURO


---***********************************************PROCEDURE ATUALIZA VALOR VENAL
DECLARE @ValTerr REAL,@ValEdif REAL, @Cadastro VARCHAR(20)
DECLARE CRS_VENAL CURSOR FOR
	SELECT VVT, VVP, CADASTRO 
	FROM arript WHERE ANOCALC = 2004 AND PARCELA = 0
	ORDER BY CADASTRO 
OPEN CRS_VENAL
FETCH NEXT FROM CRS_VENAL
INTO @ValTerr , @ValEdif , @Cadastro
WHILE (@@FETCH_STATUS = 0)
BEGIN
	PRINT 'ATUALIZANDO VALOR VENAL PARA '+@Cadastro+', AGUARDE...'
	
	UPDATE TAB_IMOVEL SET tim_valor_terreno = @ValTerr, tim_valor_edific = @ValEdif,TIM_VALOR = @ValTerr + @ValEdif
	where tim_IC_anterior = @Cadastro
	
	FETCH NEXT FROM CRS_VENAL
	INTO @ValTerr , @ValEdif , @Cadastro
END
CLOSE CRS_VENAL
DEALLOCATE CRS_VENAL


GO
