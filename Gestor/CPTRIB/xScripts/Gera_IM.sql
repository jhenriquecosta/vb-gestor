
DECLARE @IM VARCHAR(12)
DECLARE @IM_ANT VARCHAR(12)
DECLARE @CONTADOR INT
DECLARE @AUX INT
DECLARE CRS_IM CURSOR FOR
	select TCI_INSCRICAO_ANTERIOR from tab_contribuinte
	where TCI_Inscricao_anterior IS not NULL
OPEN CRS_IM 
FETCH NEXT FROM CRS_IM 
INTO @IM_ANT
SET @CONTADOR =  2
WHILE (@@FETCH_STATUS = 0)
BEGIN	
	SET @CONTADOR = @CONTADOR + 1
	SET @AUX = 11000000  + @CONTADOR
	SET @IM = CAST(@AUX AS VARCHAR(8)) + '-' + 
		substring(cast(cast(SUBSTRING(@IM_ANT,5,2) as int) + @contador + (@contador * cast(SUBSTRING(@IM_ANT,5,2) as int)+36)/2  as varchar),1,2)
--	PRINT @IM
	UPDATE TAB_CONTRIBUINTE SET TCI_IM = @IM WHERE TCI_INSCRICAO_ANTERIOR = @IM_ANT
	FETCH NEXT FROM CRS_IM 
	INTO @IM_ANT
END
CLOSE CRS_IM
DEALLOCATE CRS_IM