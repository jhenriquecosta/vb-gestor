create view vis_taxas
as
select tip_sigla_imposto + ' - '+ tip_nome_imposto as Imposto,
tpi_valor_taxa_fixa  as Valor,tpi_ano_imposto
from tab_imposto,tab_parametro_imposto as ano
where tip_cod_imposto = tpi_tip_cod_imposto
and left(tip_sigla_imposto,2) like '%TX%' 

select * from tab_parametro_imposto
select * from tab_imposto