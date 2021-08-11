begin tran
insert into tab_contribuinte select * from vttrib_pho..tab_contribuinte 
commit tran

begin tran
insert into tab_imovel select * from vttrib_pho..tab_imovel
commit tran

begin tran
insert into tab_detalhe_imovel select * from vttrib_pho..tab_detalhe_imovel
commit tran

begin tran
insert into tab_geracao_tributo select * from vttrib_pho..tab_geracao_tributo
commit tran

begin tran
insert into tab_bairro select * from vttrib_pho..tab_bairro
commit tran

begin tran
insert into tab_logradouro select * from vttrib_pho..tab_logradouro
commit tran

begin tran
insert into tab_componente  select * from vttrib_pho..tab_componente
commit tran

