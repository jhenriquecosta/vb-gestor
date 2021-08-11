create view VIS_ARREC_ATIVIDADE AS select min(tae_cae) as CAE,min(tae_nome) AS ATIVIDADE,sum(tdr_valor_real_pago) AS VALOR from tab_darm_recebido,tab_contribuinte,tab_atividade_economica
where tdr_im = tci_im and tci_tae_cae = tae_cae
group by tae_nome
