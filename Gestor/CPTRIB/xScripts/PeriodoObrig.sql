INSERT INTO TAB_PERIODO_OBRIGACAO(TPO_INSCRICAO,TPO_PERIODO_INICIAL)
SELECT TCI_IM, TCI_INICIO_ATIVIDADE FROM TAB_CONTRIBUINTE
WHERE TCI_TAE_CAE <> 0