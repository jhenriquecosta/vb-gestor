CREATE VIEW dbo.VIS_STATUS_OBRIGACAO
AS
SELECT     TGE_CODIGO, TGE_NOME
FROM         dbo.TAB_GERAL
WHERE     (TGE_CODIGO <> 0) AND (TGE_TIPO =
                          (SELECT     TGE_TIPO
                            FROM          TAB_GERAL
                            WHERE      TGE_NOME = 'STATUS OBRIGACAO'))

