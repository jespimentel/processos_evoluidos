# arquivos_repetidos

Na Promotoria, usamos uma pasta do SharePoint para manter as peças em elaboração.

Quando a peça se torna definitiva e é incorporada ao processo, ela é copiada para o OneDrive do Promotor e fica disponível para consulta de toda a equipe.

O SharePoint exporta a relação de arquivos em formato Excel.

O código 'relaciona_duplicados.py' identifica os arquivos do SharePoint que podem ser deletados.

O código 'relaciona_elegiveis_delecao.py' identifica os arquivos do SharePoint que possivelmente podem ser deletados, pois se relacionam a processos já evoluídos (com alegações finais, contrarrazões de recurso ou razões de recurso).
