from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado
from docx.document import Document 

def gerar_secao_objetivo(doc: Document, row=None):
    """
    Gera a seção de Objetivo do Relatório, usando f-string para maior clareza 
    e preparar para a extração de dados dinâmicos da 'row' (caso seja implementada).
    """
    num_monitoramento = "Xº" 
    ctr_original = "xx/xxxx"
    doc_sei = "xxxxxxx"
    cidades_str = "xxxxxxxx, xxxxxxxxx, e xxxxxxxxx" 

    if row is not None:
        num_monitoramento = str(row.get("Num Monitoramento", num_monitoramento))
        ctr_original = str(row.get("CTR Original", ctr_original))
        doc_sei = str(row.get("Doc SEI Original", doc_sei))
    
    adicionar_titulo_secao(doc, "2. OBJETIVO")
    
    texto_objetivo = (
        f"Este Relatório do {num_monitoramento} Monitoramento objetiva apresentar os resultados das "
        f"vistorias acerca das Não Conformidades pendentes registradas no Relatório de "
        f"Fiscalização Técnico-Operacional nº {ctr_original} (Doc. SEI nº {doc_sei}), referentes "
        f"aos Terminais Rodoviários Intermunicipais dos municípios de {cidades_str}, "
        "conforme o cronograma de atividades encaminhado pela Concessionária SOCICAM."
    )
    
    adicionar_paragrafo_justificado(doc, texto_objetivo)

   
    
   