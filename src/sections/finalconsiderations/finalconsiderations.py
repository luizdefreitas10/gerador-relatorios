from utils import (
    adicionar_titulo_secao,
    adicionar_paragrafo_justificado,
    adicionar_texto_centralizado,
    adicionar_texto_esquerda,
)
from datetime import datetime
from docx.document import Document 

def _adicionar_assinatura_bloco(doc: Document, nome, cargo, matricula, negrito_nome=True):
    """
    Auxiliar para adicionar um bloco de assinatura formatado e centrado.
    """
  
    if not nome or "xxxxxx" in nome.lower():
        return

    adicionar_texto_centralizado(doc, nome, negrito=negrito_nome)
    adicionar_texto_centralizado(doc, cargo, negrito=False)
    
    if matricula: 
        if "matrícula" not in matricula.lower() and "nº" not in matricula.lower():
            texto_matricula = f"Matrícula: {matricula}"
        else:
            texto_matricula = matricula
            
        adicionar_texto_centralizado(doc, texto_matricula, negrito=False)
    
    doc.add_paragraph() 

def gerar_secao_consideracoes_finais(doc: Document, row):
    """
    Gera a seção '6. CONCLUSÃO' do relatório.
    """

    adicionar_titulo_secao(doc, "6. CONCLUSÃO")

    # Extração de Placeholders para f-string
    num_monitoramento = str(row.get("Num Monitoramento", "Xº"))
    ctr_original = str(row.get("CTR Original", "xx/xxxx"))
    periodo_vistoria = str(row.get("Periodo Vistoria Texto", "xx a xx de MÊS de ANO"))
    
    texto1 = (
        f"Diante das constatações apontadas neste {num_monitoramento} Relatório de Monitoramento do Relatório de Fiscalização "
        f"Técnico-Operacional CTR {ctr_original}, referente ás vistorias técnicas realizadas no período de {periodo_vistoria}, "
        f"solicitamos seu envio para a SOCICAM para que sejam informados das Não Conformidades."
    )

    adicionar_paragrafo_justificado(doc, texto1)

    # Pega a data atual no formato dd/mm/aaaa
    data_atual = datetime.now().strftime("%d/%m/%Y")
    
    adicionar_texto_centralizado(doc, f"Recife, {data_atual}.", negrito=False, tamanho_fonte=11)
    
    doc.add_paragraph() 

    assinantes = str(row.get("Assinatura", "")).split(";")
    
    matricula_analista_fixa = "Matricula: nºxxxxxx/xx" 
    
    for assinante in assinantes:
        nome = assinante.strip()
        
        _adicionar_assinatura_bloco(
            doc, 
            nome, 
            "Analista de Regulação", 
            matricula_analista_fixa,
            negrito_nome=True
        )

    adicionar_texto_esquerda(doc, "Ciente.", negrito=False)
    
    coordenador = str(row.get("Coordenador", "")).strip()
  
    matricula_coordenador_fixa = "Matricula: nº209640/01"
    
    _adicionar_assinatura_bloco(
        doc, 
        coordenador, 
        "Coordenadora de Transportes e Rodovias", 
        matricula_coordenador_fixa,
        negrito_nome=True
    )