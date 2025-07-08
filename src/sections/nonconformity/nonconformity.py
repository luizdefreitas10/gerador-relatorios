from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
import os
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
    adicionar_texto_centralizado,
)


def gerar_secao_nao_conformidades_constatadas(
    doc, row, nao_conformidades_df, fotos_dir
):
    """
    Gera a seção '3. NÃO CONFORMIDADES CONSTATADAS' com base na planilha.

    Parâmetros:
    - doc: objeto Document.
    - row: linha da fiscalização (Series do DataFrame fiscalizacoes_df).
    - nao_conformidades_df: DataFrame da aba 'Não-conformidades'.
    - fotos_dir: caminho da pasta onde estão as imagens.
    """
    adicionar_titulo_secao(doc, "3. NÃO CONFORMIDADES CONSTATADAS")

    adicionar_paragrafo_justificado(
        doc,
        "A seguir, apresentam-se as não conformidades registradas durante a vistoria técnica:",
    )

    id_fisc = row["ID da Fiscalização"]

    nc_fisc = nao_conformidades_df[
        nao_conformidades_df["ID da Fiscalização"] == id_fisc
    ]

    for nc_id, grupo_nc in nc_fisc.groupby("Nº"):
        descricao = grupo_nc["Não Conformidade"].iloc[0]
        doc.add_heading(f"{nc_id} - {descricao}", level=1)

        for _, linha in grupo_nc.iterrows():
            nome_foto = linha["Foto"]
            legenda = linha["Legenda da Foto"]
            foto_path = os.path.join(fotos_dir, str(nome_foto))

            if os.path.exists(foto_path):
                doc.add_picture(foto_path, width=Inches(3))
                doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                adicionar_texto_centralizado(doc, legenda)
            else:
                doc.add_paragraph(f"⚠️ Foto não encontrada: {nome_foto}")
