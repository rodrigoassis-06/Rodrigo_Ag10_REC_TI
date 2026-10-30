import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import os

def criar_roteiro():
    doc = docx.Document()

    for sec in doc.sections:
        sec.top_margin = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin = Inches(1)
        sec.right_margin = Inches(1)

    # Cabecalho
    p_head = doc.add_paragraph()
    r_head = p_head.add_run(
        "CENTRO PAULA SOUZA - EAD TEC\n"
        "CURSO TECNICO EM DESENVOLVIMENTO DE SISTEMAS\n"
        "AGENDA 10 - TECNOLOGIAS DA INFORMACAO I (RECUPERACAO)"
    )
    r_head.bold = True
    r_head.font.size = Pt(11)
    r_head.font.color.rgb = RGBColor(60, 60, 60)
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_div = doc.add_paragraph()
    r_div = p_div.add_run("-" * 70)
    r_div.font.color.rgb = RGBColor(180, 180, 180)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Titulo
    p_title = doc.add_paragraph()
    r_title = p_title.add_run("ROTEIRO EXPLICATIVO: IMPLANTACAO DO SCRUM NO APP 'ORGANIZA+'")
    r_title.bold = True
    r_title.font.size = Pt(14)
    r_title.font.color.rgb = RGBColor(37, 99, 235)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    p_id = doc.add_paragraph()
    r_id = p_id.add_run("Aluno: Rodrigo Assis | Formato: Rodrigo_Ag10_REC_TI")
    r_id.bold = True
    r_id.font.size = Pt(10)
    p_id.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # 1. Analise do Cenario e Desafios
    doc.add_heading("1. Analise do Cenario e Principais Desafios", level=1)
    doc.add_paragraph(
        "Na startup responsavel pelo aplicativo 'Organiza+', a equipe de 5 desenvolvedores enfrenta tres desafios criticos:\n\n"
        "1. Mudancas Constantes de Requisitos: O cliente solicita recorrentemente novas funcionalidades "
        "(calendario interativo, integracao com redes sociais e personalizacao), gerando risco de retrabalho constante.\n\n"
        "2. Prazos Curtos e Entregas Frequentes: Exigencia de disponibilizar versoes funcionais a cada duas semanas.\n\n"
        "3. Equipe Heterogenea: Cinco profissionais com niveis desiguais de maturidade e vivencia em metodologias ageis."
    )

    # 2. Elementos do Scrum Aplicados
    doc.add_heading("2. Elementos do Scrum Aplicados a Solucao", level=1)
    doc.add_paragraph(
        "Para responder com eficacia a esse cenario, estruturaram-se os elementos do framework Scrum da seguinte forma:"
    )

    table = doc.add_table(rows=1, cols=3)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr = table.rows[0].cells
    hdr[0].text = "Categoria"
    hdr[1].text = "Elemento Scrum"
    hdr[2].text = "Aplicacao Pratica no Organiza+"
    for cell in hdr:
        cell.paragraphs[0].runs[0].bold = True

    dados = [
        ("Papeis", "Product Owner (PO)", "Centraliza e prioriza as solicitacoes do cliente no Product Backlog."),
        ("Papeis", "Scrum Master", "Facilita os ritos, treina os 5 desenvolvedores e elimina bloqueios."),
        ("Papeis", "Developers", "Equipe multidisciplinar auto-organizavel focada na construcao do incremento."),
        ("Eventos", "Sprint (2 semanas)", "Ciclo time-box fixo correspondente a demanda de entregas quinzenais."),
        ("Eventos", "Sprint Planning", "Planejamento inicial de 4h definindo meta e itens do Sprint Backlog."),
        ("Eventos", "Daily Scrum", "Alinhamento diario de 15 minutos em pe para detectar impedimentos."),
        ("Eventos", "Sprint Review", "Demonstracao da versao funcional ao cliente no encerramento da Sprint."),
        ("Eventos", "Sprint Retrospective", "Reuniao interna de licoes aprendidas e melhoria continua do time."),
        ("Artefatos", "Product Backlog", "Lista unica, dinamica e ordenada de todas as funcionalidades desejadas."),
        ("Artefatos", "Sprint Backlog", "Conjunto de tarefas selecionadas e congeladas para a Sprint vigente."),
        ("Artefatos", "Incremento (DoD)", "Versao funcional testada segundo a Definicao de Pronto pre-acordada.")
    ]

    for cat, elem, apl in dados:
        row = table.add_row().cells
        row[0].text = cat
        row[1].text = elem
        row[2].text = apl

    doc.add_paragraph()

    # 3. Gestao de Mudancas
    doc.add_heading("3. Estrategia para Gestao de Mudancas Contínuas", level=1)
    doc.add_paragraph(
        "A principal dor relatada (mudancas frequentes solicitadas pelo cliente) e solucionada pela 'Regra de Ouro' do Scrum:\n"
        "O escopo de uma Sprint ativa nao pode ser alterado durante seus 14 dias de execucao. "
        "Quando o cliente trouxer novas ideias (ex: redes sociais ou calendario), o Product Owner registra a solicitacao no "
        "Product Backlog, analisa o valor de negocio e a prioriza juntamente com o cliente. "
        "Dessa forma, o time mantem foco e estabilidade na Sprint atual, e a nova funcionalidade podera ser puxada na Sprint Planning seguinte."
    )

    # 4. Dialogo com a IA e Infografico
    doc.add_heading("4. Evidencias da Co-criacao com Inteligencia Artificial", level=1)
    doc.add_paragraph(
        "A estruturacao deste roteiro e do infografico contou com apoio da Inteligencia Artificial, que atuou como especialista Scrum "
        "auxiliando na definicao do time-box ideal, protecao do fluxo de trabalho e nivelamento da equipe heterogenea."
    )

    base_dir = os.path.dirname(os.path.abspath(__file__))
    img_ia = os.path.join(base_dir, "interacao_ia.png")
    if os.path.exists(img_ia):
        doc.add_paragraph().add_run("Registro do Dialogo com a IA:").bold = True
        doc.add_picture(img_ia, width=Inches(6.0))
        p_cap1 = doc.add_paragraph("Figura 1: Dialogo com a IA sobre o framework Scrum e resolucao dos desafios do projeto.")
        p_cap1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap1.runs[0].font.size = Pt(9)
        p_cap1.runs[0].italic = True

    doc.add_paragraph()

    img_info = os.path.join(base_dir, "infografico_scrum.png")
    if os.path.exists(img_info):
        doc.add_heading("5. Infografico Síntese do Projeto", level=1)
        doc.add_picture(img_info, width=Inches(6.0))
        p_cap2 = doc.add_paragraph("Figura 2: Infografico de implantacao do Scrum para o aplicativo Organiza+.")
        p_cap2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap2.runs[0].font.size = Pt(9)
        p_cap2.runs[0].italic = True

    doc.add_paragraph()

    # 6. Conclusao
    doc.add_heading("6. Conclusao", level=1)
    doc.add_paragraph(
        "A adocao do Scrum na startup soluciona de forma estruturada as tensoes entre flexibilidade de escopo e previsibilidade de entrega. "
        "O cliente passa a receber versoes funcionais testadas quinzenalmente na Sprint Review, enquanto os desenvolvedores trabalham protegidos "
        "contra mudancas abruptas no meio do ciclo produtivo, promovendo maturidade tecnica e previsibilidade institucional."
    )

    out_docx = os.path.join(base_dir, "Rodrigo_Ag10_REC_TI.docx")
    doc.save(out_docx)
    print("Documento salvo em:", out_docx)

if __name__ == "__main__":
    criar_roteiro()
