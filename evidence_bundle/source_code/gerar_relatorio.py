from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_TAB_ALIGNMENT
from pathlib import Path

OUT=Path('/home/ubuntu/work_report')
doc=Document()
sec=doc.sections[0]
sec.page_width=Cm(21); sec.page_height=Cm(29.7)
sec.top_margin=Cm(3); sec.left_margin=Cm(3); sec.right_margin=Cm(2); sec.bottom_margin=Cm(2)
sec.header_distance=Cm(1.25); sec.footer_distance=Cm(1.25)

styles=doc.styles
styles['Normal'].font.name='Arial'; styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial'); styles['Normal'].font.size=Pt(12)
styles['Normal'].paragraph_format.line_spacing=1.5; styles['Normal'].paragraph_format.space_after=Pt(0); styles['Normal'].paragraph_format.first_line_indent=Cm(1.25)
for name,size,bold in [('Heading 1',12,True),('Heading 2',12,True),('Heading 3',12,True)]:
    s=styles[name]; s.font.name='Arial'; s._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial'); s.font.size=Pt(size); s.font.bold=bold; s.font.color.rgb=RGBColor(0,0,0)
    s.paragraph_format.space_before=Pt(18 if name=='Heading 1' else 12); s.paragraph_format.space_after=Pt(6); s.paragraph_format.line_spacing=1.5; s.paragraph_format.keep_with_next=True

def set_cell_shading(cell, fill='D9D9D9'):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill); tcPr.append(shd)
def set_cell_text(cell,text,bold=False,size=10):
    cell.text=''; p=cell.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after=Pt(0); p.paragraph_format.line_spacing=1.0
    r=p.add_run(str(text)); r.font.name='Arial'; r._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial'); r.font.size=Pt(size); r.bold=bold
    cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
def add_page_number(paragraph):
    paragraph.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=paragraph.add_run(); fld=OxmlElement('w:fldSimple'); fld.set(qn('w:instr'),'PAGE'); r._r.addnext(fld)
def header_footer(section):
    hp=section.header.paragraphs[0]; hp.text=''; hp.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    r=hp.add_run('ASTEROID IMPACT INTELLIGENCE'); r.font.name='Arial'; r.font.size=Pt(9); r.font.bold=True
    fp=section.footer.paragraphs[0]; fp.text=''; add_page_number(fp)
    for r in fp.runs: r.font.name='Arial'; r.font.size=Pt(10)
header_footer(sec)

def p(text='', align=None, bold=False, italic=False, indent=True, space_after=0):
    para=doc.add_paragraph(); para.paragraph_format.line_spacing=1.5; para.paragraph_format.space_after=Pt(space_after)
    para.paragraph_format.first_line_indent=Cm(1.25) if indent else Cm(0)
    if align is not None: para.alignment=align
    run=para.add_run(text); run.font.name='Arial'; run._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial'); run.font.size=Pt(12); run.bold=bold; run.italic=italic
    return para

def heading(text, level=1):
    h=doc.add_heading(text, level=level); h.paragraph_format.first_line_indent=Cm(0); return h

def caption(text):
    para=p(text, align=WD_ALIGN_PARAGRAPH.LEFT, bold=True, indent=False); para.paragraph_format.space_before=Pt(6); para.paragraph_format.space_after=Pt(4); return para

def source(text):
    para=p('Fonte: '+text, align=WD_ALIGN_PARAGRAPH.LEFT, italic=True, indent=False); para.paragraph_format.space_after=Pt(8); 
    for r in para.runs: r.font.size=Pt(10)

def table(headers, rows, widths=None):
    t=doc.add_table(rows=1, cols=len(headers)); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
    for i,h in enumerate(headers): set_cell_text(t.rows[0].cells[i],h,True,10); set_cell_shading(t.rows[0].cells[i],'D9D9D9')
    for row in rows:
        cells=t.add_row().cells
        for i,val in enumerate(row): set_cell_text(cells[i],val,False,9)
    if widths:
        for row in t.rows:
            for i,w in enumerate(widths): row.cells[i].width=Cm(w)
    return t

def pagebreak(): doc.add_page_break()

# Capa
for _ in range(3): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('FORD ENTER & SENAI CIMATEC', align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, indent=False)
p('BRUNO SOARES FERREIRA RAMOS', align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, indent=False)
for _ in range(5): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('ASTEROID IMPACT INTELLIGENCE', align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, indent=False)
p('Solução de Machine Learning em nuvem para classificação de asteroides potencialmente perigosos', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
for _ in range(7): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('SALVADOR', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('2026', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
pagebreak()

# Folha de rosto
for _ in range(5): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('BRUNO SOARES FERREIRA RAMOS', align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, indent=False)
for _ in range(4): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('ASTEROID IMPACT INTELLIGENCE', align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, indent=False)
p('Solução de Machine Learning em nuvem para classificação de asteroides potencialmente perigosos', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
for _ in range(3): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
para=p('Relatório técnico apresentado à Ford Enter e ao SENAI CIMATEC como requisito do Desafio 04 — Serviços de Computação em Nuvem para Machine Learning.', align=WD_ALIGN_PARAGRAPH.JUSTIFY, indent=False); para.paragraph_format.left_indent=Cm(8); para.paragraph_format.right_indent=Cm(0)
for _ in range(6): p('', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
p('SALVADOR', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False); p('2026', align=WD_ALIGN_PARAGRAPH.CENTER, indent=False)
pagebreak()

# Resumo
heading('RESUMO',1)
p('Este relatório apresenta o desenvolvimento de uma solução de Machine Learning em ambiente de computação em nuvem para classificação de asteroides próximos da Terra como potencialmente perigosos (Potentially Hazardous Asteroids — PHA). A solução utiliza dados oficiais da NASA/JPL Small-Body Database (SBDB), armazenados em bucket versionado, preparados por meio de uma rotina reprodutível e utilizados no treinamento e na avaliação de três algoritmos de classificação. O modelo selecionado foi o Gradient Boosting por histogramas, registrado como versão champion e disponibilizado por uma API Flask com os endpoints /health e /predict. A avaliação final alcançou acurácia de 0,7607, recall de 0,7412, F1 de 0,2689 e ROC AUC de 0,8415 para a classe PHA. Foram realizados testes funcionais, teste de carga com 300 requisições e oito clientes concorrentes, monitoramento estruturado, verificação de data drift pelo índice PSI e análise de queda de desempenho. O experimento foi executado no modo padrão de bucket local emulado, mantendo a mesma interface de armazenamento prevista para uma migração ao Google Cloud Storage. O sistema não estima colisão, probabilidade de impacto ou dano; classifica o rótulo oficial PHA a partir de elementos orbitais.', indent=False)
p('Palavras-chave: computação em nuvem. Machine Learning. MLOps. asteroides próximos da Terra. monitoramento de modelos.', indent=False)
pagebreak()

heading('LISTA DE FIGURAS',1)
p('Figura 1 — Arquitetura da solução Asteroid Impact Intelligence', indent=False)
heading('LISTA DE TABELAS',1)
p('Tabela 1 — Serviços e funções da solução', indent=False)
p('Tabela 2 — Distribuição do conjunto de dados', indent=False)
p('Tabela 3 — Comparação dos experimentos', indent=False)
p('Tabela 4 — Métricas finais no conjunto de teste', indent=False)
heading('SUMÁRIO',1)
for item in ['1 INTRODUÇÃO','2 OBJETIVOS','3 FUNDAMENTAÇÃO E ESCOPO DA SOLUÇÃO','4 ARQUITETURA E SERVIÇOS UTILIZADOS','5 DESENVOLVIMENTO DA SOLUÇÃO','6 RESULTADOS E AVALIAÇÃO','7 MONITORAMENTO, GOVERNANÇA E MLOPS','8 OPORTUNIDADES DE MELHORIA','9 CONSIDERAÇÕES FINAIS','REFERÊNCIAS']:
    p(item, indent=False)
pagebreak()

heading('1 INTRODUÇÃO',1)
p('A computação em nuvem permite disponibilizar recursos de armazenamento, processamento e serviços de inferência sob demanda, favorecendo a construção de aplicações de Inteligência Artificial com maior elasticidade, rastreabilidade e disponibilidade. No ciclo de vida de Machine Learning, a nuvem também facilita a organização dos dados, o registro de experimentos, o versionamento de modelos e o monitoramento contínuo da solução.')
p('Este trabalho desenvolve uma aplicação educacional denominada Asteroid Impact Intelligence. O projeto consulta o catálogo NASA/JPL Small-Body Database por meio da SBDB Query API, seleciona asteroides próximos da Terra e treina classificadores para reproduzir, a partir de elementos orbitais, o rótulo oficial PHA. A fonte NASA/JPL é utilizada como referência de dados e de documentação da API (JET PROPULSION LABORATORY, 2021).')
p('É importante delimitar o problema: o modelo não prevê colisão com a Terra, não calcula probabilidade de impacto e não estima danos. Seu objetivo é exclusivamente classificar o rótulo oficial “Potentially Hazardous Asteroid” fornecido pelo catálogo, evitando o uso de variáveis que definem diretamente esse rótulo e poderiam causar vazamento de informação.')
heading('1.1 Problema e justificativa',2)
p('A identificação automatizada de padrões em catálogos astronômicos é um caso de uso adequado para demonstrar um fluxo completo de Machine Learning em nuvem. A solução permite observar, em um único projeto, as etapas de coleta, armazenamento, preparação, treinamento, registro, inferência, monitoramento e governança, conforme proposto no desafio acadêmico.')

heading('2 OBJETIVOS',1)
heading('2.1 Objetivo geral',2)
p('Projetar e documentar uma solução de Machine Learning em ambiente de computação em nuvem para classificar asteroides próximos da Terra como PHA ou não PHA, integrando armazenamento versionado, treinamento, API de inferência, monitoramento e evidências de execução.')
heading('2.2 Objetivos específicos',2)
for txt in ['Coletar dados oficiais da NASA/JPL SBDB com paginação, retentativas e registro do instante de coleta;','Armazenar dados brutos, dados processados, modelos, logs e relatórios em estrutura versionada;','Preparar os dados sem utilizar MOID e H, evitando vazamento do rótulo oficial;','Comparar regressão logística, Random Forest e Gradient Boosting por validação cruzada estratificada;','Registrar as versões v1, v2 e champion com metadados, métricas e hashes SHA-256;','Disponibilizar uma API de inferência com validação de entrada e logs estruturados;','Avaliar carga, latência, erros, data drift, model drift e necessidade de retreinamento.']:
    para=p('• '+txt, indent=False); para.paragraph_format.left_indent=Cm(0.5)

heading('3 FUNDAMENTAÇÃO E ESCOPO DA SOLUÇÃO',1)
p('O fluxo de trabalho foi estruturado segundo práticas de MLOps: separação entre dados brutos e processados, reprodutibilidade por semente aleatória, validação cruzada, registro de modelos, monitoramento de inferência e geração de manifesto de evidências. Em uma implantação cloud nativa, os componentes locais emulados podem ser substituídos por serviços gerenciados equivalentes, como Google Cloud Storage ou Amazon S3, endpoints de inferência e ferramentas de observabilidade.')
p('A arquitetura utiliza oito atributos orbitais como variáveis de entrada: excentricidade (e), semieixo maior (a), inclinação (i), longitude do nodo ascendente (om), argumento do periélio (w), anomalia média (ma), distância do periélio (q) e período orbital (per). O alvo é pha, convertido de Y para 1 e N para 0. Os campos MOID e H são mantidos para auditoria, mas excluídos do treinamento porque participam do critério oficial de classificação PHA.')
caption('Tabela 1 — Serviços e funções da solução')
table(['Camada','Serviço/tecnologia','Função'],[
['Fonte de dados','NASA/JPL SBDB Query API','Catálogo oficial e coleta paginada de NEAs'],
['Desenvolvimento','Google Colab / Jupyter','Notebook gerenciado e ambiente colaborativo'],
['Armazenamento','Bucket local emulado; equivalente GCS/S3','Dados, modelos, logs e relatórios versionados'],
['Treinamento','Python, scikit-learn, joblib','Comparação, treinamento e serialização dos modelos'],
['Inferência','Flask','Endpoints /health e /predict'],
['Monitoramento','Logs JSON, PSI e regras de alerta','Latência, erros, drift e recomendação de retreinamento']], [3.0,5.0,8.0])
source('Elaborado pelo autor com base no notebook do projeto e no desafio acadêmico.')

heading('4 ARQUITETURA E SERVIÇOS UTILIZADOS',1)
p('A Figura 1 apresenta o fluxo arquitetural completo. A coleta consulta a fonte oficial, grava o snapshot bruto e encaminha os dados para preparação. O conjunto preparado é dividido em treino e teste, submetido à comparação de modelos e, após a seleção, o modelo champion é gravado no registro. A API carrega o champion a partir do registro e produz predições. As requisições geram logs estruturados, que alimentam os relatórios de observabilidade, drift e governança.')
pagebreak()
caption('Figura 1 — Arquitetura da solução Asteroid Impact Intelligence')
doc.add_picture(str(OUT/'diagrama_arquitetura_asteroid_ml.png'), width=Cm(15.8))
doc.paragraphs[-1].alignment=WD_ALIGN_PARAGRAPH.CENTER
source('Elaborado pelo autor. Diagrama em preto e branco, conforme solicitado.')
p('No experimento executado, o modo USE_GCP permaneceu False. Assim, o armazenamento foi realizado em um bucket local emulado denominado asteroid-impact-ml-demo. Essa decisão permite executar a atividade sem projeto cloud com faturamento, mas preserva a abstração de objetos, gerações e versionamento. A coleta, por sua vez, foi realizada na API real da NASA/JPL e exigiu acesso à internet.')

heading('5 DESENVOLVIMENTO DA SOLUÇÃO',1)
heading('5.1 Coleta e preparação dos dados',2)
p('A coleta utilizou os filtros sb-kind=a e sb-group=neo, ordenação por spkid, precisão total e paginação de 10.000 registros. A execução registrou 43.068 registros recebidos em 13 de novembro de 2026, às 10:24:36 UTC. Após a preparação, 135 registros com rótulo PHA ausente ou inválido foram removidos, totalizando 42.933 observações válidas. Não houve remoção por identificador ausente, duplicidade, objeto não marcado como NEO ou atributo orbital não numérico.')
caption('Tabela 2 — Distribuição do conjunto de dados')
table(['Partição','Não PHA','PHA','Total'],[['Total preparado','40.381','2.552','42.933'],['Treino (80%)','32.304','2.042','34.346'],['Teste (20%)','8.077','510','8.587']], [4.0,3.0,3.0,3.0])
source('Elaborado a partir de reports/dataset_version.json.')
heading('5.2 Treinamento e seleção',2)
p('Foram comparados três algoritmos: regressão logística, Random Forest e Gradient Boosting por histogramas. A seleção foi orientada principalmente pelo F1 da classe PHA, pois a distribuição é desbalanceada e a acurácia isolada poderia ocultar o desempenho insuficiente na classe minoritária. O procedimento utilizou validação cruzada estratificada com cinco folds sobre o conjunto de treino. A regressão logística foi registrada como v1; o melhor modelo não baseline foi registrado como v2; o champion foi definido pela maior média de F1 na validação cruzada.')
caption('Tabela 3 — Comparação dos experimentos')
table(['Experimento','Acurácia CV','Recall PHA CV','F1 PHA CV','ROC AUC CV'],[['Regressão Logística','0,7311','0,6352','0,2195','0,7571'],['Random Forest','0,9331','0,1626','0,2233','0,8315'],['Gradient Boosting (Hist)','0,7875','0,7253','0,2888','0,8500']], [5.2,2.4,2.6,2.4,2.4])
source('Elaborado a partir de reports/experiments.csv.')

heading('5.3 Serviço de inferência',2)
p('A API Flask expõe GET /health para verificação de disponibilidade e POST /predict para classificação. O corpo esperado contém um objeto features com os oito atributos orbitais. A aplicação rejeita atributos ausentes, atributos extras, valores não numéricos e valores não finitos. Cada chamada registra código HTTP, latência, versão do modelo e tipo de teste, sem armazenar os valores orbitais enviados.')

heading('6 RESULTADOS E AVALIAÇÃO',1)
heading('6.1 Desempenho do modelo',2)
p('O Gradient Boosting (Hist) foi selecionado como champion na versão v2. No conjunto de teste, apresentou acurácia de 0,7607, precisão de 0,1643, recall de 0,7412, F1 de 0,2689 e ROC AUC de 0,8415 para a classe PHA. O resultado confirma que o modelo recupera parte relevante da classe minoritária, mas ainda produz muitos falsos positivos. Essa limitação é compatível com a exclusão deliberada de MOID e H, atributos diretamente relacionados ao critério oficial PHA.')
caption('Tabela 4 — Métricas finais no conjunto de teste')
table(['Versão','Algoritmo','Acurácia','Precisão PHA','Recall PHA','F1 PHA','ROC AUC'],[['v1','Regressão Logística','0,7238','0,1294','0,6373','0,2151','0,7637'],['v2 champion','Gradient Boosting (Hist)','0,7607','0,1643','0,7412','0,2689','0,8415']], [2.1,4.2,2.0,2.2,2.0,1.8,2.2])
source('Elaborado a partir de reports/metrics.json e classification_report_champion.txt.')
heading('6.2 Testes funcionais e de carga',2)
p('O endpoint /health respondeu com HTTP 200 e identificou o modelo v2. Seis chamadas de predição foram processadas com sucesso. Dois testes negativos, um com features ausentes e outro com feature extra, retornaram HTTP 400, confirmando a validação de contrato da API.')
p('No teste de carga foram enviadas 300 requisições com oito clientes concorrentes. Todas retornaram HTTP 200, sem erros, com duração de 2,306 s e vazão de 130,1 requisições por segundo. As latências observadas foram 59,45 ms no percentil 50, 79,84 ms no percentil 95 e 93,95 ms no percentil 99. Esses valores representam o loopback local e não devem ser interpretados como medição de rede, balanceamento ou autoscaling em produção.')

heading('7 MONITORAMENTO, GOVERNANÇA E MLOPS',1)
p('O monitoramento registra cada requisição em formato JSON Lines e calcula total de chamadas, chamadas bem-sucedidas, erros esperados, erros inesperados e percentis de latência. As regras definidas foram: taxa de erros inesperados menor ou igual a 1%, latência p95 inferior ou igual a 500 ms e inexistência de respostas HTTP 5xx. Na execução, todas as regras foram classificadas como OK.')
p('Para data drift, foi utilizado o Population Stability Index (PSI), com limiar de 0,2. O conjunto de teste foi comparado ao conjunto de treino e também a um lote simulado com alteração de 35% em e e redução de 18% em q. O lote simulado apresentou duas features acima do limiar de drift. A acurácia caiu de 0,7607 para 0,5010, uma redução de 25,97 pontos percentuais, e o relatório recomendou retreinamento. Trata-se de uma simulação explícita, não de tráfego real.')
p('A governança foi reforçada por manifesto com caminho, tamanho e hash SHA-256 dos artefatos, além de registro do snapshot, do modo de execução, do modelo champion e dos objetos enviados ao bucket. A execução produziu 19 artefatos, incluindo dados, modelos, logs, relatórios, figura de avaliação e pacote de evidências. Os modelos v1 e v2 foram registrados com metadados, versão do scikit-learn, features, métricas e hashes próprios.')

heading('8 OPORTUNIDADES DE MELHORIA',1)
for txt in ['Substituir o bucket local emulado por Google Cloud Storage ou Amazon S3 com versionamento, políticas de acesso mínimo, criptografia e retenção;','Executar o treinamento em serviço gerenciado, como SageMaker Training ou Vertex AI, com registro de experimentos e parâmetros;','Publicar a API atrás de gateway, autenticação, TLS, controle de taxa e autoscaling;','Adicionar calibração de probabilidades, ajuste de limiar e métricas orientadas a custo de falsos positivos e falsos negativos;','Ampliar a observabilidade com métricas de negócio, dashboards, alertas persistentes e rastreamento de lineage;','Avaliar o modelo com novos snapshots temporais e validação temporal para reduzir risco de mudança de distribuição;','Implementar pipeline automatizado de retreinamento condicionado a drift, revisão humana e aprovação do modelo.']:
    para=p('• '+txt, indent=False); para.paragraph_format.left_indent=Cm(0.5)

heading('9 CONSIDERAÇÕES FINAIS',1)
p('O projeto atendeu ao fluxo solicitado no desafio: configuração de ambiente, elaboração da arquitetura, seleção de base pública, armazenamento e preparação, treinamento e avaliação, implementação de inferência, testes, monitoramento, governança e documentação. O resultado técnico mais importante foi a construção de uma cadeia reprodutível e auditável, com separação entre dados, modelos e evidências.')
p('O modelo champion apresentou desempenho superior ao baseline em F1 e ROC AUC, mas não deve ser utilizado para decisões operacionais ou científicas sem validação adicional. A principal limitação é a natureza do alvo: PHA é uma classificação oficial baseada em critérios astronômicos, e não uma previsão de impacto. A solução deve ser entendida como protótipo educacional de MLOps e classificação supervisionada.')

heading('REFERÊNCIAS',1)
refs=[
'JET PROPULSION LABORATORY. SBDB Query API. Pasadena: California Institute of Technology, [2021]. Disponível em: https://ssd-api.jpl.nasa.gov/doc/sbdb_query.html. Acesso em: 13 nov. 2026.',
'JET PROPULSION LABORATORY. Small-Body Database Query API. Pasadena: California Institute of Technology, [2021]. Disponível em: https://ssd-api.jpl.nasa.gov/sbdb_query.api. Acesso em: 13 nov. 2026.',
'SCIKIT-LEARN DEVELOPERS. Scikit-learn: machine learning in Python. [S. l.], [s. d.]. Disponível em: https://scikit-learn.org/. Acesso em: 13 nov. 2026.',
'PYTHON SOFTWARE FOUNDATION. Python documentation. [S. l.], [s. d.]. Disponível em: https://docs.python.org/3/. Acesso em: 13 nov. 2026.',
'DESAFIO 04 — SERVIÇOS DE COMPUTAÇÃO EM NUVEM PARA MACHINE LEARNING. Material didático fornecido para a atividade acadêmica. [S. l.: s. n.], 2026.',
'RAMOS, Bruno Soares Ferreira. Asteroid Impact Intelligence: notebook de implementação do fluxo de Machine Learning em nuvem. [S. l.], 2026. Notebook Jupyter fornecido com a atividade.'
]
for r in refs:
    para=p(r, indent=False); para.paragraph_format.left_indent=Cm(0); para.paragraph_format.first_line_indent=Cm(0); para.paragraph_format.space_after=Pt(12)

# set core properties
doc.core_properties.title='Asteroid Impact Intelligence — Relatório técnico'
doc.core_properties.author='Bruno Soares Ferreira Ramos'
doc.core_properties.subject='Serviços de Computação em Nuvem para Machine Learning'
doc.save(OUT/'relatorio_tecnico_asteroid_impact_intelligence.docx')
print(OUT/'relatorio_tecnico_asteroid_impact_intelligence.docx')
