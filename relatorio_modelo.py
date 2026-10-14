"""Apresentação no layout do arquivo 'como deve ficar.xlsx'.

modelo_relatorio.xlsx contém apenas cabeçalhos e estilos, sem dados do exemplo.
As abas técnicas permanecem ocultas e podem ser reexibidas pelo Excel.
"""
from collections import Counter
from copy import copy, deepcopy
from datetime import datetime
from pathlib import Path
import math
import re
import textwrap

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from classificacao_modelo import enrich_analysis, CATEGORY_NAMES, MACRO_NAMES, MACROS, OPERATIONS

MODEL_SHEETS = ['Dashboard', 'Resumo da transmissão', 'Participação', 'Perguntas de conteúdo',
                'Picos de 5 minutos', 'Ocorrências e sugestões', 'Participantes Públicos',
                'Metodologia e lacunas', 'Transcrição dos Chats', 'Resumo_Categorias']
TEMPLATE_SHEETS = {'Resumo da transmissão': 'Resumo por vídeo'}
PRIMARY_SHEETS = ['Dashboard', 'Resumo da transmissão', 'Participação', 'Perguntas de conteúdo',
                  'Picos de 5 minutos', 'Ocorrências e sugestões', 'Participantes Públicos',
                  'Metodologia e lacunas']
TECHNICAL_SHEETS = ['Metadados', 'Transcricao', 'Transcricao_tratada', 'Chat_ao_vivo',
                    'Chat_tratado', 'Linha_do_tempo', 'Qualidade_Coleta', 'Status_Fontes',
                    'Manifesto_Fontes', 'Log_coleta', 'Classificacao_modelo']
MARKER = 'youtube_modelo:1.0'
NA = 'Não disponível'


def cell(ws, row, col, value):
    c = ws.cell(row, col, NA if value is None else value)
    if isinstance(value, str):
        c.data_type = 's'
    return c


def clone_layout(src, dst):
    # Copia objetos de estilo, não IDs internos entre workbooks.
    dst.sheet_format = copy(src.sheet_format)
    for key, dim in src.column_dimensions.items():
        target = dst.column_dimensions[key]
        for attr in ['width', 'min', 'max', 'hidden', 'outlineLevel', 'collapsed']:
            setattr(target, attr, getattr(dim, attr))
    for key, dim in src.row_dimensions.items():
        target = dst.row_dimensions[key]
        for attr in ['height', 'hidden', 'outlineLevel', 'collapsed']:
            setattr(target, attr, getattr(dim, attr))
    max_row = src.max_row if src.title == 'Dashboard' else 3
    for row in src.iter_rows(max_row=max_row):
        for c in row:
            if isinstance(c, MergedCell):
                continue
            target = dst.cell(c.row, c.column, c.value)
            target.font = copy(c.font)
            target.fill = copy(c.fill)
            target.border = copy(c.border)
            target.alignment = copy(c.alignment)
            target.number_format = c.number_format
    for region in src.merged_cells.ranges:
        dst.merge_cells(str(region))
    dst.sheet_view.showGridLines = False
    dst.freeze_panes = src.freeze_panes
    dst.sheet_properties = copy(src.sheet_properties)
    dst.page_setup = copy(src.page_setup)
    dst.page_margins = copy(src.page_margins)
    dst.sheet_properties.tabColor = '17365D'


def fill_table(ws, rows, header=1):
    ncol = next(i for i in range(1, ws.max_column + 2)
                if ws.cell(header, i).value is None) - 1
    style_row = header + 1
    styles = [(copy(ws.cell(style_row, j).font), copy(ws.cell(style_row, j).fill),
               copy(ws.cell(style_row, j).border), ws.cell(style_row, j).number_format)
              for j in range(1, ncol + 1)]
    # Mantém o cabeçalho e os dados reais; elimina apenas linhas vazias do modelo limpo.
    if ws.max_row > header:
        ws.delete_rows(header + 1, ws.max_row - header)
    for i, values in enumerate(rows, header + 1):
        lines = 1
        for j, value in enumerate(values, 1):
            c = cell(ws, i, j, value)
            font, fill, border, fmt = styles[j-1]
            c.font = font
            c.fill = PatternFill('solid', fgColor='E8F2F8' if i % 2 == 0 else 'FFFFFF')
            c.border = border
            c.number_format = fmt
            c.alignment = Alignment(wrap_text=True, vertical='top')
            width = ws.column_dimensions[get_column_letter(j)].width or 13
            lines = max(lines, math.ceil(len(str(value or '')) / max(10, width - 2)))
        ws.row_dimensions[i].height = min(400, max(24, lines * 13 + 5))
    ws.auto_filter.ref = f'A{header}:{get_column_letter(ncol)}{max(header, header+len(rows))}'
    ws.freeze_panes = f'A{header+1}'
    ws.row_dimensions[header].height = max(32, ws.row_dimensions[header].height or 0)
    if rows:
        name = 'Modelo_' + str(MODEL_SHEETS.index(ws.title))
        tab = Table(displayName=name, ref=ws.auto_filter.ref)
        tab.tableStyleInfo = TableStyleInfo(name='TableStyleMedium2', showRowStripes=True)
        ws.add_table(tab)


def mask_personal(message):
    value = re.sub(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', '[e-mail omitido]', message)
    value = re.sub(r'(?<!\d)\d{3}\.?\d{3}\.?\d{3}-?\d{2}(?!\d)', '[CPF omitido]', value)
    value = re.sub(r'(?:\+55\s*)?\(?\d{2}\)?[\s.-]*9?\d{4}[\s.-]*\d{4}\b', '[telefone omitido]', value)
    return re.sub(r'(?i)\bCRC\s*[-:/]?\s*(?:[A-Z]{2}\s*)?\d[\d./-]*', '[registro omitido]', value)


def aplicar_modelo(wb, analysis):
    from relatorio_gerencial import text, norm, SHEETS
    occupied = set(MODEL_SHEETS).intersection(wb.sheetnames)
    if occupied and ('Metodologia e lacunas' not in wb or MARKER not in {wb['Metodologia e lacunas']['A22'].value,wb['Metodologia e lacunas']['A40'].value}):
        raise ValueError('Abas do modelo já existem sem identificação de geração automática.')
    for name in occupied:
        del wb[name]
    # Compatibilidade com arquivos gerados por versões anteriores.
    if 'Resumo por vídeo' in wb.sheetnames:
        del wb['Resumo por vídeo']
    model = load_workbook(Path(__file__).with_name('modelo_relatorio.xlsx'))
    for i, name in enumerate(MODEL_SHEETS):
        clone_layout(model[TEMPLATE_SHEETS.get(name, name)], wb.create_sheet(name, i))
    model.close()
    a = enrich_analysis(analysis)
    # Unifica a métrica de elogios/agradecimentos com a taxonomia principal usada
    # nas abas de apresentação. A camada técnica anterior permanece rastreável,
    # mas o indicador com esse mesmo nome passa a usar uma única definição.
    elogios_principais = a['counts']['Elogio/agradecimento']
    if 'Indicadores' in wb.sheetnames:
        ws_ind = wb['Indicadores']
        for row in range(1, ws_ind.max_row + 1):
            if ws_ind.cell(row, 1).value == 'Elogios/agradecimentos automáticos':
                ws_ind.cell(row, 2, elogios_principais)
                ws_ind.cell(row, 3, 'Taxonomia principal do relatório; mesma contagem de Resumo da transmissão e Resumo_Categorias')
                break
    if 'Analise_chat' in wb.sheetnames:
        ws_base = wb['Analise_chat']
        for row in range(1, ws_base.max_row + 1):
            if ws_base.cell(row, 1).value == 'Elogio/agradecimento':
                ws_base.cell(row, 1, 'Elogio/agradecimento (classificação técnica exclusiva)')
                break
    md = a['md']
    title = text(md.get('titulo')) or text(md.get('video_id')) or 'Vídeo sem título'
    vid = text(md.get('video_id'))
    url = f'https://www.youtube.com/watch?v={vid}' if re.fullmatch(r'[A-Za-z0-9_-]{11}', vid) else ''
    try:
        date = datetime.strptime(text(md.get('data_upload')), '%Y%m%d')
    except ValueError:
        date = None
    p, n = a['public'], len(a['public'])
    counts, phases = a['counts'], a['phases']
    available = a['chat_available']
    metric = {label: value for label, value, *_ in a['indicators']}
    content = [r for r in p if r['pergunta_conteudo']]
    comments = [r for r in p if r['comentario_conteudo']]
    substantive = [r for r in p if r['substantiva']]
    content_bins = Counter(math.floor(r['offset']/300)*5 for r in substantive if r['offset'] is not None) if a['reliable'] else Counter()
    content_peak = min(content_bins, key=lambda b:(-content_bins[b],b)) if content_bins else None
    count = lambda category: counts[category] if available else None
    percent = lambda value: value / n if n and value is not None else None
    peak = max(a['bins'], key=lambda b: a['bins'][b][0]) if a['bins'] and any(v[0] for v in a['bins'].values()) else None
    # Resumo: exatamente as mesmas 32 colunas da referência.
    values = [title, date, title, url, a['duration']/60 if a['duration'] else None,
              n if available else None, metric['Mensagens do canal oficial'],
              metric['Participantes identificáveis'], n/(a['duration']/3600) if available and a['duration'] else None,
              len(content) if available else None, percent(len(content)), len(comments) if available else None,
              count('Crítica/sugestão'), len(substantive) if available else None, percent(len(substantive)), count('Elogio/agradecimento'),
              percent(count('Elogio/agradecimento')), count('Acesso/presença/inscrição'),
              count('Material/gravação'), count('Certificado/EPC'), count('Áudio/conexão/imagem'),
              count('Contato do palestrante'), phases['Pré-evento'] if available and a['reliable'] else None,
              metric['Concentração dos 10 IDs mais ativos'], peak, a['bins'][peak][0] if peak is not None else None,
              content_peak, content_bins[content_peak] if content_peak is not None else None, *[phases['Q'+str(i)] if available and a['reliable'] and a['duration'] else None for i in range(1,5)]]
    ws = wb['Resumo da transmissão']
    # Reproduz o cabeçalho e a faixa azul, ajustando as colunas numéricas para leitura.
    for j in range(5, 33):
        ws.column_dimensions[get_column_letter(j)].width = 19
    fill_table(ws, [values])
    ws['B2'].number_format = 'dd/mm/yyyy'
    for j in [11,15,17,24]:
        ws.cell(2,j).number_format = '0.0%'
    for j in [5,9]:
        ws.cell(2,j).number_format = '0.0'
    ws.row_dimensions[1].height = 46

    ws = wb['Participação']
    ws.column_dimensions['A'].width = 54
    repeated = Counter((text(r.get('canal_autor_id')), norm(r.get('mensagem'))) for r in p)
    known_messages = sum(map(len, a['participants'].values()))
    participant_rows = [
        ['Participantes únicos no vídeo', a['identity']['participantes_publico_ids'] if available else None],
        ['Participantes com 1 mensagem', sum(len(v)==1 for v in a['participants'].values()) if available else None],
        ['Participantes com 2 a 3 mensagens', sum(2<=len(v)<=3 for v in a['participants'].values()) if available else None],
        ['Participantes com 4+ mensagens', sum(len(v)>=4 for v in a['participants'].values()) if available else None],
        ['Mensagens repetidas (excedentes)', sum(v-1 for (aid,msg),v in repeated.items() if aid and msg) if available else None],
        ['Mensagens por participante', known_messages/a['identity']['participantes_publico_ids'] if a['identity']['participantes_publico_ids'] else None]]
    ws.column_dimensions['B'].width=27
    fill_table(ws,participant_rows,2)
    ws['B8'].number_format='0.00'

    ws=wb['Picos de 5 minutos']
    for j in range(3,10): ws.column_dimensions[get_column_letter(j)].width=21
    ws.column_dimensions['K'].width=60
    ranked = sorted((b for b,v in a['bins'].items() if v[0]), key=lambda b:(-a['bins'][b][0],b))[:3]
    rows=[]
    for rank,b in enumerate(ranked,1):
        in_bin=lambda r:r['offset'] is not None and b*60<=r['offset']<(b+5)*60
        # Exemplos compactos; mensagens integrais continuam nas bases técnicas.
        examples=' | '.join(mask_personal(text(r.get('mensagem')))[:180] for r in [r for r in p if in_bin(r)][:3])
        team=' | '.join(mask_personal(text(r.get('mensagem')))[:180] for r in [r for r in a['official'] if in_bin(r)][:3])
        rows.append([title,date,rank,b,b+5,a['bins'][b][0],sum(in_bin(r) for r in substantive),sum(in_bin(r) for r in content),
                     a['bins'][b][3] if text(md.get('canal_id')) else None,examples,team])
    fill_table(ws,rows)

    ws=wb['Perguntas de conteúdo']
    ws.column_dimensions['C'].width=21
    fill_table(ws,[[title,date,r.get('tempo_video'),mask_personal(text(r.get('mensagem'))),
                    url+f'&t={max(0,int(r["offset"]))}s' if url and r['offset'] is not None and a['reliable'] else url]
                   for r in content])

    ws=wb['Transcrição dos Chats']
    for j,width in enumerate([18,48,42,20,32,85,30,27],1):ws.column_dimensions[get_column_letter(j)].width=width
    fill_table(ws,[[date,title,url,r.get('tempo_video'),r.get('autor'),mask_personal(text(r.get('mensagem'))),
                    r['categoria_modelo'],r['macrotema_modelo']] for r in a['analysis']])

    ws=wb['Resumo_Categorias']
    for j,width in enumerate([23,40,20,20],1): ws.column_dimensions[get_column_letter(j)].width=width
    categories=Counter(r['categoria_modelo'] for r in p)
    # Os percentuais de macrotema usam somente perguntas de conteúdo, como no exemplo.
    topics=Counter(r['macrotema_modelo'] for r in content)
    category_rows=[['Categoria',c,categories[c] if available else None,percent(categories[c])] for c in CATEGORY_NAMES]
    category_rows += [['Macrotema de conteúdo',c,topics[c] if available else None,topics[c]/len(content) if content else None] for c in MACRO_NAMES]
    fill_table(ws,category_rows)
    for i in range(2,ws.max_row+1):ws.cell(i,4).number_format='0.0%'

    ws=wb['Ocorrências e sugestões']
    ws.column_dimensions['C'].width=20
    ws.column_dimensions['D'].width=30
    ws.column_dimensions['H'].width=45
    fill_table(ws,[[title,date,r.get('tempo_video'),r['categoria_detalhada'],mask_personal(text(r.get('mensagem'))),
                    r['operacional_modelo'],r['categoria_modelo']=='Presença, inscrição ou certificado',url]
                   for r in p if r['operacional_modelo'] or r['sugestao_critica']])

    ws=wb['Participantes Públicos']
    ws.column_dimensions['B'].width=36
    ws.column_dimensions['D'].width=26
    content_ids=Counter(text(r.get('canal_autor_id')) for r in content)
    fill_table(ws,[[text(msgs[0].get('autor')),aid,content_ids[aid],int(content_ids[aid]>0)]
                   for aid,msgs in sorted(a['participants'].items(),key=lambda x:(-len(x[1]),x[0]))])

    ws=wb['Metodologia e lacunas']
    for j,width in enumerate([28,80,74,70],1):ws.column_dimensions[get_column_letter(j)].width=width
    method=[
        ['Escopo','Um vídeo por execução. Publicação: '+(date.strftime('%d/%m/%Y') if date else NA),
         'A data de publicação pode não ser o início real da transmissão.','Não interpretar como portfólio mensal.'],
        ['Público e equipe','Público externo = autores identificados por ID, excluindo o canal oficial e contas institucionais/equipe CRCSP, que são apresentadas separadamente.',
         'A identificação de contas institucionais/equipe por nome público é heurística e permanece auditável.','IDs e classificação original preservados em Mensagens_analise.'],
        ['Classificação textual','Regras automáticas da camada gerencial, sem revisão humana.',
         'As categorias, perguntas e ocorrências são sinais exploratórios.','Revisar a mensagem original antes de concluir.'],
        ['Perguntas de conteúdo','Intenção interrogativa (com ou sem ?) + termo de macrotema, após excluir regras operacionais e críticas. São reconhecidos padrões como como/qual/onde, posso/devo, é possível, tem como, gostaria de saber, saber se e será que.',
         'Triagem automática por linguagem e palavras-chave, não confirmação de pergunta técnica.','Mensagens interrogativas sem macrotema continuam disponíveis em Mensagens_analise para revisão.'],
        ['Comentários/substantivas','Comentário: termo temático + pelo menos 4 palavras, sem regra anterior. Substantivas = perguntas + comentários + sugestões/críticas.',
         'Classificação exploratória por palavras-chave, sem revisão humana.','Pico conteúdo = janela de 5 minutos com mais substantivas; empate pela primeira janela.'],
        ['Problemas técnicos','Menções automáticas a áudio, conexão ou imagem.',
         'Podem ser relatos, elogios ou mensagens informativas; não comprovam falhas.','Conferir gravação ou logs antes de concluir incidente.'],
        ['Picos','Top 3 janelas fixas de 5 minutos por mensagens do público externo. Fim exclusivo.',
         'Pré/pós incluídos. Empates pela primeira janela.','Aba técnica Picos_5_min preserva todas as janelas.'],
        ['Participação','Participantes do público externo identificados por IDs distintos dentro deste vídeo, excluindo canal oficial e contas institucionais/equipe.',
         'Não equivale a espectadores. Registros sem ID não entram no total de participantes.','Mensagens por participante usa somente mensagens com ID.'],
        ['Mensagens repetidas','Excedentes do mesmo ID e texto normalizado, após a deduplicação técnica.',
         'Não são duplicidades removidas da coleta.','Nenhuma mensagem é excluída pelo relatório.'],
        ['Macrotemas','Nove grupos por termos da própria mensagem. Primeiro grupo correspondente, na ordem documentada. Não usa título como evidência da mensagem.',
         'Percentuais de macrotema: perguntas do grupo / todas as perguntas de conteúdo. Outros registros podem ter Não identificado ou Não se aplica.','As regras e os termos encontrados estão na aba técnica Classificacao_modelo.'],
        ['Privacidade','E-mails, CPF, telefones e registros CRC reconhecidos por padrões são omitidos nas abas textuais.',
         'A identificação automática pode não reconhecer todos os formatos. Nomes públicos permanecem.','Texto integral preservado nas abas técnicas ocultas.'],
        ['Fontes','Metadados, Chat_tratado, Chat_ao_vivo, transcrição e controles V7.',
         'As bases técnicas principais ficam visíveis após as abas de apresentação; abas intermediárias de cálculo permanecem ocultas.','As abas técnicas visíveis permitem auditoria sem poluir a navegação principal.'],
        ['Data da coleta',text(md.get('data_hora_coleta')) or NA,
         'Visualizações e curtidas acumuladas estão nos indicadores técnicos.','Não confundir com a data de geração do relatório.'],
        ['Qualidade da coleta',text(a['quality'].get('status_geral')) or NA,
         'Gate técnico não confirma qualidade do evento.','Consultar Qualidade_Coleta e Validacao_relatorio.'],
        ['Atualização','Agregados calculados pelo Python ao gerar o arquivo.',
         'Edição manual de mensagens não recalcula as classificações.','Executar novamente gerar_relatorio.py.'],
        ['Rastreabilidade','Mensagens_analise registra linha bruta, linha tratada e chave.',
         'A ordem da Transcrição dos Chats acompanha Mensagens_analise.','Linha n do chat visível corresponde à linha n+3 da análise técnica.'],
        ['Regras e conciliação','Metodologia e controles internos de cálculo.',
         'Percentuais usam frações numéricas com formato 0,0%.','Consultar todas as verificações vigentes em Validacao_relatorio; a quantidade é gerada pelo próprio relatório.']]
    method += [
        ['Resumo por categorias','As seis categorias fornecidas são mantidas. Comentários, críticas, elogios e outro/indeterminado completam a conciliação.',
         'Categorias exclusivas: quantidade / mensagens do público externo. Macrotemas têm outro denominador.','Não somar percentuais de categorias com percentuais de macrotemas.'],
        ['Mensagens/h','Mensagens do público externo / duração do vídeo em horas, como nos exemplos enviados.',
         'Inclui mensagens pré e pós-evento no numerador.','O indicador técnico Mensagens por hora durante o vídeo continua separado na aba Indicadores.'],
        ['Dados de audiência','Visualizações e curtidas disponíveis quando fornecidas pelo YouTube; consultar Metadados.',
         'Sem retenção, espectadores únicos ou confirmação de presença.','Não usar chat como substituto dessas métricas.'],
        ['Fonte primária','Coleta V7 do vídeo '+vid,'Bases preservadas no mesmo arquivo; não foi usado o CSV de exemplo.','Data da coleta: '+(text(md.get('data_hora_coleta')) or NA)],
        ['Observação','Classificações automáticas não equivalem a revisão amostral.','O relatório não substitui os registros originais.','Consultar as regras e validar mensagens ambíguas.']]
    fill_table(ws,method,2)
    # Marcador fixo fora da tabela para permitir regeneração sem duplicar abas.
    cell(ws,40,1,MARKER)
    ws.row_dimensions[40].hidden=True

    # Dashboard usa as mesmas células, blocos de cor e ordem do modelo.
    ws=wb['Dashboard']
    for col in ['J','K','L']: ws.column_dimensions[col].width=15
    ws.row_dimensions[4].height=30
    for col in [1,3,5,7,9,11]:ws.cell(4,col).alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    ws['A1']='Análise dos chats ao vivo — '+(text(md.get('canal')) or 'YouTube')
    ws['A3']='Base: 1 vídeo. Público externo, equipe/contas institucionais e canal oficial são populações separadas por identidade. Classificações são automáticas.'
    ws['A3'].alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[3].height=32
    cards={'A5':1,'C5':values[5],'E5':values[7],'G5':'Não se aplica',
           'I5':values[15],'K5':sum(r['operacional_modelo'] for r in p) if available else None}
    for addr,value in cards.items():
        c=ws[addr]
        c.value=value if value is not None else NA
        if isinstance(c.value,str):c.data_type='s'
    ws['G5'].font=Font(name='Calibri',size=14,bold=True,color='17365D')
    ws['G5'].alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    for addr,value in {'A10':'Equipe no chat','B10':values[6],'C10':'Perguntas de conteúdo','D10':values[9],
                       'E10':'Mensagens substantivas','F10':values[13],'G10':'Problemas técnicos','H10':values[20],
                       'I10':'Material/gravação','J10':values[18],'K10':'Acesso/presença/início','L10':values[17]}.items():
        cell(ws,ws[addr].row,ws[addr].column,value)
    ws.row_dimensions[10].height=46
    data=[title,values[5],values[7],values[8],values[9],values[14],values[16],cards['K5'],values[20]]
    for j,v in enumerate(data,1):cell(ws,14,j,v)
    ws.row_dimensions[14].height=70
    for c in ws[14][:9]:
        c.alignment=Alignment(wrap_text=True,vertical='center')
        c.fill=PatternFill('solid',fgColor='C6EAF6')
    ws['G14'].number_format='0.0%'
    ws['F14'].number_format='0.0%'
    ws['D14'].number_format='0.0'
    ws.merge_cells('A26:I28')
    cell(ws,26,1,'Perguntas e problemas técnicos são sinais automáticos, não fatos validados. Consulte Metodologia e lacunas. As bases técnicas principais estão organizadas ao final do arquivo.').alignment=Alignment(wrap_text=True,vertical='center')
    ws.merge_cells('A30:I32')
    cell(ws,30,1,'Status da coleta: '+(text(a['quality'].get('status_geral')) or NA)).alignment=Alignment(wrap_text=True)
    ws.merge_cells('A34:I36')
    cell(ws,34,1,'Vídeo: '+title).alignment=Alignment(wrap_text=True)
    for j,v in enumerate(['Evento','Mensagens','Participantes'],26):cell(ws,2,j,v)
    for j,v in enumerate([title,values[5],values[7]],26):cell(ws,3,j,v)
    cell(ws,32,26,'Categoria');cell(ws,32,27,'Mensagens')
    for row,(label,value) in enumerate([('Acesso/presença/início',values[17]),('Material/gravação',values[18]),('Problemas técnicos',values[20]),('Certificado/EPC',values[19])],33):
        cell(ws,row,26,label);cell(ws,row,27,value)
    if available:
        for start,end,mincol,maxcol,anchor,title_chart in [(2,3,27,28,'K13','Mensagens e participantes'),(32,36,27,27,'K32','Demandas operacionais (sinais)')]:
            chart=BarChart()
            chart.type='bar'
            chart.title=title_chart
            chart.style=13
            chart.width,chart.height=17,8.5
            chart.add_data(Reference(ws,min_col=mincol,max_col=maxcol,min_row=start,max_row=end),titles_from_data=True)
            chart.set_categories(Reference(ws,min_col=26,min_row=start+1,max_row=end))
            chart.dataLabels=DataLabelList();chart.dataLabels.showVal=True
            for i,series in enumerate(chart.series):
                series.tx=SeriesLabel(v=str(ws.cell(start,mincol+i).value))
                series.graphicalProperties.solidFill=['2F75B5','5BB8C8'][i%2]
                series.graphicalProperties.line.solidFill=['2F75B5','5BB8C8'][i%2]
            chart.legend.position='b'
            if mincol==maxcol:chart.legend=None
            chart.visible_cells_only=False
            ws.add_chart(chart,anchor)
    ws.conditional_formatting.add('H14:I14',DataBarRule(start_type='num',start_value=0,end_type='max',color='F4B183'))
    ws.column_dimensions.group('Z','AB',hidden=True)
    ws.freeze_panes='B14'
    ws.print_area='A1:U49'
    write_classification_audit(wb,a,categories,topics,content,comments,substantive,content_bins)
    for name in MODEL_SHEETS:
        sheet=wb[name]
        # Datas permanecem valores Excel, com apresentação brasileira.
        if name not in {'Dashboard','Participação','Resumo_Categorias','Participantes Públicos','Metodologia e lacunas'}:
            col=1 if name=='Transcrição dos Chats' else 2
            for row in sheet.iter_rows(min_row=2,min_col=col,max_col=col):
                if isinstance(row[0].value,datetime):row[0].number_format='dd/mm/yyyy'
    ajustar_formatacao(wb)
    from leitura_clara import apply
    apply(wb,a)

    # Navegação final: primeiro as 8 abas de apresentação; depois as bases técnicas úteis.
    # Abas de apoio/intermediárias continuam no arquivo para rastreabilidade, mas ficam ocultas.
    visible = [name for name in PRIMARY_SHEETS + TECHNICAL_SHEETS if name in wb.sheetnames]
    for sheet in wb:
        sheet.sheet_state = 'visible' if sheet.title in visible else 'hidden'
    ordered = [wb[name] for name in visible]
    ordered += [sheet for sheet in wb.worksheets if sheet.title not in visible]
    wb._sheets = ordered
    wb.active = 0


def ajustar_formatacao(wb):
    """Ajusta somente apresentação; também pode corrigir um relatório já coletado."""
    widths = {
        'Resumo da transmissão': [34, 17, 54, 46] + [20] * 28,
        'Participação': [54, 28],
        'Picos de 5 minutos': [34, 17, 9, 15, 15, 20, 23, 22, 20, 64, 56],
        'Perguntas de conteúdo': [34, 17, 22, 85, 48],
        'Transcrição dos Chats': [17, 42, 46, 21, 30, 85, 38, 42],
        'Resumo_Categorias': [26, 54, 18, 18],
        'Ocorrências e sugestões': [34, 17, 21, 32, 85, 24, 24, 46],
        'Participantes Públicos': [38, 38, 18, 28],
        'Metodologia e lacunas': [30, 66, 60, 60],
    }

    def height(c, width, header=False):
        # Conta quebras explícitas e de palavras, com margem para fontes do Excel.
        size = c.font.sz or 11
        chars = max(5, int((width - (4 if header else 2)) * 0.85 * 11 / size))
        lines = sum(max(1, len(textwrap.wrap(part, width=chars, break_long_words=True,
                                            break_on_hyphens=False)))
                    for part in str(c.value or '').split('\n'))
        return min(409, max(27, lines * size * 1.4 + 10))

    for name, sizes in widths.items():
        if name not in wb:
            continue
        ws = wb[name]
        header = 2 if name in {'Participação', 'Metodologia e lacunas'} else 1
        for col, width in enumerate(sizes, 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        last = max((c.row for row in ws for c in row
                    if c.value is not None and not ws.row_dimensions[c.row].hidden), default=header)
        for row in ws.iter_rows(min_row=header, max_row=last, max_col=len(sizes)):
            heights = []
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                align = copy(c.alignment)
                align.wrap_text = True
                align.shrink_to_fit = False
                align.vertical = 'center'
                align.horizontal = ('center' if c.row == header or isinstance(c.value, (datetime, bool))
                                    else 'right' if isinstance(c.value, (int, float)) else 'left')
                align.indent = 0 if c.row == header else 1
                c.alignment = align
                heights.append(height(c, sizes[c.column-1], c.row == header))
            ws.row_dimensions[row[0].row].height = max(heights, default=27)
        ws.freeze_panes = ('C2' if name == 'Resumo da transmissão' else f'A{header+1}')
        ws.sheet_view.zoomScale = 85 if len(sizes) > 8 else 90
        ws.print_title_rows = f'1:{header}'
        ws.print_area = f'A1:{get_column_letter(len(sizes))}{last}'
        ws.page_setup.orientation = 'landscape'
        # Evita reduzir 32 colunas a uma única página ilegível.
        ws.sheet_properties.pageSetUpPr.fitToPage = False
        ws.page_setup.scale = 85
        if header == 2:
            ws.row_dimensions[1].height = 36

    ws = wb['Dashboard']
    for j in range(2, 13):
        ws.column_dimensions[get_column_letter(j)].width = 17
    for label_col, value_col in [('C', 'D'), ('E', 'F'), ('G', 'H'), ('I', 'J')]:
        ws.column_dimensions[label_col].width = 23
        ws.column_dimensions[value_col].width = 11
    ws.column_dimensions['K'].width = 28
    ws.column_dimensions['L'].width = 9
    ws.row_dimensions[4].height = 42
    ws.row_dimensions[10].height = 64
    ws.row_dimensions[13].height = 42
    for rownum in [9, 10, 13, 14]:
        for c in ws[rownum][:12 if rownum in [9, 10] else 9]:
            if c.value is not None:
                c.alignment = Alignment(wrap_text=True, vertical='center',
                                        horizontal='center' if rownum in [9, 13] else
                                        'right' if isinstance(c.value, (int, float)) else 'left', indent=1)
    ws.row_dimensions[14].height = max(70, height(ws['A14'], ws.column_dimensions['A'].width))
    for addr in ['A26', 'A30', 'A34']:
        ws[addr].alignment = Alignment(wrap_text=True, vertical='center', horizontal='left')
    # Usa o ID já existente como rótulo curto, sem alterar as células fonte.
    # O título integral permanece na tabela e no rodapé.
    if ws._charts and 'Metadados' in wb:
        for row in wb['Metadados']:
            if row[0].value == 'video_id':
                ws._charts[0].set_categories(Reference(wb['Metadados'], min_col=2,
                                                       min_row=row[0].row, max_row=row[0].row))
                break
    ws.freeze_panes = None
    ws.sheet_view.zoomScale = 80
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1


def write_classification_audit(wb,a,categories,topics,content,comments,substantive,bins):
    from relatorio_gerencial import heading, table
    name='Classificacao_modelo'
    if name in wb:
        if wb[name]['A3'].value!='classificacao_modelo:1.0':raise ValueError('Aba Classificacao_modelo já existe sem marcador.')
        del wb[name]
    ws=wb.create_sheet(name)
    heading(ws,'Classificação automática do modelo','Rastreabilidade por chave e linha de Chat_tratado. Nenhuma classificação foi validada por pessoa.','H')
    cell(ws,3,1,'classificacao_modelo:1.0')
    table(ws,['Chave','Linha Chat_tratado','População','Categoria','Detalhe','Macrotema','Termos encontrados',
              'Pergunta conteúdo','Comentário conteúdo','Sugestão/crítica','Substantiva','Operacional','Regra'],
          [[r.get('chave_evento'),r['_linha'],r['populacao'],r['categoria_modelo'],r['categoria_detalhada'],r['macrotema_modelo'],r['termos_macrotema'],
            r['pergunta_conteudo'],r['comentario_conteudo'],r['sugestao_critica'],r['substantiva'],r['operacional_modelo'],r['regra_modelo']] for r in a['analysis']],
          [35,20,24,38,34,50,36,23,23,22,20,20,90])
    end=ws.max_row+3
    cell(ws,end,1,'Regras de macrotemas (ordem de prioridade)')
    for i,(label,pattern) in enumerate(MACROS,end+1):cell(ws,i,1,label);cell(ws,i,2,pattern)
    # Controle independente dos agregados publicados, sem alterar o gate V7.
    checks=[('Modelo: categorias = público',len(a['public']),sum(categories.values())),
            ('Modelo: macrotemas = perguntas',len(content),sum(topics.values())),
            ('Modelo: substantivas = perguntas + comentários + críticas',len(substantive),len(content)+len(comments)+sum(r['sugestao_critica'] for r in a['public']))]
    if a['reliable']:checks.append(('Modelo: janelas conteúdo = substantivas com tempo',sum(r['offset'] is not None for r in substantive),sum(bins.values())))
    valid=wb['Validacao_relatorio']
    for label,expected,actual in checks:
        valid.append([label,expected,actual,'OK' if expected==actual else 'DIVERGÊNCIA','Taxonomia automática documentada em Classificacao_modelo'])
    for table_obj in valid.tables.values():table_obj.ref=f'A4:E{valid.max_row}'

