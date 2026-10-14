"""Apresentação acessível dos resultados, sem alterar registros ou classificação."""
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter

BLUE, GOLD, INK, PALE = '21609A', 'BD993E', '243746', 'EDF3F8'
LABELS = {
    'Mensagens público': 'Mensagens do público externo',
    'Mensagens equipe': 'Mensagens da conta do canal',
    'Participantes do Chat': 'Contas participantes no chat',
    'Perguntas conteúdo': 'Perguntas sobre o tema (automático)',
    'Perguntas conteúdo %': 'Perguntas / mensagens (%)',
    'Comentários conteúdo': 'Comentários sobre o tema (automático)',
    'Sugestões/críticas': 'Sugestões ou críticas (automático)',
    'Mensagens substantivas': 'Perguntas + comentários + sugestões',
    'Substantivas %': 'Perguntas + comentários + sugestões (%)',
    'Elogios/agradecimentos': 'Elogios ou agradecimentos (automático)',
    'Elogios %': 'Elogios / mensagens (%)',
    'Problemas técnicos': 'Menções técnicas (automático)',
    'Pré-live': 'Mensagens antes do vídeo',
    'Top 10 participantes %': 'Mensagens das 10 contas mais ativas (%)',
    'Pico bruto (min)': 'Início do maior pico (min)',
    'Pico bruto msgs': 'Mensagens no maior pico',
    'Pico conteúdo (min)': 'Início do pico de contribuições (min)',
    'Pico conteúdo msgs': 'Contribuições no pico (automático)',
    'Q1': 'Primeiro quarto do vídeo', 'Q2': 'Segundo quarto do vídeo',
    'Q3': 'Terceiro quarto do vídeo', 'Q4': 'Último quarto do vídeo',
    'Mensagens substantivas': 'Perguntas + comentários + sugestões',
}

def canonical(label):
    return next((old for old,new in LABELS.items() if new==label), label)

def val(value):
    return value if isinstance(value,(int,float)) else None

def fmt(value):
    return f'{value:,.0f}'.replace(',', '.') if val(value) is not None else 'não disponível'

def summary_text(s, institutional_msgs=None):
    n, ids, team = s.get('Mensagens público'), s.get('Participantes do Chat'), s.get('Mensagens equipe')
    inst = fmt(institutional_msgs) if institutional_msgs is not None else 'não disponível'
    return (f'O chat analisado contém {fmt(n)} mensagens do público externo e '
            f'{fmt(ids)} participantes do público externo identificados por ID. '
            f'Contas institucionais/equipe enviaram {inst} mensagens e a conta oficial do canal enviou {fmt(team)} mensagens. '
            'Esses participantes não representam o total de espectadores.')

def explain(s):
    q,c,g,total = [s.get(k) for k in ['Perguntas conteúdo','Comentários conteúdo','Sugestões/críticas','Mensagens substantivas']]
    if g == 1:
        sugestao_txt = '1 sugestão ou crítica'
    else:
        sugestao_txt = f'{fmt(g)} sugestões ou críticas'
    return (f'A triagem marcou {fmt(q)} perguntas, {fmt(c)} comentários sobre o tema e '
            f'{sugestao_txt}: {fmt(total)} mensagens ao todo. '
            'As perguntas já fazem parte desse total. As classificações precisam de revisão.')

def apply(wb,a):
    """Recebe a análise enriquecida; não recalcula categorias nem métricas."""
    ws=wb['Resumo da transmissão']
    s={canonical(c.value):ws.cell(2,c.column).value for c in ws[1] if c.value}
    # A tabela técnica mantém duração em minutos; acrescenta a versão legível.
    col=ws.max_column+1
    ws.cell(1,col,'Duração (h:mm:ss)')
    seconds=a.get('duration')
    ws.cell(2,col,seconds/86400 if seconds else 'Não disponível').number_format='[h]:mm:ss'
    ws.column_dimensions[get_column_letter(col)].width=20
    for table in ws.tables.values():
        table.ref=f'A1:{get_column_letter(col)}2'
    ws.auto_filter.ref=f'A1:{get_column_letter(col)}2'
    ws['A2']=str(s.get('Título capturado') or '')[:75]
    ws['A2'].data_type='s'
    # Nomes simples nas tabelas. As abas originais e de auditoria ficam intactas.
    names=['Resumo da transmissão','Picos de 5 minutos']
    for name in names:
        sheet=wb[name]
        for c in sheet[1]:
            if c.value in LABELS: c.value=LABELS[c.value]
        sheet.row_dimensions[1].height=82 if name=='Resumo da transmissão' else 65
        for c in sheet[1]:
            c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
            c.fill=PatternFill('solid',fgColor=BLUE)
            c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
    # Completa a visão dos participantes deste vídeo.
    people=wb['Participantes Públicos']
    people.cell(1,5,'Mensagens enviadas')
    for row in range(2,people.max_row+1):
        aid=str(people.cell(row,2).value or '')
        people.cell(row,5,len(a['participants'].get(aid,[])))
    for col in (5,):
        people.column_dimensions[get_column_letter(col)].width=22
        for row in range(1,people.max_row+1):
            c=people.cell(row,col)
            c.font=Font(name='Arial',size=10,bold=row==1,color='FFFFFF' if row==1 else INK)
            c.fill=PatternFill('solid',fgColor=BLUE if row==1 else PALE if row%2==0 else 'FFFFFF')
            c.alignment=Alignment(wrap_text=True,vertical='center')
    for table in people.tables.values(): table.ref=f'A1:E{max(1,people.max_row)}'
    people.auto_filter.ref=f'A1:E{max(1,people.max_row)}'
    people.print_area=f'A1:E{max(1,people.max_row)}'
    # Dashboard compacto, com contagens e interpretações em seções distintas.
    del wb['Dashboard']
    dash=wb.create_sheet('Dashboard',0)
    dash.sheet_view.showGridLines=False
    dash.sheet_view.zoomScale=90
    dash.sheet_properties.tabColor=BLUE
    for i in range(1,13): dash.column_dimensions[get_column_letter(i)].width=12
    for row in range(1,92): dash.row_dimensions[row].height=22

    def block(row,c1,c2,value,height=28,bg=None,size=11,bold=False,fmtcode=None):
        dash.merge_cells(start_row=row,start_column=c1,end_row=row,end_column=c2)
        c=dash.cell(row,c1,'Não disponível' if value is None else value)
        if isinstance(value,str): c.data_type='s'
        c.font=Font(name='Arial',size=size,bold=bold,color=INK if not bg or bg==PALE else 'FFFFFF')
        c.alignment=Alignment(wrap_text=True,vertical='center')
        if bg:c.fill=PatternFill('solid',fgColor=bg)
        if fmtcode:c.number_format=fmtcode
        dash.row_dimensions[row].height=max(dash.row_dimensions[row].height or 22,height)
        return c

    block(2,1,10,'Relatório da transmissão',size=18,bold=True,height=34)
    logo=Path(__file__).parent/'web/static/crcsp-logo.png'
    if logo.exists():
        image=Image(str(logo));image.width=74;image.height=60;dash.add_image(image,'K1')
    block(4,1,12,s.get('Título capturado'),height=44,bold=True)
    block(5,1,12,'Análise individual da transmissão selecionada.',height=28)
    block(6,1,3,'Data de publicação',bold=True)
    block(6,4,6,s.get('Data'),fmtcode='dd/mm/yyyy')
    block(6,7,9,'Duração do vídeo',bold=True)
    block(6,10,12,seconds/86400 if seconds else None,fmtcode='[h]:mm:ss')
    block(8,1,12,'1. Participação no chat',bg=BLUE,bold=True)
    institutional_msgs = len(a.get('institutional') or []) if a.get('chat_available') else None
    cards = [
        (1, 3, val(s.get('Mensagens público')), 'Mensagens do público externo', 'Exclui canal oficial, contas institucionais/equipe e mensagens de sistema.'),
        (4, 6, val(s.get('Participantes do Chat')), 'Participantes do público externo', 'IDs distintos do público externo; exclui contas institucionais/equipe.'),
        (7, 9, institutional_msgs, 'Mensagens institucionais/equipe', 'Contas associadas ao CRCSP identificadas separadamente do público externo.'),
        (10, 12, val(s.get('Mensagens equipe')), 'Mensagens do canal oficial', 'Somente o ID oficial responsável pelo vídeo.'),
    ]
    for left,right,value,label,note in cards:
        block(9,left,right,label,bold=True,height=32)
        block(10,left,right,value,bg=PALE,size=20,bold=True,height=36,fmtcode='#,##0')
        block(11,left,right,note,height=44,size=9)
    block(13,1,12,summary_text(s, institutional_msgs),height=52)
    block(14,1,12,'Contas com nome público associado ao CRCSP são classificadas separadamente como institucionais/equipe. Mensagens sem identificador entram no total de mensagens, mas não no total de participantes.',height=44,size=10)
    block(16,1,12,'2. O que aparece nas mensagens',bg=BLUE,bold=True)
    block(17,1,12,'Classificação automática por palavras-chave. As quantidades abaixo são pistas para revisão, não avaliações confirmadas.',height=32,size=10)
    for row,(label,key) in enumerate([('Possíveis perguntas sobre o tema','Perguntas conteúdo'),('Possíveis comentários sobre o tema','Comentários conteúdo'),('Sugestões ou críticas identificadas','Sugestões/críticas'),('Total dos três grupos acima','Mensagens substantivas'),('Elogios ou agradecimentos identificados','Elogios/agradecimentos')],19):
        block(row,1,9,label,bold=row==22,bg=PALE if row==22 else None)
        block(row,10,12,val(s.get(key)),bold=True,fmtcode='#,##0')
    block(25,1,12,explain(s),height=48)
    block(27,1,12,'3. Pedidos e menções para a organização revisar',bg=BLUE,bold=True)
    for row,(label,key) in enumerate([('Acesso, presença ou início','Acesso/presença/início'),('Materiais ou gravação','Material/gravação'),('Certificado ou EPC','Certificado/EPC'),('Áudio, conexão ou imagem','Problemas técnicos'),('Contato do palestrante','Contato palestrante')],29):
        block(row,1,9,label)
        block(row,10,12,val(s.get(key)),fmtcode='#,##0')
    block(35,1,12,'Uma menção a áudio ou conexão não comprova falha. Leia as mensagens em Ocorrências e sugestões antes de concluir que houve um problema.',height=40,size=10)
    block(37,1,12,'4. Como o chat se distribuiu',bg=BLUE,bold=True)
    cats=wb['Resumo_Categorias']
    rows=[r for r in cats.iter_rows(min_row=2,values_only=True) if r[0]=='Categoria']
    for i,row in enumerate(rows,2):
        dash.cell(i,15,row[1]);dash.cell(i,16,row[2])
    dash.cell(1,15,'Categoria automática');dash.cell(1,16,'Mensagens')
    if a['chat_available']:
        chart=BarChart();chart.type='bar';chart.title='Mensagens por categoria automática'
        chart.add_data(Reference(dash,min_col=16,min_row=1,max_row=1+len(rows)),titles_from_data=True)
        chart.set_categories(Reference(dash,min_col=15,min_row=2,max_row=1+len(rows)))
        chart.legend=None;chart.width=27;chart.height=9
        chart.x_axis.title='Mensagens do público externo';chart.dataLabels=DataLabelList();chart.dataLabels.showVal=True
        chart.series[0].graphicalProperties.solidFill=BLUE;chart.visible_cells_only=False
        dash.add_chart(chart,'A39')
    else:block(39,1,12,'Chat indisponível: não interpretar como ausência de participação.',height=38)
    if a['reliable'] and a['bins']:
        dash.cell(1,18,'Início do intervalo (min)');dash.cell(1,19,'Mensagens do público externo')
        for i,(minute,counts) in enumerate(a['bins'].items(),2):
            dash.cell(i,18,minute);dash.cell(i,19,counts[0])
        line=LineChart();line.title='Mensagens em intervalos de 5 minutos';line.width=27;line.height=7
        line.add_data(Reference(dash,min_col=19,min_row=1,max_row=len(a['bins'])+1),titles_from_data=True)
        line.set_categories(Reference(dash,min_col=18,min_row=2,max_row=len(a['bins'])+1))
        line.x_axis.title='Minutos em relação ao início';line.y_axis.title='Mensagens';line.legend=None;line.visible_cells_only=False
        dash.add_chart(line,'A58')
    else:block(58,1,12,'Tempo do chat indisponível ou não confiável. Não é possível mostrar a evolução por minuto.',height=40)
    block(73,1,12,'5. Onde consultar os detalhes',bg=BLUE,bold=True)
    for row,(name,description) in enumerate([('Resumo da transmissão','Todos os indicadores e percentuais deste vídeo'),('Perguntas de conteúdo','Mensagens classificadas como perguntas para revisar'),('Ocorrências e sugestões','Pedidos, menções operacionais e sugestões'),('Chat_tratado','Base técnica do chat tratado e normalizado'),('Participantes Públicos','Contas, mensagens e perguntas identificadas'),('Metodologia e lacunas','Critérios, denominadores e limitações')],75):
        block(row,1,5,name).hyperlink=f"#'{name}'!A1"
        block(row,6,12,description,height=30)
    block(83,1,12,'As abas principais vêm primeiro. As bases técnicas de coleta e auditoria ficam organizadas depois delas; apenas abas internas de cálculo permanecem ocultas. Os agregados são atualizados ao gerar novamente o relatório.',height=42,size=10)
    dash.column_dimensions.group('O','S',hidden=True)
    dash.print_area='A1:L84';dash.page_setup.orientation='portrait'
    dash.sheet_properties.pageSetUpPr.fitToPage=True;dash.page_setup.fitToWidth=1;dash.page_setup.fitToHeight=0
    wb.active=0
