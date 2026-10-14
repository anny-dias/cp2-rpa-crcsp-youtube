"""Taxonomia local, reproduzível e exploratória. Sem revisão humana implícita."""
import re
import unicodedata

MACROS = [
 ('Tributação e Reforma Tributária', r'\b(ibs|cbs|tribut\w*|reforma tributaria|icms|iss|ipi|simples nacional|aliquota\w*|credito presumido|ncm)\b'),
 ('Outros temas profissionais', r'\b(profission\w*|conselho profissional|registro profissional|educacao continuada)\b'),
 ('Perícia, societário e prestação de contas', r'\b(pericia\w*|perito\w*|societari\w*|prestacao de contas|dissolucao|contrato social)\b'),
 ('IRPF, DIRF e obrigações fiscais', r'\b(irpf|dirf|ecf|ecd|dctf\w*|esocial|reinf|sped|declaracao de renda|obrigac\w* fisc\w*)\b'),
 ('Carreira, gestão e novos negócios', r'\b(carreira\w*|gestao|negocio\w*|empreend\w*|lideranca|honorario\w*|precificacao|soft skills)\b'),
 ('Tecnologia, IA e futuro da profissão', r'\b(ia|inteligencia artificial|tecnologi\w*|automatiza\w*|automacao|chatgpt|software\w*)\b'),
 ('Contabilidade, demonstrações e auditoria', r'\b(contabil\w*|auditoria\w*|auditor\w*|balanco\w*|demonstrac\w*|ifrs|cpc|nbc|ativo\w*|passivo\w*|depreciacao)\b'),
 ('Fraude, ética e PLD', r'\b(fraude\w*|etica|pld|lavagem de dinheiro|coaf|compliance)\b'),
 ('Temas sociais, trabalhistas e terceiro setor', r'\b(autismo|inclusao|diversidade|trabalhista\w*|nr[ -]?1|psicossocia\w*|terceiro setor|clt|sindicato\w*)\b'),
]
MACRO_NAMES = [name for name,_ in MACROS]
CATEGORY_NAMES = ['Pergunta sobre o conteúdo','Presença, inscrição ou certificado',
 'Material, gravação ou contato','Problema técnico ou acesso','Horário ou início da transmissão',
 'Cumprimento, boas vindas','Comentário sobre o conteúdo','Sugestão/crítica',
 'Elogio/agradecimento','Outro/indeterminado']
OPERATIONS = [
 ('Problema técnico ou acesso','Áudio/conexão/imagem',r'\b(audio|som|microfone|conexao|trav\w*|login|portal|nao (?:consigo )?(?:ouvir|escutar|acessar)|sem imagem|sem acesso)\b'),
 ('Presença, inscrição ou certificado','Certificado/EPC',r'\b(certificad\w*|epc|pontuacao|pontos)\b'),
 ('Presença, inscrição ou certificado','Acesso/presença/inscrição',r'\b(presenca|inscri\w*|lista de presenca)\b'),
 ('Material, gravação ou contato','Material/gravação',r'\b(material|materiais|slides?|apostila\w*|gravacao|gravado|pdf)\b'),
 ('Material, gravação ou contato','Contato do palestrante',r'\b(contato|email|e-mail|instagram|telefone|whatsapp)\b'),
 ('Horário ou início da transmissão','Acesso/presença/inscrição',r'\b(horario|comec\w*|inicio|iniciar\w*|atraso\w*|evento correto|live certa)\b'),
 ('Problema técnico ou acesso','Acesso/presença/inscrição',r'\b(acesso|link)\b'),
]
QUESTION = r'(?:\?|\b(?:como|qual|quais|quando|onde|por que|porque|quanto\w*|quem|duvida\w*)\b|\b(?:e possivel|tem como|gostaria de saber|queria saber|quero saber|alguem sabe|saber se|sera que|ha possibilidade|haveria possibilidade)\b|\b(?:posso|podemos|devo|devemos)\b|\b(?:pode|podem|deve|devem)\s+(?:ser|usar|utilizar|fazer|considerar|aplicar|informar|emitir|lancar|calcular|deduzir|compensar|registrar|entregar|enviar|adotar|reconhecer|contabilizar)\b)'
CRITIC = r'\b(sugiro|sugest\w*|critica\w*|ruim|pessim\w*|deveria\w*|poderia\w*)\b'
THANKS = r'\b(obrigad\w*|agradec\w*|parabens|excelente\w*|otim\w*|maravilh\w*)\b'
GREETING = r'\b(bom dia|boa tarde|boa noite|ola|oi|bem vind\w*)\b'


def normalize(value):
 return ''.join(c for c in unicodedata.normalize('NFKD',str(value or '').lower()) if not unicodedata.combining(c))


def classify_message(message):
 value=normalize(message)
 matches=[(name, sorted(set(re.findall(pattern,value)))) for name,pattern in MACROS if re.search(pattern,value)]
 # Um macrotema por mensagem; empate pela ordem pública de regras.
 topic=matches[0][0] if matches else 'Não identificado'
 result={'categoria_modelo':'Outro/indeterminado','categoria_detalhada':'Outro/indeterminado',
         'macrotema_modelo':'Não se aplica','termos_macrotema':'; '.join(matches[0][1]) if matches else '',
         'pergunta_conteudo':False,'comentario_conteudo':False,'sugestao_critica':False,
         'substantiva':False,'operacional_modelo':False,'regra_modelo':'Sem correspondência'}
 for category,fine,pattern in OPERATIONS:
  if re.search(pattern,value):
   result.update(categoria_modelo=category,categoria_detalhada=fine,operacional_modelo=fine!='Contato do palestrante',regra_modelo=pattern)
   return result
 if re.search(CRITIC,value):
  result.update(categoria_modelo='Sugestão/crítica',categoria_detalhada='Crítica/sugestão',
                sugestao_critica=True,substantiva=True,macrotema_modelo=topic,regra_modelo=CRITIC)
 elif matches and re.search(QUESTION,value):
  result.update(categoria_modelo='Pergunta sobre o conteúdo',categoria_detalhada='Pergunta conteúdo',
                pergunta_conteudo=True,substantiva=True,macrotema_modelo=topic,regra_modelo=QUESTION)
 elif re.search(THANKS,value):
  result.update(categoria_modelo='Elogio/agradecimento',categoria_detalhada='Elogio/agradecimento',regra_modelo=THANKS)
 elif re.search(GREETING,value):
  result.update(categoria_modelo='Cumprimento, boas vindas',categoria_detalhada='Cumprimento',regra_modelo=GREETING)
 elif matches and len(value.split())>=4:
  result.update(categoria_modelo='Comentário sobre o conteúdo',categoria_detalhada='Comentário conteúdo',
                comentario_conteudo=True,substantiva=True,macrotema_modelo=topic,regra_modelo='Termo temático e pelo menos 4 palavras; sem intenção interrogativa detectada')
 return result


def enrich_analysis(a):
 # Não modifica o tratamento V7 nem os indicadores da camada anterior.
 from collections import Counter
 a=dict(a)
 a['analysis']=[dict(r,**classify_message(r.get('mensagem'))) for r in a['analysis']]
 a['public']=[r for r in a['analysis'] if r['populacao'] in {'Público externo','Público externo sem ID'}]
 a['official']=[r for r in a['analysis'] if r['populacao']=='Canal oficial']
 a['counts']=Counter(r['categoria_detalhada'] for r in a['public'])
 return a
