import unittest
from classificacao_modelo import classify_message,MACRO_NAMES

class ClassificacaoTest(unittest.TestCase):
 def test_operational_precedes_topic(self):
  row=classify_message('Onde está o certificado da palestra de auditoria?')
  self.assertEqual(row['categoria_modelo'],'Presença, inscrição ou certificado')
  self.assertFalse(row['pergunta_conteudo'])
 def test_questions_and_comments(self):
  row=classify_message('Como calcular a alíquota do IBS?')
  self.assertTrue(row['pergunta_conteudo'])
  self.assertEqual(row['macrotema_modelo'],'Tributação e Reforma Tributária')
  row=classify_message('A auditoria exige avaliar evidências e riscos relevantes.')
  self.assertTrue(row['comentario_conteudo'])
  self.assertTrue(row['substantiva'])
 def test_no_theme_from_nowhere(self):
  row=classify_message('Quem está aqui hoje?')
  self.assertFalse(row['pergunta_conteudo'])
  self.assertEqual(row['macrotema_modelo'],'Não se aplica')
 def test_exclusive_criticism(self):
  row=classify_message('Sugiro aprofundar a auditoria de estimativas?')
  self.assertTrue(row['sugestao_critica'])
  self.assertFalse(row['pergunta_conteudo'])
  self.assertTrue(row['substantiva'])
 def test_greeting_and_nine_groups(self):
  self.assertEqual(len(MACRO_NAMES),9)
  row=classify_message('Bom dia, pessoal da contabilidade!')
  self.assertEqual(row['categoria_modelo'],'Cumprimento, boas vindas')
  self.assertFalse(row['substantiva'])
 def test_contact_separate_from_operational_card(self):
  row=classify_message('Qual o contato do palestrante?')
  self.assertEqual(row['categoria_modelo'],'Material, gravação ou contato')
  self.assertFalse(row['operacional_modelo'])

 def test_questions_without_question_mark(self):
  examples = [
   'Gostaria de saber se o IBS permite crédito nessa operação',
   'É possível utilizar crédito de ICMS nessa situação',
   'Posso aplicar essa regra de auditoria na empresa',
   'Devo contabilizar esse ativo dessa forma',
   'Tem como calcular a alíquota do IBS neste caso',
   'Será que o CPC permite esse reconhecimento contábil',
  ]
  for message in examples:
   with self.subTest(message=message):
    row=classify_message(message)
    self.assertTrue(row['pergunta_conteudo'])
    self.assertTrue(row['substantiva'])

 def test_question_without_mark_still_respects_operational_priority(self):
  row=classify_message('Gostaria de saber onde encontro o certificado')
  self.assertEqual(row['categoria_modelo'],'Presença, inscrição ou certificado')
  self.assertFalse(row['pergunta_conteudo'])

if __name__=='__main__':unittest.main()
