# Auditoria final da entrega CP2

Validação final realizada sobre o projeto e o relatório de demonstração.

## Regra auditável de participantes

- Participante único = `canal_autor_id` público e preenchido.
- Nome exibido (`autor`) não é usado como chave de identidade.
- Mensagens sem ID permanecem nas métricas de mensagens, mas não entram no total de participantes únicos.
- Conta oficial do canal é segregada do público.

No relatório de demonstração:
- 364 IDs únicos do público;
- 1 ID da conta oficial do canal;
- 365 IDs únicos de autores no chat no total.

## Testes

- SELFTEST: OK.
- Testes automatizados: 28/28 aprovados.
- Validações de conciliação do relatório: todas em OK.
- Erros de fórmula pesquisados: nenhum `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?` ou `#N/A`.

## Observação metodológica

Perguntas são sinalizadas por intenção interrogativa, com ou sem `?`, respeitando prioridade para categorias operacionais como certificado, acesso, material e problemas técnicos. Classificações automáticas são sinais para análise e não substituem validação humana.


## Identidade V7.1

A divergência histórica de 363/364 foi rastreada a uma conta adicional `@Equipe_DES_CRCSP`. O arquivo real contém 365 IDs identificados, reconciliados em 1 ID do canal oficial, 1 ID institucional/equipe e 363 IDs do público externo. A métrica **Participantes do público** usa somente os 363 IDs externos.

A função `participant_identity_metrics()` em `etl_core.py` é a fonte única. `relatorio_gerencial.py`, `relatorio_modelo.py`, `Qualidade_Coleta` e `Resumo` usam a mesma definição. O relatório contém conciliações por mensagem e por ID, e o utilitário `auditar_identidade.py` verifica as abas principais.

## Fechamento final do Dashboard

O Dashboard separa explicitamente três populações do chat: **público externo**, **contas institucionais/equipe CRCSP** e **canal oficial**. Para o vídeo de validação, a reconciliação é 745 mensagens do público externo + 61 mensagens institucionais/equipe + 3 mensagens do canal oficial = 809 registros tratados. Participantes únicos do público externo são calculados exclusivamente por `canal_autor_id`.
