# CP2 – RPA CRCSP: Coleta e Análise de Transmissões do YouTube

Projeto desenvolvido em Python para automatizar a coleta, tratamento e consolidação de dados de palestras e transmissões do CRCSP publicadas no YouTube.

## Problema resolvido

A análise manual de uma transmissão exige abrir o vídeo, localizar metadados, transcrição e chat, organizar os dados e gerar indicadores. Este projeto automatiza esse fluxo e produz um arquivo Excel gerencial.

## Fluxo

1. O usuário informa o link de uma transmissão do YouTube.
2. O programa valida o link e realiza uma integração HTTP externa com `requests`.
3. O `yt-dlp` coleta metadados, legendas/transcrição e, quando disponível, replay do chat ao vivo.
4. O pipeline ETL limpa, padroniza e integra os dados.
5. O projeto executa controles de qualidade.
6. Um relatório Excel gerencial é gerado automaticamente.

## Tecnologias

- Python 3.11/3.12
- requests – integração HTTP externa
- yt-dlp – coleta de dados do YouTube
- pandas – tratamento e transformação dos dados
- openpyxl – geração/formatação do Excel
- faster-whisper – fallback opcional de transcrição

## Estrutura principal

- `youtube_palestra_etl.py`: programa principal e orquestração do RPA/ETL.
- `entrada_youtube.py`: validação e normalização do link informado.
- `integracao_web.py`: integração HTTP externa para validação do vídeo.
- `etl_core.py`: regras de limpeza, tratamento e qualidade.
- `relatorio_gerencial.py` e `relatorio_modelo.py`: construção do relatório final.
- `gerar_relatorio.py`: reprocessa um Excel técnico já coletado e recria a apresentação.
- `saida_segura.py`: publicação segura do arquivo Excel.
- `requirements.txt`: dependências do projeto.

## Instalação

No terminal do VSCode:

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Instale as dependências:

```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Execução

Modo interativo:

```bash
python youtube_palestra_etl.py
```

Cole o link quando solicitado.

Também é possível passar o link diretamente:

```bash
python youtube_palestra_etl.py "https://www.youtube.com/watch?v=3r3tAUYtpv4"
```


## Identificação de perguntas

A triagem não depende apenas do caractere `?`. O projeto também reconhece padrões linguísticos frequentes em perguntas, como `como`, `qual`, `onde`, `posso`, `devo`, `é possível`, `tem como`, `gostaria de saber`, `saber se` e `será que`. Para a aba **Perguntas de conteúdo**, a mensagem também precisa conter um macrotema profissional e não pode ter sido classificada antes como ocorrência operacional ou crítica. A classificação é automática e exploratória, por isso a mensagem original permanece disponível para conferência.

## Saída

O arquivo final segue o padrão:

```text
YouTube_<ID_DO_VIDEO>_GERENCIAL.xlsx
```

O Excel é organizado para análise de um único vídeo por execução. As primeiras abas são as de apresentação (`Dashboard`, `Resumo da transmissão`, `Participação`, `Perguntas de conteúdo`, `Picos de 5 minutos`, `Ocorrências e sugestões`, `Participantes Públicos` e `Metodologia e lacunas`). Em seguida ficam as bases técnicas de metadados, transcrição, chat, linha do tempo, qualidade da coleta e rastreabilidade. Abas intermediárias usadas apenas para cálculo permanecem ocultas para evitar duplicidade visual.

## Teste rápido

```bash
python selftest.py
```

## Observações

- O projeto trabalha apenas com conteúdo publicamente acessível pelo YouTube.
- Algumas transmissões podem não disponibilizar chat ou legendas.
- O FFmpeg é necessário somente quando o fallback do Whisper precisar processar áudio.


### Padronização dos indicadores

- **Participantes do público**: IDs/autores externos únicos; a conta oficial do canal não entra nesse total.
- **Autores únicos no chat**: quando exibido em controle técnico, pode incluir a conta oficial e é identificado explicitamente como total de autores.
- **Elogios/agradecimentos**: as abas principais e o indicador gerencial usam a mesma taxonomia principal do relatório.


### Regra de identidade dos participantes
Participantes únicos são contabilizados pelo `canal_autor_id` público do YouTube, e não pelo nome exibido. Assim, usuários diferentes com o mesmo nome continuam distintos, mudanças de nome não duplicam a mesma conta e mensagens sem ID permanecem nas mensagens, mas não entram no total de participantes únicos.

## Regra canônica de participantes

A identidade no chat é deduplicada exclusivamente por `canal_autor_id`. A versão 7.1 separa três grupos:

- `canal_oficial`: ID exatamente igual ao canal do vídeo;
- `equipe_crcsp`: conta institucional identificada pelo nome público associado ao CRCSP;
- `publico_externo`: demais IDs identificados.

Nomes exibidos não são usados como chave de pessoa. Mensagens sem ID permanecem na análise de mensagens, mas não entram no total de participantes únicos. Todas as abas gerenciais consomem a mesma métrica canônica, e `auditar_identidade.py` reprova o arquivo se houver divergência entre abas.

Para auditar um relatório já gerado:

```bash
python auditar_identidade.py caminho/do/relatorio.xlsx
```

## Fechamento final do Dashboard

O Dashboard separa explicitamente três populações do chat: **público externo**, **contas institucionais/equipe CRCSP** e **canal oficial**. Para o vídeo de validação, a reconciliação é 745 mensagens do público externo + 61 mensagens institucionais/equipe + 3 mensagens do canal oficial = 809 registros tratados. Participantes únicos do público externo são calculados exclusivamente por `canal_autor_id`.
