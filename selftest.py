#!/usr/bin/env python3
import pandas as pd
from etl_core import *

info={"id":"v","channel_id":"crc","duration":40,"was_live":True,"live_status":"was_live"}
tr=pd.DataFrame([[0,10,"00:00:00.000","00:00:10.000","fala 1","youtube"],[10,20,"00:00:10.000","00:00:20.000","fala 2","youtube"],[20,40,"00:00:20.000","00:00:40.000","fala 3","youtube"]],columns=RAW_TRANS_COLS)
raw=pd.DataFrame([[1,"00:00:01.000",1_000_000,"textMessage","Mesmo Nome","a","Oi","","","m1"],[2,"00:00:02.000",2_000_000,"textMessage","Mesmo Nome","b","Olá","","","m2"]],columns=RAW_CHAT_COLS)
ct,st=clean_chat(raw,info);tt=treat_transcript(tr,"v");tl=build_timeline(tt,ct,"v")
assert st["participantes_externos_unicos"]==2, "Participantes devem ser contados por IDs distintos, mesmo com o mesmo nome exibido"
assert st["autores_unicos_chat"]==2
states={"metadados":{"status":"ok","mandatory":True},"transcricao":{"status":"ok","mandatory":True},"chat":{"status":"ok","mandatory":True}}
q,status=build_quality(info,tr,tt,raw,ct,st,tl,states)
assert status=="APROVADO_PRODUCAO", q.to_string(index=False)
assert "comentarios_verificados" not in set(q["indicador"])
assert set(source_status_table(states)["fonte"])=={"Metadados","Transcrição","Chat ao vivo"}
# Live sem chat deve reprovar.
empty=pd.DataFrame(columns=RAW_CHAT_COLS);ct2,st2=clean_chat(empty,info);tl2=build_timeline(tt,ct2,"v")
states2={"metadados":{"status":"ok","mandatory":True},"transcricao":{"status":"ok","mandatory":True},"chat":{"status":"falhou","mandatory":True}}
q2,status2=build_quality(info,tr,tt,empty,ct2,st2,tl2,states2)
assert status2=="NAO_APROVADO_PRODUCAO"
# Conta institucional CRCSP é separada do público, mesmo com ID próprio.
raw3=pd.DataFrame([[3,"00:00:03.000",3_000_000,"textMessage","@Equipe_DES_CRCSP","team1","Aviso da equipe","","","m3"],[4,"00:00:04.000",4_000_000,"textMessage","Canal","crc","Aviso oficial","","","m4"]],columns=RAW_CHAT_COLS)
ct3,st3=clean_chat(pd.concat([raw,raw3],ignore_index=True),info)
assert st3["autores_unicos_chat"]==4
assert st3["canal_oficial_ids"]==1
assert st3["equipe_crcsp_ids"]==1
assert st3["participantes_externos_unicos"]==2
assert st3["autores_unicos_chat"]==st3["participantes_externos_unicos"]+st3["institucionais_ids"]
print("SELFTEST OK - V7.1.0 IDENTIDADE CENTRALIZADA / FAIL-CLOSED")
