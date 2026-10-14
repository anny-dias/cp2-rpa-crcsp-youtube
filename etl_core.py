# -*- coding: utf-8 -*-
from __future__ import annotations
import hashlib, re
from datetime import datetime, timezone
from typing import Any, Dict, Optional, Tuple
import pandas as pd

PIPELINE_VERSION = "7.1.0"
RAW_TRANS_COLS = ["inicio_segundos","fim_segundos","inicio","fim","texto","fonte"]
RAW_CHAT_COLS = ["offset_segundos","tempo_video","timestamp_usec","tipo","autor","canal_autor_id","mensagem","valor_pago","moeda","id_mensagem"]


def seconds_to_hms(seconds: Optional[float]) -> str:
    if seconds is None or pd.isna(seconds): return ""
    total_ms=int(round(float(seconds)*1000)); sign="-" if total_ms<0 else ""; total_ms=abs(total_ms)
    h,rem=divmod(total_ms,3_600_000); m,rem=divmod(rem,60_000); s,ms=divmod(rem,1000)
    return f"{sign}{h:02d}:{m:02d}:{s:02d}.{ms:03d}"


def safe_int(v):
    try:
        if pd.isna(v): return None
        return int(v)
    except Exception: return None


def safe_float(v):
    try:
        if pd.isna(v): return None
        return float(v)
    except Exception: return None


def normalize_text(t: str) -> str:
    t="" if pd.isna(t) else str(t)
    t=re.sub(r"\[(?:música|musica|aplausos?|risos?)\]"," ",t,flags=re.I)
    t=re.sub(r"^\s*>>+\s*","",t)
    return re.sub(r"\s+"," ",t).strip()


def sanitize_log_detail(text: str, max_len: int=1200) -> str:
    text=str(text or "")
    text=re.sub(r"https?://\S+","[URL_REMOVIDA]",text,flags=re.I)
    text=re.sub(r"\b(?:\d{1,3}\.){3}\d{1,3}\b","[IP_REMOVIDO]",text)
    text=re.sub(r"(?i)(signature|sig|token|expire|ip|key|auth|cookie|authorization)=([^&\s]+)",r"\1=[REMOVIDO]",text)
    text=re.sub(r"(?i)(cookie|authorization):\s*.+",r"\1: [REMOVIDO]",text)
    text=re.sub(r"\s+"," ",text).strip()
    return text[:max_len]+(" ...[TRUNCADO]" if len(text)>max_len else "")


def transcript_diagnostics(raw: pd.DataFrame, duration: float) -> Dict[str, Any]:
    out={"inicio":None,"fim":None,"span_pct":None,"covered_seconds":0.0,"coverage_pct":None,
         "start_gap":None,"end_gap":None,"max_gap":None,"gaps_gt_30":0,"gaps_gt_120":0,"intervalos":0}
    if raw.empty:return out
    pairs=[]
    for _,r in raw.iterrows():
        a=safe_float(r.get("inicio_segundos")); b=safe_float(r.get("fim_segundos"))
        if a is None: continue
        if b is None or b<a:b=a
        if duration:
            a=min(max(0.0,a),duration); b=min(max(0.0,b),duration)
        else:
            a=max(0.0,a); b=max(0.0,b)
        pairs.append((a,b))
    if not pairs:return out
    pairs.sort(); merged=[]
    for a,b in pairs:
        if not merged or a>merged[-1][1]+0.25: merged.append([a,b])
        else: merged[-1][1]=max(merged[-1][1],b)
    start,end=merged[0][0],merged[-1][1]
    gaps=[max(0.0,merged[i+1][0]-merged[i][1]) for i in range(len(merged)-1)]
    covered=sum(max(0.0,b-a) for a,b in merged)
    out.update({
        "inicio":start,"fim":end,"span_pct":round((max(0,end-start)/duration)*100,2) if duration else None,
        "covered_seconds":round(covered,3),"coverage_pct":round((covered/duration)*100,2) if duration else None,
        "start_gap":round(start,3),"end_gap":round(max(0.0,duration-end),3) if duration else None,
        "max_gap":round(max(gaps),3) if gaps else 0.0,"gaps_gt_30":sum(g>30 for g in gaps),
        "gaps_gt_120":sum(g>120 for g in gaps),"intervalos":len(merged)})
    return out


def canonical_author_type(author: Any, author_id: Any, channel_id: str = "") -> str:
    """Classifica o autor sem usar nome como substituto do ID do canal oficial."""
    a="" if pd.isna(author) else str(author).strip()
    aid="" if pd.isna(author_id) else str(author_id).strip()
    cid=str(channel_id or "").strip()
    low=a.lower()
    if cid and aid==cid:
        return "canal_oficial"
    if "crcsp" in low or "crc-sp" in low or "crc_sp" in low:
        return "equipe_crcsp"
    if not a:
        return "sem_autor"
    return "publico_externo"


def participant_identity_metrics(df: pd.DataFrame, channel_id: str = "") -> Dict[str, int]:
    """Retorna métricas canônicas de identidade do chat.

    A chave de pessoa é sempre ``canal_autor_id``. Nomes exibidos nunca são
    usados para deduplicar participantes. Contas institucionais reconhecidas
    pelo ETL (``canal_oficial`` e ``equipe_crcsp``) são separadas do público.
    Registros sem ID permanecem nas contagens de mensagens, mas não contam como
    participantes únicos.
    """
    base = df.copy() if df is not None else pd.DataFrame()
    if base.empty or "canal_autor_id" not in base.columns:
        return {
            "autores_identificados_total": 0,
            "canal_oficial_ids": 0,
            "equipe_crcsp_ids": 0,
            "institucionais_ids": 0,
            "participantes_publico_ids": 0,
            "sem_id_mensagens": 0,
        }
    ids = base["canal_autor_id"].fillna("").astype(str).str.strip()
    valid = ids.ne("")
    autores = base.get("autor", pd.Series([""] * len(base), index=base.index))
    # Recalcula a tipologia sempre pela regra canônica. Não confia em uma coluna
    # autor_tipo herdada de versões antigas do arquivo.
    tipos = pd.Series([canonical_author_type(a, aid, channel_id) for a, aid in zip(autores, ids)], index=base.index)
    canal_mask = tipos.eq("canal_oficial")
    equipe_mask = tipos.eq("equipe_crcsp")
    institucional = canal_mask | equipe_mask
    publico = valid & ~institucional
    return {
        "autores_identificados_total": int(ids[valid].nunique()),
        "canal_oficial_ids": int(ids[valid & canal_mask].nunique()),
        "equipe_crcsp_ids": int(ids[valid & equipe_mask].nunique()),
        "institucionais_ids": int(ids[valid & institucional].nunique()),
        "participantes_publico_ids": int(ids[publico].nunique()),
        "sem_id_mensagens": int((~valid).sum()),
    }


def clean_chat(raw: pd.DataFrame, info: Dict[str, Any]) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    stats={"chat_bruto":len(raw),"duplicidades_removidas":0,"mensagens_sistema":0,"mensagens_pre_evento":0,
           "mensagens_durante_evento":0,"mensagens_pos_evento":0,"chat_valido":0,"participantes_unicos":0,
           "participantes_externos_unicos":0,"offset_anchors":0,"offset_mad_ms":None,"offset_confiavel":False}
    cols=["video_id","pipeline_version","chave_evento","offset_segundos","tempo_video","periodo","timestamp_usec","timestamp_utc","tipo","autor","autor_tipo","canal_autor_id","mensagem","valor_pago","moeda","id_mensagem","offset_original_segundos"]
    if raw.empty:return pd.DataFrame(columns=cols),stats
    df=raw.copy();df["offset_original_segundos"]=pd.to_numeric(df["offset_segundos"],errors="coerce")
    seen_ids=set();seen_fallback=set();keep=[]
    for _,r in df.iterrows():
        mid="" if pd.isna(r.get("id_mensagem")) else str(r.get("id_mensagem")).strip()
        if mid:key=("id",mid);dup=key in seen_ids;seen_ids.add(key)
        else:key=("fallback",str(r.get("timestamp_usec") or ""),str(r.get("autor") or ""),str(r.get("mensagem") or ""),str(r.get("tipo") or ""));dup=key in seen_fallback;seen_fallback.add(key)
        keep.append(not dup);stats["duplicidades_removidas"]+=int(dup)
    df=df.loc[keep].copy();bases=[]
    for _,r in df.iterrows():
        ts=safe_int(r.get("timestamp_usec"));off=safe_float(r.get("offset_original_segundos"))
        if ts is not None and off is not None and off>.25 and str(r.get("tipo"))!="sistema":bases.append(ts-off*1_000_000)
    zero_usec=int(pd.Series(bases).median()) if bases else None;stats["offset_anchors"]=len(bases)
    if bases:
        med=float(pd.Series(bases).median());mad=float(pd.Series([abs(x-med) for x in bases]).median())
        stats["offset_mad_ms"]=round(mad/1000,3);stats["offset_confiavel"]=len(bases)>=3 and mad<=5_000_000
    corr=[];utc=[]
    for _,r in df.iterrows():
        ts=safe_int(r.get("timestamp_usec"));off=safe_float(r.get("offset_original_segundos"))
        if ts is not None and zero_usec is not None and (off is None or off<=.001):off=(ts-zero_usec)/1_000_000
        corr.append(off);utc.append(datetime.fromtimestamp(ts/1_000_000,tz=timezone.utc).isoformat(timespec="milliseconds") if ts else "")
    df["offset_segundos"]=corr;df["tempo_video"]=[seconds_to_hms(x) for x in corr];df["timestamp_utc"]=utc
    duration=safe_float(info.get("duration")) or 0
    def period(x):
        if x is None or pd.isna(x):return "indeterminado"
        if x<0:return "pre_evento"
        if duration and x>duration:return "pos_evento"
        return "durante_evento"
    df["periodo"]=[period(x) for x in corr]
    cid=str(info.get("channel_id") or info.get("uploader_id") or "")
    df["autor_tipo"]=[canonical_author_type(a,b,cid) for a,b in zip(df["autor"],df["canal_autor_id"])]
    stats["mensagens_sistema"]=int((df["tipo"].astype(str)=="sistema").sum());t=df.loc[df["tipo"].astype(str)!="sistema"].copy();vid=str(info.get("id") or "video")
    t["video_id"]=vid;t["pipeline_version"]=PIPELINE_VERSION;keys=[]
    for _,r in t.iterrows():
        mid="" if pd.isna(r.get("id_mensagem")) else str(r.get("id_mensagem")).strip();material="|".join([vid,str(r.get("timestamp_usec") or ""),str(r.get("autor") or ""),str(r.get("mensagem") or ""),str(r.get("tipo") or "")])
        keys.append(f"{vid}:{mid}" if mid else f"{vid}:sha256:{hashlib.sha256(material.encode()).hexdigest()[:24]}")
    t["chave_evento"]=keys
    stats["mensagens_pre_evento"]=int((t["periodo"]=="pre_evento").sum());stats["mensagens_durante_evento"]=int((t["periodo"]=="durante_evento").sum());stats["mensagens_pos_evento"]=int((t["periodo"]=="pos_evento").sum());stats["chat_valido"]=len(t)
    # Uma única função canônica alimenta todas as métricas de identidade.
    identity=participant_identity_metrics(t,cid)
    stats["autores_unicos_chat"]=identity["autores_identificados_total"]
    stats["canal_oficial_ids"]=identity["canal_oficial_ids"]
    stats["equipe_crcsp_ids"]=identity["equipe_crcsp_ids"]
    stats["institucionais_ids"]=identity["institucionais_ids"]
    stats["participantes_externos_unicos"]=identity["participantes_publico_ids"]
    stats["participantes_unicos"]=identity["participantes_publico_ids"]
    return t[cols].sort_values(["offset_segundos","timestamp_usec"],na_position="last",kind="stable").reset_index(drop=True),stats


def treat_transcript(raw: pd.DataFrame, video_id: str) -> pd.DataFrame:
    cols=["video_id","pipeline_version","chave_bloco","inicio_segundos","fim_segundos","inicio","fim","duracao_segundos","segmentos","texto","fonte"]
    if raw.empty:return pd.DataFrame(columns=cols)
    rows=[];texts=[];st=en=None;cnt=0;src=""
    def emit():
        nonlocal texts,st,en,cnt,src
        if texts and st is not None:
            txt=re.sub(r"\s+"," "," ".join(texts)).strip();material=f"{video_id}|{st:.3f}|{en:.3f}|{txt}"
            rows.append([video_id,PIPELINE_VERSION,f"{video_id}:transcript:{hashlib.sha256(material.encode()).hexdigest()[:24]}",st,en,seconds_to_hms(st),seconds_to_hms(en),max(0,en-st),cnt,txt,src])
        texts=[];st=en=None;cnt=0;src=""
    for _,r in raw.sort_values("inicio_segundos",na_position="last",kind="stable").iterrows():
        txt=normalize_text(r.get("texto",""));a=safe_float(r.get("inicio_segundos"));b=safe_float(r.get("fim_segundos"))
        if not txt or a is None:continue
        b=a if b is None else b
        if st is None:st,en,src=a,b,str(r.get("fonte") or "")
        projected=len(" ".join(texts+[txt]));gap=0 if en is None else a-en
        if texts and (b-st>60 or projected>900 or gap>5):emit();st,en,src=a,b,str(r.get("fonte") or "")
        texts.append(txt);en=max(en if en is not None else b,b);cnt+=1
    emit();return pd.DataFrame(rows,columns=cols)


def build_timeline(tt: pd.DataFrame,ct: pd.DataFrame,video_id: str) -> pd.DataFrame:
    rows=[]
    for _,r in tt.iterrows():rows.append([video_id,PIPELINE_VERSION,r["chave_bloco"],r["inicio_segundos"],r["inicio"],"durante_evento","transcricao","fala","","",r["texto"]])
    for _,r in ct.iterrows():rows.append([video_id,PIPELINE_VERSION,r["chave_evento"],r["offset_segundos"],r["tempo_video"],r["periodo"],"chat_ao_vivo",r["tipo"],r["autor"],r["autor_tipo"],r["mensagem"]])
    cols=["video_id","pipeline_version","registro_id","tempo_segundos","tempo","periodo","origem","tipo","autor","autor_tipo","conteudo"]
    return pd.DataFrame(rows,columns=cols).sort_values(["tempo_segundos","origem","registro_id"],na_position="last",kind="stable").reset_index(drop=True) if rows else pd.DataFrame(columns=cols)


def source_status_table(source_states: Dict[str,Any]) -> pd.DataFrame:
    rows=[]
    label={"metadados":"Metadados","transcricao":"Transcrição","chat":"Chat ao vivo"}
    for key in ["metadados","transcricao","chat"]:
        x=source_states.get(key,{})
        status=x.get("status","indeterminado")
        rows.append([label[key],status.upper(),"SIM" if x.get("mandatory") else "NAO",x.get("method","") or x.get("source",""),x.get("detail","")])
    return pd.DataFrame(rows,columns=["fonte","status","obrigatoria","metodo","detalhe"])


def build_quality(info,tr,tt,cr,ct,stats,timeline,source_states=None,raw_expected=None):
    source_states=source_states or {};dur=safe_float(info.get("duration")) or 0;td=transcript_diagnostics(tr,dur)
    rawids=cr["id_mensagem"].fillna("").astype(str).str.strip() if "id_mensagem" in cr else pd.Series(dtype=str);raw_dup=int(rawids[rawids.ne("")].duplicated().sum())
    pre=int((ct["periodo"]=="pre_evento").sum()) if not ct.empty else 0;during=int((ct["periodo"]=="durante_evento").sum()) if not ct.empty else 0;post=int((ct["periodo"]=="pos_evento").sum()) if not ct.empty else 0;period_sum=pre+during+post
    key_dups=int(ct["chave_evento"].duplicated().sum()) if "chave_evento" in ct else 0;system_treated=int((ct["tipo"].astype(str)=="sistema").sum()) if "tipo" in ct else 0
    neg_ok=all(str(t).startswith("-") for o,t in zip(ct.get("offset_segundos",[]),ct.get("tempo_video",[])) if pd.notna(o) and float(o)<0);sorted_ok=bool(pd.to_numeric(timeline["tempo_segundos"],errors="coerce").dropna().is_monotonic_increasing) if not timeline.empty else True
    raw_preserved=True
    if raw_expected:raw_preserved=(len(tr)==raw_expected.get("transcricao",len(tr)) and len(cr)==raw_expected.get("chat",len(cr)))
    chat_expected=bool(info.get("was_live") or info.get("live_status") in ("was_live","is_live"));chat_verified=(not chat_expected) or len(cr)>0;metadata_ok=bool(info.get("id")) and dur>0
    transcript_ok=(not tr.empty and dur>0 and td["fim"] is not None and td["fim"]>=dur*.95 and (td["inicio"] or 0)<=max(120,dur*.05) and (td["max_gap"] is None or td["max_gap"]<=900) and (td["coverage_pct"] is None or td["coverage_pct"]>=60))
    offset_ok=(len(cr)==0) or bool(stats.get("offset_confiavel")) or not pd.to_numeric(cr.get("offset_segundos",pd.Series(dtype=float)),errors="coerce").fillna(0).eq(0).any()
    checks={"metadados_validos":metadata_ok,"transcricao_obtida":not tr.empty,"transcricao_cobertura_temporal":transcript_ok,"chat_obtido_quando_esperado":chat_verified,"chat_offset_confiavel":offset_ok,"chat_periodos_fecham":period_sum==len(ct),"chat_chaves_unicas":key_dups==0,"chat_sem_sistema":system_treated==0,"tempo_negativo_formatado":neg_ok,"linha_tempo_ordenada":sorted_ok,"dados_brutos_preservados":raw_preserved}
    for k,v in source_states.items():
        if isinstance(v,dict) and v.get("mandatory") and v.get("status") in {"falhou","indeterminado","truncado"}:checks[f"fonte_{k}_ok"]=False
    failed=[k for k,v in checks.items() if not v];status="APROVADO_PRODUCAO" if not failed else "NAO_APROVADO_PRODUCAO";warnings=[]
    if raw_dup:warnings.append(f"{raw_dup} duplicidade(s) da fonte bruta removidas somente no analítico")
    if stats.get("mensagens_sistema",0):warnings.append("Mensagens de sistema preservadas no bruto e excluídas do analítico")
    if td.get("coverage_pct") is not None and td["coverage_pct"]<80:warnings.append("Cobertura efetiva da legenda abaixo de 80%; pode haver silêncio ou trechos sem legenda")
    rows=[["pipeline_version",PIPELINE_VERSION,"Versão do ETL"],["status_geral",status,"Gate fail-closed"],["checks_falhos"," | ".join(failed) if failed else "Nenhum","Qualquer falha impede uso automático"],["dados_brutos_preservados","SIM" if raw_preserved else "NAO","Camadas brutas não sobrescritas"],["duracao_video_segundos",dur,"Metadado"],["transcricao_segmentos_brutos",len(tr),"Fonte integral"],["transcricao_blocos_tratados",len(tt),"Camada NLP"],["transcricao_inicio_segundos",td["inicio"],"Primeiro segmento"],["transcricao_fim_segundos",td["fim"],"Último segmento"],["transcricao_span_pct",td["span_pct"],"(fim-início)/duração"],["transcricao_cobertura_efetiva_pct",td["coverage_pct"],"União dos intervalos legendados / duração"],["transcricao_segundos_cobertos",td["covered_seconds"],"União dos intervalos legendados"],["transcricao_gap_inicial_segundos",td["start_gap"],"Tempo até primeiro segmento"],["transcricao_gap_final_segundos",td["end_gap"],"Tempo após último segmento"],["transcricao_maior_gap_segundos",td["max_gap"],"Maior lacuna interna"],["transcricao_gaps_maiores_30s",td["gaps_gt_30"],"Diagnóstico"],["transcricao_gaps_maiores_120s",td["gaps_gt_120"],"Diagnóstico"],["chat_registros_brutos",len(cr),"Fonte integral"],["chat_duplicidades_id_fonte",raw_dup,"Bruto preservado"],["chat_duplicidades_removidas",stats.get("duplicidades_removidas",0),"Somente analítico"],["chat_offset_anchors",stats.get("offset_anchors",0),"Âncoras timestamp-offset"],["chat_offset_mad_ms",stats.get("offset_mad_ms"),"Dispersão das âncoras"],["chat_offset_confiavel","SIM" if offset_ok else "NAO","Validação temporal"],["chat_mensagens_pre_evento",pre,"Tabela final"],["chat_mensagens_durante_evento",during,"Tabela final"],["chat_mensagens_pos_evento",post,"Tabela final"],["chat_registros_validos",len(ct),"Tabela final"],["chat_chaves_evento_duplicadas",key_dups,"Deve ser zero"],["chat_autores_unicos_total",stats.get("autores_unicos_chat",0),"IDs únicos de autores no chat; inclui contas institucionais e público; exclui registros sem ID"],["chat_canal_oficial_ids",stats.get("canal_oficial_ids",0),"IDs que coincidem exatamente com o canal do vídeo"],["chat_equipe_crcsp_ids",stats.get("equipe_crcsp_ids",0),"IDs de contas institucionais identificadas pelo nome público CRCSP, separadas do canal oficial"],["chat_institucionais_ids",stats.get("institucionais_ids",0),"União de canal oficial e contas institucionais/equipe"],["chat_participantes_publico_unicos",stats.get("participantes_externos_unicos",0),"IDs únicos do público externo; exclui canal oficial, contas institucionais/equipe e registros sem ID"],["linha_tempo_ordenada","SIM" if sorted_ok else "NAO","Crescente"],["alertas_informativos"," | ".join(warnings) if warnings else "Nenhum","Alertas não substituem checks"]]
    return pd.DataFrame(rows,columns=["indicador","valor","criterio"]),status
