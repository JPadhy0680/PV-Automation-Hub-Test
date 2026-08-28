from __future__ import annotations
import base64, hashlib, hmac, io
from pathlib import Path
from typing import Dict, Tuple
import pandas as pd
import requests
import streamlit as st

LISTEDNESS_COLUMNS = ["Active Ingredients", "PT", "Expectedness", "Comment"]

def norm(v):
    import re
    s=str(v or "").lower()
    s=re.sub(r"[^a-z0-9\s+\-]", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def dataframe_to_index(df: pd.DataFrame) -> Dict[Tuple[str,str], Dict[str,str]]:
    if df is None or df.empty: return {}
    cols={str(c).strip().lower():c for c in df.columns}
    ai=cols.get("active ingredients") or cols.get("drug name")
    pt=cols.get("pt")
    exp=cols.get("expectedness")
    comment=cols.get("comment")
    if not ai or not pt:
        st.warning("Listedness master requires Active Ingredients and PT columns.")
        return {}
    out={}
    for _,r in df.iterrows():
        k=(norm(r.get(ai,"")), norm(r.get(pt,"")))
        if all(k): out[k]={"Expectedness":str(r.get(exp,"Expected") if exp else "Expected").strip() or "Expected", "Comment":str(r.get(comment,"") if comment else "").strip()}
    return out

def assess(index, ingredient, pt):
    """Return stored expectedness/comment. Missing pairs are explicit, not Unexpected."""
    row = index.get((norm(ingredient), norm(pt)))
    if row is None:
        return "No exact match found", ""
    return row.get("Expectedness", ""), row.get("Comment", "")


def has_exact_match(index, ingredient, pt):
    return (norm(ingredient), norm(pt)) in index

def _secret(name, default=""):
    try: return str(st.secrets.get(name,default))
    except Exception: return default

def password_ok(password):
    return hmac.compare_digest(
        str(password or ""),
        "1403267112"
    )

def github_upsert_row(ingredient, pt, expectedness, comment):
    owner=_secret("GITHUB_OWNER"); repo=_secret("GITHUB_REPO"); branch=_secret("GITHUB_BRANCH","main")
    path=_secret("GITHUB_LISTEDNESS_PATH","data/Listedness_CX.xlsx"); token=_secret("GITHUB_TOKEN")
    if not all([owner,repo,path,token]): raise RuntimeError("GitHub secrets are incomplete.")
    url=f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    headers={"Authorization":f"Bearer {token}","Accept":"application/vnd.github+json","X-GitHub-Api-Version":"2022-11-28"}
    got=requests.get(url,headers=headers,params={"ref":branch},timeout=30); got.raise_for_status(); meta=got.json()
    raw=base64.b64decode(meta["content"])
    df=pd.read_excel(io.BytesIO(raw),engine="openpyxl")
    for c in LISTEDNESS_COLUMNS:
        if c not in df.columns: df[c]=""
    mask=(df["Active Ingredients"].map(norm)==norm(ingredient)) & (df["PT"].map(norm)==norm(pt))
    row={"Active Ingredients":ingredient.strip().upper(),"PT":pt.strip(),"Expectedness":expectedness,"Comment":comment.strip()}
    if mask.any():
        for c,v in row.items(): df.loc[mask,c]=v
    else: df=pd.concat([df,pd.DataFrame([row])],ignore_index=True)
    buf=io.BytesIO(); df.to_excel(buf,index=False,engine="openpyxl")
    payload={"message":f"Update listedness: {ingredient} / {pt}","content":base64.b64encode(buf.getvalue()).decode(),"sha":meta["sha"],"branch":branch}
    put=requests.put(url,headers=headers,json=payload,timeout=30); put.raise_for_status()
    return put.json().get("commit",{}).get("sha","")

def render_listedness_updater(key_prefix, ingredient_options, pt_options):
    with st.expander("🔐 Update Listedness Master on GitHub"):
        st.caption("Use only when the Active Ingredient + PT combination is missing or needs correction.")
        ingredient=st.selectbox("Active Ingredient", sorted(set(x for x in ingredient_options if x)), key=f"{key_prefix}_ing") if ingredient_options else st.text_input("Active Ingredient",key=f"{key_prefix}_ing")
        pt=st.selectbox("PT", sorted(set(x for x in pt_options if x)), key=f"{key_prefix}_pt") if pt_options else st.text_input("PT",key=f"{key_prefix}_pt")
        expectedness=st.selectbox("Expectedness",["Expected","Unexpected"],key=f"{key_prefix}_exp")
        comment=st.text_input("Comment",key=f"{key_prefix}_comment")
        password=st.text_input("Administrator password",type="password",key=f"{key_prefix}_pwd")
        if st.button("Commit update to GitHub",key=f"{key_prefix}_save"):
            if not password_ok(password): st.error("Invalid password or password hash is not configured.")
            elif not ingredient or not pt: st.error("Active Ingredient and PT are required.")
            else:
                try:
                    sha=github_upsert_row(ingredient,pt,expectedness,comment)
                    st.success(f"Listedness master updated. Commit: {sha[:10]}")
                    st.cache_data.clear()
                except Exception as exc: st.error(f"GitHub update failed: {exc}")


def render_missing_listedness_update(key_prefix, ingredient, pt):
    """Render update controls immediately below one unmatched listedness pair."""
    safe_key = f"{key_prefix}_{abs(hash((norm(ingredient), norm(pt))))}"
    st.warning(f"No exact listedness match found for {ingredient} + {pt}.")
    c1, c2 = st.columns(2)
    with c1:
        expectedness = st.selectbox(
            "Expectedness",
            ["Expected", "Unexpected"],
            key=f"{safe_key}_expectedness",
        )
        password = st.text_input(
            "Administrator password",
            type="password",
            key=f"{safe_key}_password",
        )
    with c2:
        comment = st.text_area(
            "Comment",
            key=f"{safe_key}_comment",
            height=100,
        )
    if st.button("Update this listedness pair", key=f"{safe_key}_save"):
        if not password_ok(password):
            st.error("Invalid administrator password.")
            return False
        try:
            sha = github_upsert_row(ingredient, pt, expectedness, comment)
            st.success(f"Listedness master updated. Commit: {sha[:10]}")
            st.cache_data.clear()
            return True
        except Exception as exc:
            st.error(f"GitHub update failed: {exc}")
    return False
