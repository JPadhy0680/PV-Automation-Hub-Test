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
    row=index.get((norm(ingredient),norm(pt)))
    return (row or {}).get("Expectedness","Unexpected"), (row or {}).get("Comment","")

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


def github_upsert_rows(rows):
    """Add or update multiple listedness pairs in one GitHub commit."""
    owner = _secret("GITHUB_OWNER")
    repo = _secret("GITHUB_REPO")
    branch = _secret("GITHUB_BRANCH", "main")
    path = _secret("GITHUB_LISTEDNESS_PATH", "data/Listedness_CX.xlsx")
    token = _secret("GITHUB_TOKEN")
    if not all([owner, repo, path, token]):
        raise RuntimeError("GitHub secrets are incomplete.")

    url = f"https://api.github.com/repos/{owner}/{repo}/contents/{path}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    got = requests.get(url, headers=headers, params={"ref": branch}, timeout=30)
    got.raise_for_status()
    meta = got.json()
    raw = base64.b64decode(meta["content"])
    df = pd.read_excel(io.BytesIO(raw), engine="openpyxl")

    for column in LISTEDNESS_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    added = 0
    updated = 0
    for item in rows:
        ingredient = str(item.get("Active Ingredients", "")).strip()
        pt = str(item.get("PT", "")).strip()
        expectedness = str(item.get("Expectedness", "")).strip()
        comment = str(item.get("Comment", "")).strip()
        if not ingredient or not pt or expectedness not in {"Expected", "Unexpected"}:
            continue

        mask = (
            (df["Active Ingredients"].map(norm) == norm(ingredient))
            & (df["PT"].map(norm) == norm(pt))
        )
        row = {
            "Active Ingredients": ingredient.upper(),
            "PT": pt,
            "Expectedness": expectedness,
            "Comment": comment,
        }
        if mask.any():
            for column, value in row.items():
                df.loc[mask, column] = value
            updated += 1
        else:
            df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            added += 1

    if added == 0 and updated == 0:
        raise ValueError("No completed listedness rows were provided.")

    buffer = io.BytesIO()
    df.to_excel(buffer, index=False, engine="openpyxl")
    payload = {
        "message": f"Batch update listedness: {added} added, {updated} updated",
        "content": base64.b64encode(buffer.getvalue()).decode(),
        "sha": meta["sha"],
        "branch": branch,
    }
    put = requests.put(url, headers=headers, json=payload, timeout=30)
    put.raise_for_status()
    sha = put.json().get("commit", {}).get("sha", "")
    return {"added": added, "updated": updated, "sha": sha}


def has_exact_match(index, ingredient, pt):
    """Return True only when the exact normalized Active Ingredient + PT pair exists."""
    return (norm(ingredient), norm(pt)) in index


def render_missing_listedness_update(key_prefix, ingredient, pt):
    """Show one inline update form for a missing listedness pair."""
    safe_key = f"{key_prefix}_{abs(hash((norm(ingredient), norm(pt))))}"
    st.warning(f"No exact listedness match found for {ingredient} + {pt}.")
    expectedness = st.selectbox(
        "Expectedness",
        ["Expected", "Unexpected"],
        key=f"{safe_key}_expectedness",
    )
    comment = st.text_input("Comment", key=f"{safe_key}_comment")
    password = st.text_input(
        "Administrator password",
        type="password",
        key=f"{safe_key}_password",
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
