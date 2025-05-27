import os
import base64
import json
import re
import io
import imghdr            # <-- เพิ่มบรรทัดนี้
import requests
import pandas as pd
import streamlit as st
from PIL import Image
from openpyxl import load_workbook

API_KEY = "AIzaSyDb8iBV1EWqLvjheG_44gh3vQHfpmYGOCI"

# --------------------------------------
#  Utilities
# --------------------------------------
def encode_image(file) -> tuple[str, str]:
    """Convert an uploaded image to (base64, mime-type)."""
    raw = file.getvalue()
    kind = imghdr.what(None, raw) or "jpeg"
    mime = f"image/{kind}"
    return base64.b64encode(raw).decode("utf-8"), mime

def _kv_from_text(txt: str) -> float | None:
    txt_u = txt.upper()
    best = None
    for chunk in re.split(r"[\/,;]", txt_u):
        if re.search(r"\bK?VA\b|\bKA\b|\bA\b|\bAMP", chunk): continue
        if "BIL" in chunk or "IMPULSE" in chunk:            continue
        for m in re.finditer(r"(\d+(?:\.\d+)?)\s*([K]?V)(?![A-Z])", chunk):
            n, unit = float(m.group(1)), m.group(2).upper()
            kv = n if unit=="KV" else n/1000
            if kv>1500: continue
            best = kv if best is None else max(best, kv)
    return best

def split_value_unit(raw: object) -> tuple[str, str]:
    s = str(raw or "").strip()
    # case: starts with number
    m = re.match(r"^(-?\d+(?:\.\d+)?)(.*)$", s)
    if m:
        val, unit = m.group(1), m.group(2).strip()
        if "/" in unit or "-" in unit:
            return s, ""
        return val, unit
    # case: percent
    m2 = re.match(r"^(-?\d+(?:\.\d+)?)\s*%$", s)
    if m2: return m2.group(1), "%"
    # case: unit suffix
    m3 = re.match(r"^(-?\d+(?:\.\d+)?)\s*([°%A-Za-zµΩ]+)$", s)
    if m3: return m3.group(1), m3.group(2)
    # else text/fallback
    return s, ""

def extract_data_from_image(api_key: str, img_b64: str, mime: str, prompt: str) -> str:
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        "gemini-2.5-flash-preview-04-17:generateContent"
        f"?key={api_key}"
    )
    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inlineData": {"mimeType": mime, "data": img_b64}}
            ]
        }],
        "generationConfig": {"temperature":0.2, "topP":0.85, "maxOutputTokens":9000}
    }
    r = requests.post(url, headers={"Content-Type":"application/json"}, data=json.dumps(payload))
    if r.ok and r.json().get("candidates"):
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]
    return f"API ERROR {r.status_code}: {r.text}"

# --------------------------------------
#  POWTR-CODE generator & validator
# --------------------------------------
def generate_powtr_code(d: dict) -> str:
    phase = "3"
    if any("1PH" in str(v).upper() or "SINGLE" in str(v).upper() for v in d.values()):
        phase = "1"
    high_kv = None
    for k,v in d.items():
        if any(t in k.upper() for t in ("VOLT","HV","LV","RATED","SYSTEM")):
            kv = _kv_from_text(str(v))
            if kv is not None:
                high_kv = kv if high_kv is None else max(high_kv, kv)
    if   high_kv is None: v_char="-"
    elif high_kv>765:     return "POWTR \\ POWTR-3-OO"
    elif high_kv>=345:    v_char="E"
    elif high_kv>=100:    v_char="H"
    elif high_kv>=1:      v_char="M"
    else:                 v_char="L"

    t_char="-"
    for v in d.values():
        u=str(v).upper()
        if "DRY" in u: t_char="D"; break
        if any(o in u for o in ("OIL","ONAN","OFAF","OA","FOA")):
            t_char="O"; break

    tap_char="F"
    for v in d.values():
        u=str(v).upper()
        if any(x in u for x in ("ON-LOAD","OLTC")):
            tap_char="O"; break

    code = f"POWTR-{phase}{v_char}{t_char}{tap_char}"
    return f"POWTR \\ {code}"

def add_powtr_codes(results: list[dict]) -> list[dict]:
    for r in results:
        d = r.get("extracted_data", {})
        if isinstance(d, dict) and not any(k in d for k in ("error","raw_text")):
            d["POWTR_CODE"] = generate_powtr_code(d)
    return results

def is_positive_oltc(val: object) -> bool:
    if pd.isna(val): return False
    v = str(val).strip().lower()
    neg = {"","-","n/a","na","none","null","no","fixed","0","off"}
    if v in neg: return False
    return "oltc" in v or "on-load" in v or v in {"y","yes"}

def validate_powtr_code(row: pd.Series) -> pd.Series:
    current = str(row.get("Classification","")).strip()
    phase = "3"
    if any("1PH" in str(v).upper() or "SINGLE" in str(v).upper() for v in row.values):
        phase="1"
    high_v=None
    for c,v in row.items():
        if "voltage" in str(c).lower():
            m = re.search(r"(\d+\.?\d*)", str(v))
            if m:
                x = float(m.group(1))
                high_v = x if high_v is None or x>high_v else high_v
    if   high_v is None:  v_char="-"
    elif high_v>765:      return pd.Series([current=="POWTR-3-OO","POWTR-3-OO"])
    elif high_v>=345:     v_char="E"
    elif high_v>=100:     v_char="H"
    elif high_v>=1:       v_char="M"
    else:                 v_char="L"
    t_char="-"
    for _,v in row.items():
        u=str(v).upper()
        if "DRY" in u: t_char="D"; break
        if any(o in u for o in ("OIL","ONAN","OFAF")): t_char="O"; break
    tap = "O" if any(is_positive_oltc(row[c]) for c in row.index if "oltc" in str(c).lower()) else "F"
    code = f"POWTR-{phase}{v_char}{t_char}{tap}"
    return pd.Series([current==code, code])

def process_excel(df: pd.DataFrame) -> pd.DataFrame:
    res = df.apply(validate_powtr_code, axis=1)
    df["Is_Correct"] = res[0]
    df["Correct_POWTR_CODE"] = res[1]
    if "Classification" in df.columns:
        cols = list(df.columns)
        cols.remove("Is_Correct"); cols.remove("Correct_POWTR_CODE")
        idx = cols.index("Classification")+1
        cols[idx:idx] = ["Is_Correct","Correct_POWTR_CODE"]
        df = df[cols]
    return df

def fill_template_from_validated(validated_path, template_path,
                                 key_col_validated="Correct_POWTR_CODE",
                                 sheet_name="AssetAttr") -> io.BytesIO:
    df = pd.read_excel(validated_path)
    df = df[df["Is_Correct"]==True].set_index(key_col_validated)
    wb = load_workbook(template_path)
    ws = wb[sheet_name]
    ws.delete_rows(1)
    header = next(ws.iter_rows(min_row=1, max_row=1))
    cols = {c.value: c.column for c in header if c.value}
    for row in ws.iter_rows(min_row=2):
        a = row[cols["ASSETNUM"]-1].value
        if a in df.index:
            rec = df.loc[a]
            row[cols["ASSETNUM"]-1].value      = rec.name
            row[cols["SITEID"]-1].value        = rec["Plant"]
            row[cols["HIERARCHYPATH"]-1].value = rec["Location Description"]
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

# --------------------------------------
#  Streamlit UI
# --------------------------------------
st.set_page_config(page_title="Transformer Toolkit", layout="wide")
tab1, tab2, tab3, tab4 = st.tabs([
    "สกัดจากรูปภาพ",
    "ประมวลผลจาก validated",
    "🔎 ตรวจสอบ POWTR-CODE",
    "สกัด NAMEPLATE อะไรก็ได้จากรูปภาพ"
])

# --- Tab 1: extract from image ---
with tab1:
    st.subheader("💡 สกัดข้อมูลจากรูปภาพ")
    attr_file = st.file_uploader("อัปโหลด Attributes (xlsx)", type=["xlsx","xls"], key="tab1_attr")
    imgs = st.file_uploader("อัปโหลดรูปภาพ (หลายไฟล์)", type=["jpg","png","jpeg"],
                             accept_multiple_files=True, key="tab1_imgs")
    if st.button("ประมวลผลภาพ", key="btn_tab1") and attr_file and imgs:
        prompt = generate_prompt_from_excel(attr_file)
        st.expander("Prompt").write(prompt)
        results=[]; bar=st.progress(0); status=st.empty()
        for i,f in enumerate(imgs,1):
            bar.progress(i/len(imgs))
            status.write(f"Processing {i}/{len(imgs)} – {f.name}")
            b64,mime=encode_image(f)
            resp=extract_data_from_image(API_KEY,b64,mime,prompt)
            try: js=json.loads(resp[resp.find("{"):resp.rfind("}")+1])
            except: js={"error":resp}
            results.append({"file_name":f.name,"extracted_data":js})
        results=add_powtr_codes(results)
        rows=[]
        for r in results:
            d=r["extracted_data"]; fn=r["file_name"]
            asset,site,code = d.get("ASSETNUM",""), d.get("SITEID",""), d.get("POWTR_CODE","")
            if "error" in d:
                rows.append({"file_name":fn,"ASSETNUM":asset,"SITEID":site,"POWTR_CODE":code,"ATTRIBUTE":"Error","VALUE":d["error"]})
            else:
                for k,v in d.items():
                    if k in ("ASSETNUM","SITEID","POWTR_CODE"): continue
                    rows.append({"file_name":fn,"ASSETNUM":asset,"SITEID":site,"POWTR_CODE":code,"ATTRIBUTE":k,"VALUE":v})
        df1=pd.DataFrame(rows)
        st.dataframe(df1)
        buf=io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w: df1.to_excel(w,index=False)
        buf.seek(0)
        st.download_button("ดาวน์โหลด extracted_long.xlsx", buf, "extracted_long.xlsx")

# --- Tab 2: from validated ---
with tab2:
    st.subheader("🔍 ประมวลผลจาก validated")
    attr2 = st.file_uploader("Attributes Excel (xlsx)", type=["xlsx","xls"], key="tab2_attr")
    val_file = st.file_uploader("validated_powtr_codes.xlsx", type=["xlsx"], key="tab2_val")
    if st.button("ประมวลผล validated", key="btn_tab2") and attr2 and val_file:
        df_attr2=pd.read_excel(attr2)
        canonical=df_attr2[df_attr2.columns[0]].dropna().astype(str).tolist()
        dfv=pd.read_excel(val_file)
        dfv=dfv[dfv["Is_Correct"]==True]
        st.dataframe(dfv)
        rows=[]; prev=None
        for _,r in dfv.iterrows():
            asset=r.get("Location","")
            plant=r.get("Plant","")
            site=(plant[:3]+"0") if plant else ""
            code=r.get("Correct_POWTR_CODE","")
            if prev and asset!=prev:
                rows.append({k:"" for k in ["file_name","ASSETNUM","SITEID","POWTR_CODE","ATTRIBUTE","VALUE","MEASUREUNIT"]})
            prev=asset
            for attr in canonical:
                raw=r.get(attr,"-")
                if attr.strip().lower().startswith("serial"):
                    val,unit=str(raw).strip(),""
                else:
                    val,unit=split_value_unit(raw)
                rows.append({
                    "file_name":r.get("Plant",""),
                    "ASSETNUM":asset,"SITEID":site,
                    "POWTR_CODE":code,"ATTRIBUTE":attr,
                    "VALUE":val,"MEASUREUNIT":unit
                })
        df2=pd.DataFrame(rows)
        st.dataframe(df2)
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="openpyxl") as w: df2.to_excel(w,index=False)
        buf.seek(0)
        st.download_button("ดาวน์โหลด extracted_long_from_validated.xlsx", buf,"extracted_long_from_validated.xlsx")

# --- Tab 3: POWTR validator ---
with tab3:
    st.subheader("🔎 ตรวจสอบ POWTR-CODE")
    uploaded=st.file_uploader("Upload Excel to validate", type=["xlsx","xls"], key="tab3_upl")
    if uploaded:
        df=pd.read_excel(uploaded)
        res=process_excel(df)
        st.dataframe(res)
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="openpyxl") as w: res.to_excel(w,index=False)
        buf.seek(0)
        st.download_button("Download validated_powtr_codes.xlsx", buf,"validated_powtr_codes.xlsx")

# --- Tab 4: free-form nameplate extraction ---
with tab4:
    st.subheader("🖼️ สกัด NAMEPLATE อะไรก็ได้จากรูปภาพ")
    imgs4=st.file_uploader("อัปโหลดรูป Nameplate", type=["jpg","png","jpeg"],
                           accept_multiple_files=True, key="tab4_imgs")
    if st.button("ประมวลผล Nameplate", key="btn_tab4") and imgs4:
        prompt4=(
            "กรุณาสกัดข้อมูลทั้งหมดจากป้ายประจำเครื่อง (nameplate) ในรูปนี้ "
            "แล้วจัดให้อยู่ใน JSON โดยใช้ key เป็น field ภาษาอังกฤษ และ value เป็นค่าที่อ่านได้"
        )
        st.expander("Prompt Nameplate").write(prompt4)
        rows=[]
        for f in imgs4:
            b64,mime=encode_image(f)
            resp=extract_data_from_image(API_KEY,b64,mime,prompt4)
            try: js=json.loads(resp[resp.find("{"):resp.rfind("}")+1])
            except: js={"error":resp}
            if isinstance(js,dict):
                for k,v in js.items():
                    rows.append({"file_name":f.name,"attribute":k,"value":v})
            else:
                rows.append({"file_name":f.name,"attribute":"error","value":js})
        df4=pd.DataFrame(rows)
        st.dataframe(df4)
        buf=io.BytesIO()
        with pd.ExcelWriter(buf,engine="openpyxl") as w: df4.to_excel(w,index=False)
        buf.seek(0)
        st.download_button("ดาวน์โหลด nameplate_extracted.xlsx", buf,"nameplate_extracted.xlsx")
