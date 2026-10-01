
import streamlit as st
import pandas as pd
import numpy as np
import io, os, re, tempfile, subprocess, shutil
import openpyxl
from datetime import datetime

st.set_page_config(page_title="Bank Statement → Tally Accounting Voucher", page_icon="🏦", layout="wide")

TARGET_HEADERS = [
    "Voucher Date","Voucher Type Name","Voucher Number",
    "Buyer/Supplier - Address","Buyer/Supplier - Pincode",
    "Ledger Name","Ledger Amount","Ledger Amount Dr/Cr",
    "Item Name","Billed Quantity","Item Rate","Item Rate per",
    "Item Amount","Change Mode "
]

DATE_ALIASES = ["date","txn date","transaction date","transaction dt","tran date","value date","value dt","posting date","posting dt","entry date"]
NARR_ALIASES = ["narration","description","particulars","transaction particulars","remarks","details","transaction details","memo"]
REF_ALIASES = ["chq no","cheque no","cheque number","chq./ref.no.","chq/ref","ref no","reference no","reference","utr","utr no","transaction id","txn id"]
DEBIT_ALIASES = ["debit","debit amount","withdrawal","withdrawal amt","withdrawal amount","debit amt","dr","dr amount","paid out","amount debited"]
CREDIT_ALIASES = ["credit","credit amount","deposit","deposit amt","deposit amount","credit amt","cr","cr amount","paid in","amount credited"]
AMOUNT_ALIASES = ["amount","transaction amount","txn amount"]

def norm(s):
    return re.sub(r"[^a-z0-9]+","",str(s).strip().lower())

def money(x):
    if x is None or (isinstance(x,float) and np.isnan(x)): return np.nan
    s=str(x).strip()
    if not s: return np.nan
    neg=False
    if s.startswith("(") and s.endswith(")"): neg=True
    s=re.sub(r"[₹$€£,\s]", "", s)
    s=re.sub(r"(?i)\binr\b", "", s)
    s=re.sub(r"[^0-9.\-]", "", s)
    if not s or s in ("-","."): return np.nan
    try:
        v=float(s)
        return -abs(v) if neg else v
    except: return np.nan

def decrypt_if_needed(data, password):
    # Try as-is first. If encrypted, decrypt using msoffcrypto.
    if not password:
        return data
    try:
        import msoffcrypto
        src=io.BytesIO(data)
        office=msoffcrypto.OfficeFile(src)
        if office.is_encrypted():
            office.load_key(password=password)
            out=io.BytesIO()
            office.decrypt(out)
            return out.getvalue()
    except Exception:
        pass
    return data

def read_excel_bytes(data, name, password=""):
    raw=decrypt_if_needed(data,password)
    ext=os.path.splitext(name.lower())[1]
    frames={}
    errors=[]
    engines=[]
    if ext==".xlsx":
        engines=["openpyxl"]
    elif ext==".xls":
        engines=["xlrd", "openpyxl"]
    else:
        engines=["openpyxl", "xlrd"]
    for engine in engines:
        try:
            xls=pd.ExcelFile(io.BytesIO(raw), engine=engine)
            for sheet in xls.sheet_names:
                try:
                    frames[sheet]=pd.read_excel(io.BytesIO(raw),sheet_name=sheet,header=None,engine=engine)
                except Exception as e:
                    errors.append(f"{sheet}: {e}")
            if frames: return frames, errors
        except Exception as e:
            errors.append(str(e))
    raise ValueError("Could not read the Excel workbook. If it is password protected, enter the correct password. Details: " + " | ".join(errors[-3:]))

def read_csv_bytes(data):
    for enc in ["utf-8-sig","utf-8","cp1252","latin1"]:
        try:
            txt=data.decode(enc)
            return {"CSV": pd.read_csv(io.StringIO(txt), header=None, sep=None, engine="python")}
        except Exception:
            pass
    raise ValueError("Could not read the CSV file.")

def read_pdf_bytes(data):
    try:
        import pdfplumber
    except ImportError:
        raise ValueError("PDF support requires pdfplumber. Install requirements.txt.")
    tables=[]
    texts=[]
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for page in pdf.pages:
            try:
                for t in (page.extract_tables() or []):
                    if t: tables.append(pd.DataFrame(t))
            except Exception: pass
            try:
                txt=page.extract_text()
                if txt: texts.append(txt)
            except Exception: pass
    frames={}
    for i,t in enumerate(tables):
        frames[f"PDF Table {i+1}"]=t
    if frames: return frames
    if texts:
        rows=[]
        for txt in texts:
            for line in txt.splitlines():
                parts=re.split(r"\s{2,}|\t",line.strip())
                if len(parts)>=3: rows.append(parts)
        if rows: frames["PDF Text"]=pd.DataFrame(rows)
    if not frames:
        raise ValueError("No readable table/text was found in this PDF. If it is scanned, OCR is required.")
    return frames

def read_upload(upload, password=""):
    """Read an uploaded statement and ALWAYS return {sheet_name: dataframe}."""
    data = upload.getvalue()
    ext = os.path.splitext(upload.name.lower())[1]
    if ext in [".xlsx", ".xls", ".xlsm"]:
        result = read_excel_bytes(data, upload.name, password)
        # read_excel_bytes returns (frames, errors); normalize here so callers
        # never accidentally iterate over the tuple itself.
        if isinstance(result, tuple):
            frames, errors = result
        else:
            frames, errors = result, []
        if not isinstance(frames, dict):
            raise ValueError("Excel reader returned an invalid workbook structure.")
        return frames
    if ext == ".csv":
        return read_csv_bytes(data)
    if ext == ".pdf":
        return read_pdf_bytes(data)
    raise ValueError("Supported formats: XLSX, XLS, CSV, PDF.")

def find_header_and_map(df):
    if df is None or df.empty: return None
    scan=min(60,len(df))
    best=None
    for r in range(scan):
        vals=[norm(x) for x in df.iloc[r].tolist()]
        score=0
        idx={}
        def locate(aliases):
            for j,v in enumerate(vals):
                if not v: continue
                for a in aliases:
                    na=norm(a)
                    if v==na or na in v or v in na:
                        return j
            return None
        d=locate(DATE_ALIASES); n=locate(NARR_ALIASES)
        db=locate(DEBIT_ALIASES); cr=locate(CREDIT_ALIASES); am=locate(AMOUNT_ALIASES); rf=locate(REF_ALIASES)
        if d is not None: score+=3
        if n is not None: score+=3
        if db is not None: score+=2
        if cr is not None: score+=2
        if am is not None: score+=1
        if rf is not None: score+=1
        if score>=6 and (db is not None or cr is not None or am is not None):
            best=(r,{"date":d,"narr":n,"debit":db,"credit":cr,"amount":am,"ref":rf},score)
    return best

def parse_frame(df, source):
    found=find_header_and_map(df)
    if not found: return pd.DataFrame()
    h,m,score=found
    rows=[]
    body=df.iloc[h+1:].copy()
    for _,row in body.iterrows():
        date=row.iloc[m["date"]] if m["date"] is not None else None
        narration=row.iloc[m["narr"]] if m["narr"] is not None else ""
        ref=row.iloc[m["ref"]] if m["ref"] is not None else ""
        debit=money(row.iloc[m["debit"]]) if m["debit"] is not None else np.nan
        credit=money(row.iloc[m["credit"]]) if m["credit"] is not None else np.nan
        amount=money(row.iloc[m["amount"]]) if m["amount"] is not None else np.nan
        if pd.isna(debit): debit=0.0
        if pd.isna(credit): credit=0.0
        if debit==0 and credit==0 and not pd.isna(amount):
            # If a single Amount column exists, use sign to determine side.
            if amount < 0: debit=abs(amount)
            elif amount > 0: credit=amount
        parsed_date=pd.to_datetime(date,errors="coerce",dayfirst=True)
        text=str(narration).strip()
        if pd.isna(parsed_date) or not text or text.lower() in ("nan","none","total","opening balance","closing balance"):
            continue
        if debit==0 and credit==0: continue
        if debit<0: debit=abs(debit)
        if credit<0: credit=abs(credit)
        # Avoid rows with both sides unless the source truly has both.
        if debit>0 and credit>0:
            if debit>=credit: credit=0.0
            else: debit=0.0
        rows.append({
            "Date":parsed_date,
            "Narration":text,
            "Reference":"" if str(ref).lower()=="nan" else str(ref).strip(),
            "Withdrawal":round(float(debit),2),
            "Deposit":round(float(credit),2),
            "Source":source
        })
    return pd.DataFrame(rows)

def parse_statement(frames):
    # Defensive normalization for older callers / cached Streamlit state.
    if isinstance(frames, tuple):
        frames = frames[0]
    if not isinstance(frames, dict):
        raise ValueError("Invalid workbook data. Please re-upload the statement.")
    allparts=[]
    for name,df in frames.items():
        p=parse_frame(df,name)
        if not p.empty: allparts.append(p)
    if not allparts: return pd.DataFrame()
    out=pd.concat(allparts,ignore_index=True)
    out=out.drop_duplicates(subset=["Date","Narration","Reference","Withdrawal","Deposit"])
    return out.sort_values(["Date"]).reset_index(drop=True)

def load_mapping(upload,password=""):
    if upload is None: return []
    frames=read_upload(upload,password)
    if isinstance(frames, tuple):
        frames=frames[0]
    for _,df in frames.items():
        if df.empty: continue
        # Use first row as header
        hdr=[str(x).strip() for x in df.iloc[0].tolist()]
        data=df.iloc[1:].copy(); data.columns=hdr
        cols={norm(c):c for c in data.columns}
        k=next((cols[norm(x)] for x in ["Keyword","Narration Keyword","Match"] if norm(x) in cols),None)
        l=next((cols[norm(x)] for x in ["Ledger","Ledger Name","Tally Ledger"] if norm(x) in cols),None)
        if k and l:
            rules=[]
            for _,r in data.iterrows():
                kw=str(r[k]).strip()
                led=str(r[l]).strip()
                if kw and led and kw.lower()!="nan" and led.lower()!="nan":
                    rules.append((kw,led))
            return sorted(rules,key=lambda x:len(x[0]),reverse=True)
    return []

def map_ledger(narration,rules,default):
    n=narration.lower()
    for kw,led in rules:
        if kw.lower() in n: return led, "Matched"
    return default, "Suspense/Unmatched"

def build_target(transactions, bank_ledger, default_ledger, rules):
    rows=[]
    review=[]
    for i,t in transactions.iterrows():
        vnum=f"BANK-{i+1:06d}"
        vtype="Payment" if t["Withdrawal"]>0 else "Receipt"
        amt=t["Withdrawal"] if t["Withdrawal"]>0 else t["Deposit"]
        contra,status=map_ledger(t["Narration"],rules,default_ledger)
        date=t["Date"].strftime("%d-%b-%Y")
        ref=t["Reference"]
        # Two ledger rows per voucher: counter ledger first, bank second.
        if vtype=="Payment":
            sides=[(contra,amt,"Dr"),(bank_ledger,amt,"Cr")]
        else:
            sides=[(bank_ledger,amt,"Dr"),(contra,amt,"Cr")]
        for ledger,amount,dc in sides:
            rows.append([
                date,vtype,vnum,"","",ledger,round(amount,2),dc,
                "","","","","",""
            ])
        review.append({
            "Voucher Number":vnum,"Date":date,"Voucher Type":vtype,
            "Narration":t["Narration"],"Reference":ref,
            "Ledger":contra,"Amount":amt,"Match Status":status
        })
    return pd.DataFrame(rows,columns=TARGET_HEADERS),pd.DataFrame(review)

def make_xlsx(df, review, template_bytes):
    # Preserve exact Accounting Voucher headers and add review as a separate sheet.
    wb=openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws=wb["Accounting Voucher"]
    # clear rows below header
    if ws.max_row>1: ws.delete_rows(2,ws.max_row-1)
    for row in df.itertuples(index=False,name=None):
        ws.append(list(row))
    if "Conversion Review" in wb.sheetnames:
        del wb["Conversion Review"]
    rv=wb.create_sheet("Conversion Review")
    if not review.empty:
        rv.append(list(review.columns))
        for row in review.itertuples(index=False,name=None): rv.append(list(row))
    out=io.BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

st.title("🏦 Bank Statement → Tally Accounting Voucher Converter")
st.caption("Converts bank statements into the exact Accounting Voucher worksheet structure supplied by you.")
st.caption("Converter build: FIXED v3 — Excel reader normalized")

with st.sidebar:
    st.header("Tally Settings")
    bank_ledger=st.text_input("Bank Ledger Name","SBI")
    default_ledger=st.text_input("Unmatched Ledger","Suspense")
    password=st.text_input("Excel Password",type="password")
    st.divider()
    st.caption("Optional ledger mapping")
    mapping_file=st.file_uploader("Ledger Master (optional)",type=["xlsx","xls","csv"],key="mapping")
    st.caption("Mapping columns: Keyword, Ledger")

statement=st.file_uploader("Upload Bank Statement",type=["xlsx","xls","xlsm","csv","pdf"])
if statement:
    try:
        with st.spinner("Reading statement..."):
            frames=read_upload(statement,password)
            tx=parse_statement(frames)
        if tx.empty:
            st.error("No transactions could be detected. Check the password, file type, or statement layout.")
        else:
            st.success(f"Detected {len(tx)} transactions.")
            st.dataframe(tx,use_container_width=True)
            try:
                rules=load_mapping(mapping_file,password) if mapping_file else []
            except Exception as e:
                rules=[]; st.warning(f"Ledger master could not be read; using Suspense for unmatched rows. {e}")
            target,review=build_target(tx,bank_ledger,default_ledger,rules)
            st.subheader("Tally Accounting Voucher Preview")
            st.dataframe(target,use_container_width=True)
            st.subheader("Conversion Review")
            st.dataframe(review,use_container_width=True)
            template_path=os.path.join(os.path.dirname(__file__),"AccountingVouchers_Tally_Template.xlsx")
            with open(template_path,"rb") as f: template_bytes=f.read()
            xbytes=make_xlsx(target,review,template_bytes)
            st.download_button("⬇️ Download Tally Accounting Voucher Excel",xbytes,
                file_name="AccountingVouchers_Converted.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            st.info("Each bank transaction becomes one Tally voucher with two ledger rows: bank ledger and counter/suspense ledger. Unmatched transactions remain mapped to Suspense and are listed in Conversion Review.")
    except Exception as e:
        st.error(str(e))
else:
    st.info("Upload a bank statement to begin.")
