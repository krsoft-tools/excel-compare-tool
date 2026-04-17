from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import HTMLResponse
import pandas as pd

app = FastAPI()


@app.get("/")
def root():
    return {"message": "API is running"}


# 👉 normalization
def normalize_value(val):
    if pd.isna(val) or val == "":
        return ""
    try:
        f = float(val)
        if f.is_integer():
            return str(int(f))
        return str(f)
    except:
        return str(val).strip()


# 👉 CORE LOGIC
async def compare_files_logic(file1, file2, key):
    df1 = pd.read_excel(file1.file)
    df2 = pd.read_excel(file2.file)

    df1 = df1.loc[:, ~df1.columns.str.contains('^Unnamed')]
    df2 = df2.loc[:, ~df2.columns.str.contains('^Unnamed')]

    if key not in df1.columns or key not in df2.columns:
        return {"error": f"Column '{key}' not found"}

    # remove empty rows
    df1 = df1[df1[key].notna()]
    df2 = df2[df2[key].notna()]

    df1 = df1[df1[key].astype(str).str.strip() != ""]
    df2 = df2[df2[key].astype(str).str.strip() != ""]

    # normalize
    df1[key] = df1[key].apply(normalize_value)
    df2[key] = df2[key].apply(normalize_value)

    for col in df1.columns:
        df1[col] = df1[col].apply(normalize_value)
    for col in df2.columns:
        df2[col] = df2[col].apply(normalize_value)

    # missing
    keys1 = set(df1[key])
    keys2 = set(df2[key])

    missing_in_file2 = sorted(list(keys1 - keys2))
    missing_in_file1 = sorted(list(keys2 - keys1))

    # merge
    merged = df1.merge(df2, on=key, how="inner", suffixes=("_1", "_2"))

    differences = []

    for _, row in merged.iterrows():
        diff = {}

        for col in df1.columns:
            if col == key:
                continue

            val1 = row[f"{col}_1"]
            val2 = row[f"{col}_2"]

            if val1 != val2:
                diff[col] = {
                    "file1": val1,
                    "file2": val2
                }

        if diff:
            differences.append({
                "id": row[key],
                "differences": diff
            })

    return {
        "rows_compared": len(merged),
        "differences_found": len(differences),
        "differences": differences,
        "missing_in_file2": missing_in_file2,
        "missing_in_file1": missing_in_file1
    }


# 👉 API endpoint
@app.post("/compare")
async def compare_files(
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
    key: str = Form(...)
):
    return await compare_files_logic(file1, file2, key)


# 👉 UI PAGE
@app.get("/ui", response_class=HTMLResponse)
def ui():
    return """
    <html>
    <head>
        <title>Excel Compare</title>
        <style>
            body {
                font-family: Arial;
                background: #f5f7fa;
                display: flex;
                justify-content: center;
                padding-top: 50px;
            }
            .card {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 10px 30px rgba(0,0,0,0.1);
                width: 400px;
            }
            input, button {
                width: 100%;
                margin-bottom: 10px;
                padding: 10px;
            }
            button {
                background: #4CAF50;
                color: white;
                border: none;
                cursor: pointer;
            }
        </style>
    </head>
    <body>
        <div class="card">
            <h2>Compare Excel Files</h2>
            <form action="/compare-ui" method="post" enctype="multipart/form-data">
                <input type="file" name="file1">
                <input type="file" name="file2">
                <input type="text" name="key" placeholder="Key column">
                <button type="submit">Compare</button>
            </form>
        </div>
    </body>
    </html>
    """


# 👉 UI RESULT
@app.post("/compare-ui", response_class=HTMLResponse)
async def compare_ui(file1: UploadFile = File(...), file2: UploadFile = File(...), key: str = Form(...)):
    result = await compare_files_logic(file1, file2, key)

    if "error" in result:
        return f"<h3>{result['error']}</h3>"

    html = """
    <html>
    <head>
    <style>
        body { font-family: Arial; background: #f5f7fa; padding: 40px; }
        .container { max-width: 800px; margin: auto; }
        .card {
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 10px; border-bottom: 1px solid #eee; }
        th { background: #fafafa; }
        .diff { background: #ffe5e5; }
        .missing {
            padding: 10px;
            margin: 5px 0;
            background: #fff2cc;
            border-left: 5px solid orange;
        }
        .ok { color: #888; }
    </style>
    </head>
    <body>
    <div class="container">
    """

    html += f"""
    <div class="card">
        <h2>Result</h2>
        <p>Rows compared: <b>{result['rows_compared']}</b></p>
        <p>Differences: <b>{result['differences_found']}</b></p>
    </div>
    """

    # 👉 differences
    html += "<div class='card'><h3>Differences</h3>"

    if not result["differences"]:
        html += "<p class='ok'>No differences found ✅</p>"
    else:
        html += "<table><tr><th>ID</th><th>Column</th><th>File1 (baseline)</th><th>File2</th></tr>"

        for d in result["differences"]:
            for col, vals in d["differences"].items():
                html += f"""
                <tr>
                    <td>{d['id']}</td>
                    <td>{col}</td>
                    <td class='diff'><b>{vals['file1']}</b></td>
                    <td class='diff'>{vals['file2']}</td>
                </tr>
                """

        html += "</table>"

    html += "</div>"

    # 👉 missing file2
    html += "<div class='card'><h3>Missing in file2</h3>"
    if result["missing_in_file2"]:
        for m in result["missing_in_file2"]:
            html += f"<div class='missing'>Missing ID: <b>{m}</b></div>"
    else:
        html += "<p class='ok'>None</p>"
    html += "</div>"

    # 👉 missing file1
    html += "<div class='card'><h3>Missing in file1</h3>"
    if result["missing_in_file1"]:
        for m in result["missing_in_file1"]:
            html += f"<div class='missing'>Missing ID: <b>{m}</b></div>"
    else:
        html += "<p class='ok'>None</p>"
    html += "</div>"

    html += "</div></body></html>"

    return html