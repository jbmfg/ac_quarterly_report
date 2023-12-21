import re
import json
import requests
import sqlite_connector

def _normalize_json(data):
    if not data:
        result = [None, None]
    else:
        data = [data] if isinstance(data, dict) else data
        all_keys = list(set([k for row in data for k in row.keys()]))
        nones = set([k for row in data for k in row.keys() if not row[k]])
        for x, row in enumerate(data):
            for key in list(row):
                if isinstance(row[key], dict):
                    row[key] = str(row[key])
                elif isinstance(row[key], list):
                    if data[x][key]:
                        data[x][key] = ", ".join(row[key][0])
                    else:
                        data[x][key] = ""
        result = [data, nones]
    return result

def _normalize_url(url):
    url = re.sub(r"http.{1}:\/\/", "", url)
    url = f"https://{url}"
    return url

def get_data(server, api_key):
    headers = {
        "X-Auth-Token": api_key,
        "Content-Type": "application/json; charset=utf-8"
    }
    server = _normalize_url(server)
    r = requests.get(server, headers=headers)
    if r.status_code == 200:
        r = _normalize_json(r.json())[0]
    return r

def create_table_txt(name, data):
    text = []
    text.append(f"class {name.title()}(Base):\n")
    text.append(f'    __tablename__ = "{name}"\n')
    data, nones = _normalize_json(data)
    if not data: return
    for x, i in enumerate(data[-1].keys()):
        extra = "primary_key=True" if x==0 else ""
        dt = type(data[0][i])
        if i in nones:
            dt = "str"
            text.append(f'    {i}: sqlalchemy.orm.Mapped[typing.Optional[{dt}]] = sqlalchemy.orm.mapped_column({extra})\n')
        else:
            dt = str(dt)
            dt = dt.replace("<class \'", "").replace("\'>", "")
            text.append(f'    {i}: sqlalchemy.orm.Mapped[typing.Optional[{dt}]] = sqlalchemy.orm.mapped_column({extra})\n')
    text.append("\n\n")
    with open("tables.txt", "a") as f:
        for line in text:
            f.write(line)
    return text

if __name__ == "__main__":
    db = sqlite_connector.sqlite_db("ben.db")
    import os
    for i in os.listdir('fake_data'):
        filename = os.path.join("fake_data", i)
        tablename = i.replace(".json", "")
        with open(filename, "r") as f:
            print(filename)
            data = json.load(f)
            create_table_txt(tablename, data)
