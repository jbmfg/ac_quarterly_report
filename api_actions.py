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
                    #data[x].pop(key)
                elif isinstance(row[key], list):
                    data[x][key] = ", ".join(row[key])
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

def create_table_txt():
    text = []
    for name, url, limit in resources:
        #if name != "appTemplate": continue
        r = get_data(url+str(limit))
        if r.json():
            text.append(f"class {name.title()}(Base):\n")
            text.append(f'    __tablename__ = "{name}"\n')
            data, nones = _normalize_json(r.json())
            for x, i in enumerate(data[0].keys()):
                extra = "primary_key=True" if x==0 else ""
                dt = type(data[0][i])
                if i in nones:
                    dt = "str"
                    text.append(f'    {i}: sqlalchemy.orm.Mapped[typing.Optional[{dt}]] = sqlalchemy.orm.mapped_column({extra})\n')
                else:
                    dt = str(dt)
                    dt = dt.replace("<class \'", "").replace("\'>", "")
                    text.append(f'    {i}: sqlalchemy.orm.Mapped[{dt}] = sqlalchemy.orm.mapped_column({extra})\n')
            text.append("\n\n")
    with open("tables.txt", "w") as f:
        for line in text:
            f.write(line)
    return text

if __name__ == "__main__":
    db = sqlite_connector.sqlite_db("ben.db")
    resources = (
        ("publisher", "/api/bit9platform/v1/publisher?limit=", 0),
        ("computer", "/api/bit9platform/v1/computer?q=daysOffline<32&limit=", 0),
        ("updater", "/api/bit9platform/v1/updater?limit=", 0),
        ("policy", "/api/bit9platform/v1/policy?limit=", 0),
        ("scriptRule", "/api/bit9platform/v1/scriptRule?limit=", 0),
        ("serverPerformance", "/api/bit9platform/v1/serverPerformance?limit=", 0),
        ("trustedDirectory", "/api/bit9platform/v1/trustedDirectory?limit=", 0),
        ("trustedUser", "/api/bit9platform/v1/trustedUser?limit=", 0),
        ("serverConfig", "/api/bit9platform/v1/serverConfig?limit=", 0),
        ("driftReport", "/api/bit9platform/v1/driftReport?limit=", 0),
        ("global_approval_counts", "/api/bit9platform/v1/fileRule?q=sourceType!3&filestate:2&lazyApproval:false&group=datecreated&limit=", 0),
        ("global_approval_counts_2", "/api/bit9platform/v1/fileRule?q=sourceType!3&filestate:2&lazyApproval:false&group=datecreated&grouptype=n&groupstep=1", 0),
        ("td_approval_counts", "/api/bit9platform/v1/fileRule?q=sourceType:2&group=sourceId&limit=", 0),
        ("approval_summary", "/api/bit9platform/v1/fileRule?group=sourceType&limit=", 0),
        ("rule_hits", "/api/bit9platform/v1/event?group=ruleName&limit=", 0),
        ("rapfig_events", "/api/bit9platform/v1/event?q=rapfigName!&q=timestamp>-30d&group=rapfigName&subgroup=RuleName&limit=", 0),
        ("extensions", "/api/bit9platform/v1/fileCatalog?q=dateCreated>-30d&group=fileExtension&limit=", 0),
        ("unapprovedWriters", "/api/bit9platform/v1/event?expand=fileCatalogId&q=subtype:1003&q=fileCatalogId_effectiveState:Unapproved&q=param1:DiscoveredBy[Kernel:Rename]*|DiscoveredBy[Kernel:Create]*|DiscoveredBy[Kernel:Write]*&limit=", 10000),
        ("customRule", "/api/bit9platform/restricted/customRule?limit=", 0),
        ("approvalRequestSummary", "/api/bit9platform/restricted/approvalRequestSummary?limit=", 0),
        ("appTemplate", "/api/bit9platform/restricted/appTemplate?limit=", 0),
        ("block_events", "/api/bit9platform/v1/event?q=subtype:801&limit=", 10000),
        ("agent_config", "/api/bit9platform/restricted/agentConfig?limit=", 0),
        ("cache_checks", "/api/bit9platform/v1/event?q=subtype:426&q=timestamp>-30d&limit=", 0),
        ("oldest_event", "/api/bit9platform/v1/event?sort=timestamp&limit=", 1),
        ("event_count_30d", "/api/bit9platform/v1/event?q=timestamp>-30d&limit=", -1)
        )

    #create_table_txt()
    #input("stop")
    insert_data()
