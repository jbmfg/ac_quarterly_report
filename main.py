import datetime
import json
import os
import re
import requests
import dateparser
import api_actions
import sqlite_connector
from collections import defaultdict
from bs4 import BeautifulSoup

class ConfigFileMissing(Exception):
    pass

def printProgressBar (header, iteration, total, prefix = '', suffix = '', decimals = 1, length = 100, fill = '█', printEnd = "\r"):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
        printEnd    - Optional  : end character (e.g. "\r", "\r\n") (Str)
    """
    total -= 1
    if iteration == 0:
        print(header)
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filledLength = int(length * iteration // total)
    bar = fill * filledLength + '-' * (length - filledLength)
    print(f'\r{prefix} |{bar}| {percent}% {suffix}', end = printEnd)
    # Print New Line on Complete
    if iteration == total:
        print()

def _check_config():
    result = ""
    if "config.json" not in os.listdir():
        result = "No config file found"
    else:
        config = _open_config()
        undef_keys = []
        for key in config:
            if not config[key]:
                undef_keys.append(key)
        if undef_keys:
            result = f"Config file is missing the values for {', '.join(undef_keys)}"
    return result

def _open_config():
    with open("config.json", "r") as f:
        data = json.load(f)
    return data

def fake_requests():
    import os
    db = sqlite_connector.sqlite_db("ac_exec_report.db")
    for i in [f for f in os.listdir("./fake_data") if f.endswith(".json")]:
        if "tops_" in i: continue
        table = i.replace(".json", "")
        data = json.load(open(os.path.join("./fake_data", i), "r"))
        import api_actions
        data = api_actions._normalize_json(data)[0]
        db.insert_data(table, data)

def get_api_data(db, config):
    server, api_key = config["server_address"], config["api_key"]
    resources = (
        ("appTemplate", "/api/bit9platform/restricted/appTemplate?limit=", 0),
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
        ("block_events", "/api/bit9platform/v1/event?q=subtype:801&limit=", 10000),
        ("agent_config", "/api/bit9platform/restricted/agentConfig?limit=", 0),
        ("cache_checks", "/api/bit9platform/v1/event?q=subtype:426&q=timestamp>-30d&limit=", 0),
        ("oldest_event", "/api/bit9platform/v1/event?sort=timestamp&limit=", 1),
        ("event_count_30d", "/api/bit9platform/v1/event?q=timestamp>-30d&limit=", -1)
        )
    for x, resource in enumerate(resources):
        name, url, limit = resource
        db.delete_data(name)
        printProgressBar("Getting API resources...", x, len(resources))
        location = f"{server}{url}{str(limit)}"
        data = api_actions.get_data(location, api_key)
        if not data: continue
        for x in range(0, len(data), 500):
            if x + 500 <= len(data):
                db.insert_data(name, data[x:x+500])
            else:
                db.insert_data(name, data[x:len(data)])

def get_version_support_status(db):
    urls = [
        ["server", "/8.10/cb-ac-oer/GUID-21E6E704-237F-4415-8B50-DE380C6D9ECA.html"],
        ["win", "/services/cb-appc-oer-winagent-desktop/GUID-32AC0CFE-6CB3-4E96-A7B3-D9BA1F32BE0F.html"],
        ["mac", "/services/cb-appc-oer-macosagent/GUID-D6B6EEA2-8665-4938-9D27-87D100903A30.html"],
        ["lin", "/services/cb-appc-oer-linuxagent/GUID-054C2FCD-DA66-4D83-AABD-0F90EC543DAA.html"],
        ]
    levels = []
    for os, url in urls:
        url = "https://docs.vmware.com/en/VMware-Carbon-Black-App-Control" + url
        data = requests.get(url).content
        soup = BeautifulSoup(data, features="html.parser")
        table_bodies = soup.find_all("tbody")
        for i in table_bodies:
            rows = i.find_all("tr")
            row_data = []
            for row in rows:
                entries = row.find_all("td")
                items = [entry.get_text() for entry in entries]
                if items[0].lower().strip().startswith("none"):
                    continue
                if any(items):
                    levels += [[os] + items]
    for x, row in enumerate(levels):
        if len(row) == 4:
            levels[x].insert(3, None)
        if len(row) == 6:
            # the server tables include the GA date, which we dont care about
            levels[x] = row[:2] + row[3:]
            row = row[:2] + row[3:]
        for xx, ri in enumerate(row):
            # First two items are os and version
            if xx < 2: continue
            # Parse the date fields
            if ri:
                # Replace the day with 01 for items that match format %Y/%m
                if len(ri) in (6, 7):
                    ri = dateparser.parse(ri).replace(day=1)
                    # Add 31 days and then subtract the day number to get the last day of the month
                    ri = ri + datetime.timedelta(days=31)
                    ri = ri - datetime.timedelta(days=ri.day)
                else:
                    ri = dateparser.parse(ri)
                levels[x][xx] = ri.date().isoformat()
            else:
                levels[x][xx] = "Unknown"
        today = datetime.datetime.date(datetime.datetime.now())
    for i in levels:
        st = datetime.datetime.strptime(i[2], "%Y-%m-%d").date() if i[2] != "Unknown" else "Unknown"
        ex = datetime.datetime.strptime(i[3], "%Y-%m-%d").date() if i[3] != "Unknown" else "Unknown"
        eol = datetime.datetime.strptime(i[4], "%Y-%m-%d").date() if i[4] != "Unknown" else "Unknown"
        if eol != "Unknown" and today > eol:
            i.append("eol")
        elif ex != "Unknown" and today > ex:
            i.append("ex")
        elif st != "Unknown":
            i.append("st")
        else:
            i.append("Unknown")
    fields = ["os", "release", "enter_standard", "enter_extended", "enter_eol", "current_level"]
    levels = [dict(zip(fields, level)) for level in levels]
    db.delete_data("eol")
    db.insert_data("eol", levels)

def deployed_version_status(db):
    # Recursive function to get current level
    def cur_lev(os, v):
        print(v)
        while len(v) > 0:
            if v[:-1] in lu_dict[os]:
                return lu_dict[os][v[:-1]]
            else:
                return cur_lev(os, v[:-1])
        return "EOL"
    lookup = db.query_data("select os, REPLACE(release, '.x', ''), UPPER(current_level) from eol;")
    lu_dict = defaultdict(lambda: defaultdict(str))
    for os, version, level in lookup:
        lu_dict[os][version] = level
    print(json.dumps(lu_dict, indent=2))
    query = """
    select distinct agentVersion,
    substring(osShortName, 0, 4)
    from computer
    where lastPollDate >= (select max(date(lastPollDate, '-30 days')) from computer)
    order by agentVersion, osShortName;
    """
    all_versions = [list(i) for i in db.query_data(query)]
    for x, row in enumerate(all_versions):
        version, os = row
        print(f"THIS IS V: {version}, {os}")
        all_versions[x].append(cur_lev(os.lower(), version))
    fields = ["version", "os", "support_level"]
    all_versions = [dict(zip(fields, level)) for level in all_versions]
    db.delete_data("deployed_version_status")
    db.insert_data("deployed_version_status", all_versions)


def main():
    # Check config
    result = _check_config()
    if result:
        raise ConfigFileMissing(result)
    else:
        config = _open_config()
    db = sqlite_connector.sqlite_db("ac_exec_report.db")
    #get_api_data(db, config)
    #get_version_support_status(db)
    deployed_version_status(db)

if __name__ == "__main__":
    main()
    #fake_requests()
