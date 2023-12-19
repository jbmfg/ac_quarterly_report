import json
import os
import api_actions
import sqlite_connector


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

def main():
    # Check config
    result = _check_config()
    if result:
        raise ConfigFileMissing(result)
    else:
        config = _open_config()
    server, api_key = config["server_address"], config["api_key"]
    db = sqlite_connector.sqlite_db("ac_exec_report.db")
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
        printProgressBar("Getting API resources...", x, len(resources))
        location = f"{server}{url}{str(limit)}"
        data = api_actions.get_data(location, api_key)
        if not data: continue
        for x in range(0, len(data), 500):
            if x + 500 <= len(data):
                db.insert_data(name, data[x:x+500])
            else:
                db.insert_data(name, data[x:len(data)])



if __name__ == "__main__":
    main()
