import requests
from bs4 import BeautifulSoup
import lxml
import re


def get_bbt_data(scrm_id):
    url = "https://bbt.ext.net.nokia.com/asmxservices/wsbidbox.asmx?op=GetSWFDetails"
    payload = f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
          <soap:Body>
            <GetSWFDetails xmlns="http://tempuri.org/">
                <SWFNumber>{scrm_id}</SWFNumber>
            </GetSWFDetails>
          </soap:Body>
        </soap:Envelope>"""
    # headers
    headers = {"Content-Type": "text/xml; charset=utf-8"}

    try:
        response = requests.request("POST", url, headers=headers, data=payload)
        xml = BeautifulSoup(response.text, "html.parser")
        cust = xml.find("customer").text
        market = xml.find("market").text
        opp_name = xml.find("opportunityname").text
        country = xml.find("country").text
        region = xml.find("region").text
        subregion = xml.find("subregion").text
        if market != None and len(market) > 0:
            return {
                "customer": cust,
                "market": market,
                "opportunityname": opp_name,
                "country": country,
                "region": region,
                "subregion": subregion,
            }
        else:
            return {}
        # op_name = xml.find('OpportunityName').text
    except Exception as e:
        return {}


def getParents(children, selected_sec):
    for child in children:
        if not re.search(r"0.[a-z]$", child):
            for parent_idx in range(len(child.split("."))):
                parent_section = ".".join(child.split(".")[0:parent_idx])
                if len(parent_section) > 0:
                    if parent_section in selected_sec:
                        pass
                    else:
                        selected_sec.append(parent_section)

    return selected_sec


def bson_to_json(type, collection, filename=None, domain=None, user=None):
    if type == "sections_mappings":
        return {filename: collection["sections"]}
    elif type == "master_mappings":
        return {domain: {filename: collection["mappings"]}}
    elif type == "download_history_mappings":
        return {user: {domain: {filename: collection["mappings"]}}}
    elif type == "heirarchy_mappings":
        return {filename: collection["version_heirarchy"]}
    elif type == "custom_vars_mappings":
        return {domain: {filename: collection["custom_vars"]}}


def json_to_bson(cur_json, additional_params):
    bson_payload = []
    for cur_key in cur_json.keys():
        additional_params.update({cur_key: cur_json[cur_key]})
        bson_payload.append(additional_params)
    return bson_payload


def format_data(filename=None, domain=None, user=None):
    return {
        "filter": {
            "user_file": filename,
            "user_domain": domain,
            "user_name": user,
        }
    }


def create_ver_heirarchy(sections):
    version_hierarchy = {}
    for section in sections:
        version = section
        node = {"children": []}
        sub_vers = version.split(".")
        node_ver = ".".join(sub_vers)
        if len(sub_vers) > 1:
            parent = ".".join(sub_vers[0 : len(sub_vers) - 1])
            if parent not in version_hierarchy.keys():
                version_hierarchy[parent] = {"children": []}
            version_hierarchy[parent]["children"].append(node_ver)
            version_hierarchy[node_ver] = node
            # print(f"Parent: {parent} - Child: {version}")
        else:
            version_hierarchy[node_ver] = node
    return version_hierarchy


def get_sensitive_sections(restricted_keywords, content, sections):
    restricted_keywrds_sections = []
    for section in sections:
        if (
            section["id"] in content.keys()
            and len(re.findall(restricted_keywords, content[section["id"]].lower())) > 0
        ):
            restricted_keywrds_sections.append(section["id"])
    return restricted_keywrds_sections


def get_restricted_keywrds_list(restricted_ls):
    restricted_keywords = dict(restricted_ls.find_one({"type": "legal"}))
    if restricted_keywords != None:
        restricted_keywords = restricted_keywords["keywords"]
    else:
        restricted_keywords = []
    restricted_keywords = "|".join(restricted_keywords)
    restricted_keywords = f"(?:{restricted_keywords})"
    return restricted_keywords
