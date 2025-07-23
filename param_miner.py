from org.zaproxy.zap.extension.pscan import PluginPassiveScanner
from org.zaproxy.addon.commonlib.scanrules import ScanRuleMetadata
from org.zaproxy.zap.extension.script import ScriptVars
from java.net import URLDecoder

def getMetadata():
    return ScanRuleMetadata.fromYaml("""
id: 77777
name: Collect Request Parameters with Values
description: Collects all request parameters and values from URL and body.
risk: INFO
confidence: HIGH
cweId: 0
wascId: 0
status: alpha
""")

def appliesToHistoryType(historyType):
    return PluginPassiveScanner.getDefaultHistoryTypes().contains(historyType)

def scan(helper, msg, src):
    uri = msg.getRequestHeader().getURI()
    method = msg.getRequestHeader().getMethod()
    body = msg.getRequestBody().toString()

    # Key for global storage
    global_key = "collected_params"

    # Get existing global param store
    stored = ScriptVars.getGlobalVar(global_key)
    if stored is None:
        param_map = {}
    else:
        try:
            param_map = eval(stored)
        except:
            param_map = {}

    # --- Extract from query ---
    query = uri.getQuery()
    if query:
        for pair in query.split("&"):
            if "=" in pair:
                k, v = pair.split("=", 1)
                k = URLDecoder.decode(k, "UTF-8")
                v = URLDecoder.decode(v, "UTF-8")
                param_map[k] = v
            else:
                k = URLDecoder.decode(pair, "UTF-8")
                param_map[k] = ""

    # --- Extract from body ---
    if method.upper() in ["POST", "PUT", "PATCH"]:
        for pair in body.split("&"):
            if "=" in pair:
                k, v = pair.split("=", 1)
                k = URLDecoder.decode(k, "UTF-8")
                v = URLDecoder.decode(v, "UTF-8")
                param_map[k] = v
            else:
                k = URLDecoder.decode(pair, "UTF-8")
                param_map[k] = ""

    # Store updated param map
    ScriptVars.setGlobalVar(global_key, str(param_map))
