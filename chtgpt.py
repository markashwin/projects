{
    "type": "function",
    "function": {
      "name": "zap_ascan_scan",
      "description": "Starts an active scan against the given URL.",
      "parameters": {
        "type": "object",
        "properties": {
          "url": {"type": "string", "description": "(positional or keyword)"},
          "recurse": {"type": "string", "description": "(optional)"},
          "in_scope_only": {"type": "string", "description": "(optional)"}
        },
        "required": ["url"]
      },
      "x-original-path": "zap.ascan.scan"
    }
  }
