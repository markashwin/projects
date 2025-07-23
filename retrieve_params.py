"""
Standalone scripts have no template.
They are only evaluated when you run them. 
"""

from org.zaproxy.zap.extension.script import ScriptVars

stored = ScriptVars.getGlobalVar("collected_params")

# Collecting stored params
if stored is None:
    print("No parameters collected.")
else:
    try:
        param_map = eval(stored)
        print("\n--- Collected Request Parameters ---")
        for key in param_map:
            value = param_map[key]
            print("{} = {}".format(key, value))
    except:
        print("Error parsing parameter map.")



