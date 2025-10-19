import inspect
from cai.sdk.agents.tool import function_tool

def expose_zap_methods(zap_instance):
    tools = []
    
    for section_name in dir(zap_instance):
        section = getattr(zap_instance, section_name)
        if not hasattr(section, "__dict__"):
            continue
        
        for name, func in inspect.getmembers(section, inspect.ismethod):
            if name.startswith("_"):
                continue  # skip internals
            
            @function_tool(name_override=f"zap_{section_name}_{name}")
            def wrapped_func(*args, **kwargs):
                """Auto-generated tool wrapper for ZAP function."""
                return func(*args, **kwargs)
            
            tools.append(wrapped_func)
    
    return tools
