import inspect
import json

def expose_class_methods(instance, prefix=None):
    tools = []
    for name, method in inspect.getmembers(instance, predicate=inspect.ismethod):
        if name.startswith("_"):
            continue
        sig = inspect.signature(method)
        props, required = {}, []
        for param, param_data in sig.parameters.items():
            props[param] = {"type": "string", "description": param}
            if param_data.default == inspect.Parameter.empty:
                required.append(param)

        tools.append({
            "type": "function",
            "function": {
                "name": f"{prefix}_{name}" if prefix else name,
                "description": method.__doc__ or f"Call {name}",
                "parameters": {"type": "object", "properties": props, "required": required},
            },
        })
    return tools
