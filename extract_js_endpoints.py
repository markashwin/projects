# Script Type: Standalone
# Engine: Jython

from org.parosproxy.paros.model import Model
from org.parosproxy.paros.model import SiteNode
from java.util import HashSet
import os

output_path = "C:/zap_output/site_tree_urls.txt"
seen = HashSet()

def traverse(node):
    msg = node.getHistoryReference()
    if msg is not None:
        try:
            uri = msg.getHttpMessage().getRequestHeader().getURI().toString()
            
            if uri.endswith(".js") and not seen.contains(uri):
                seen.add(uri)
                f.write(uri + "\n")
                print(uri)
        except:
            pass
    # Recurse into children
    for i in range(node.getChildCount()):
        traverse(node.getChildAt(i))

root = Model.getSingleton().getSession().getSiteTree().getRoot()

with open(output_path, "w") as f:
    traverse(root)

print("✅ Done. Extracted", seen.size(), "unique URLs.")
