import re
import requests
import os
import sys

URL = "https://raw.githubusercontent.com/wikimedia/mediawiki/master/includes/Languages/Data/ZhConversion.php"
OPENCC_DIR = r".venv\Lib\site-packages\opencc\dictionary"

def extract_mediawiki_dict(array_name, text):
    pattern = rf"const {array_name}\s*=\s*\[(.*?)\];"
    match = re.search(pattern, text, re.DOTALL)
    if not match:
        print(f"Array {array_name} not found!")
        return {}
    array_content = match.group(1)
    kv_pattern = r"'([^']+)'\s*=>\s*'([^']+)'"
    kv_pairs = re.findall(kv_pattern, array_content)
    return {k: v for k, v in kv_pairs}

def load_opencc_dict(filename):
    d = {}
    path = os.path.join(OPENCC_DIR, filename)
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line: continue
            parts = line.split("\t")
            if len(parts) == 2:
                d[parts[0]] = parts[1].split(" ")
    return d

print("Downloading ZhConversion.php...")
try:
    response = requests.get(URL)
    response.raise_for_status()
    content = response.text
except Exception as e:
    print(f"Failed to download ZhConversion.php: {e}")
    sys.exit(1)

zh2tw = extract_mediawiki_dict('ZH_TO_TW', content)
zh2hant = extract_mediawiki_dict('ZH_TO_HANT', content)
zh2tw.update(zh2hant)

print(f"Extracted {len(zh2tw)} mappings from Wikipedia.")

# Separate into chars and phrases
wiki_chars = {}
wiki_phrases = {}
for k, v in zh2tw.items():
    if len(k) == 1:
        wiki_chars[k] = v
    else:
        wiki_phrases[k] = v

print("Loading OpenCC dictionaries...")
st_chars = load_opencc_dict("STCharacters.txt")
st_phrases = load_opencc_dict("STPhrases.txt")
tw_phrases = load_opencc_dict("TWPhrases.txt")

# OpenCC's s2tw is STPhrases -> STCharacters -> TWPhrases
# So we merge wiki phrases into STPhrases and TWPhrases? 
# To simplify, we'll build custom phrases (STPhrases + TWPhrases + wiki_phrases)
# and custom chars (STCharacters + wiki_chars).

merged_phrases = {}
# Base: STPhrases
for k, v in st_phrases.items():
    merged_phrases[k] = v
# Apply TWPhrases on top
for k, v in tw_phrases.items():
    if k in merged_phrases:
        merged_phrases[k] = v + [x for x in merged_phrases[k] if x not in v]
    else:
        merged_phrases[k] = v
# Apply wiki phrases
for k, v in wiki_phrases.items():
    if k in merged_phrases:
        if v not in merged_phrases[k]:
            merged_phrases[k].append(v)
    else:
        merged_phrases[k] = [v]

merged_chars = {}
for k, v in st_chars.items():
    merged_chars[k] = v
for k, v in wiki_chars.items():
    if k in merged_chars:
        if v not in merged_chars[k]:
            merged_chars[k].append(v)
    else:
        merged_chars[k] = [v]

print("Writing custom dictionaries...")
def write_dict(d, out_path):
    with open(out_path, "w", encoding="utf-8") as f:
        # Sort by key length descending, then alphabetical
        for k, v_list in sorted(d.items(), key=lambda x: (-len(x[0]), x[0])):
            v_str = " ".join(v_list)
            f.write(f"{k}\t{v_str}\n")

write_dict(merged_phrases, "my_STPhrases.txt")
write_dict(merged_chars, "my_STCharacters.txt")

print("Generating my_s2tw.json...")
json_content = """{
  "name": "Custom Simplified to Traditional (Taiwan)",
  "segmentation": {
    "type": "mmseg",
    "dict": {
      "type": "text",
      "file": "my_STPhrases.txt"
    }
  },
  "conversion_chain": [
    {
      "dict": {
        "type": "text",
        "file": "my_STPhrases.txt"
      }
    },
    {
      "dict": {
        "type": "text",
        "file": "my_STCharacters.txt"
      }
    }
  ]
}"""
with open("my_s2tw.json", "w", encoding="utf-8") as f:
    f.write(json_content)

print("Done!")
