import os, base64, json

assets_dir = r"C:\Users\Win\.gemini\antigravity-ide\scratch\pronouns-presentation\assets"
output_file = r"C:\Users\Win\.gemini\antigravity-ide\scratch\pronouns-presentation\assets_base64.js"

data = {}
for fname in os.listdir(assets_dir):
    if fname.endswith(".jpg"):
        key = os.path.splitext(fname)[0]
        path = os.path.join(assets_dir, fname)
        with open(path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
            data[key] = f"data:image/jpeg;base64,{b64}"

with open(output_file, "w", encoding="utf-8") as f:
    f.write("window.SLIDE_ASSETS = " + json.dumps(data) + ";\n")

print(f"Generated assets_base64.js with {len(data)} images. File size: {os.path.getsize(output_file)} bytes")
