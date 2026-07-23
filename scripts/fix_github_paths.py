import json, os

p = "moteur_rapport/config.json"
c = json.load(open(p))
c["output_dir"] = os.environ["GITHUB_WORKSPACE"]
json.dump(c, open(p, "w"), indent=2)
print(f'output_dir configure: {c["output_dir"]}')
