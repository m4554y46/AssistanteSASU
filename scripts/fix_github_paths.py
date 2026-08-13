import json, os

p = "moteur_rapport/config.json"
c = json.load(open(p))
c["output_dir"] = os.environ["GITHUB_WORKSPACE"]
json.dump(c, open(p, "w"), indent=2)
print(f'output_dir configure: {c["output_dir"]}')

pa = "agents/config.json"
a = json.load(open(pa))
a["output_dir"] = os.path.join(os.environ["GITHUB_WORKSPACE"], "agents", "output").replace("\\", "/")
json.dump(a, open(pa, "w"), indent=2)
print(f'agents output_dir configure: {a["output_dir"]}')
