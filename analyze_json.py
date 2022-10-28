import os
import json

res = {}

def main(path):
    for root, _, files in os.walk(path):
        for file in files[:1]:
            if ".csv" in file:
                continue

            filepath = f"{root}\{file}"

            with open(filepath, "r", encoding="UTF-8") as f:
                obj = json.load(f)

                analyze(res, obj)
                for k, v in res.items():
                    if type(v).__name__ == "list":
                        res[k] = v[0]

                with open("result.json", "w", encoding="UTF-8") as f:
                    print(json.dumps(res), file=f)
                

def analyze(d, obj):
    if type(obj).__name__ == "dict":
        for k, v in obj.items():
            if type(v).__name__ == "dict":
                d[k] = {}
                analyze(d[k], v)
            elif type(v).__name__ == "list":
                d[k] = []
                analyze(d[k], v)
            else:
                d[k] = type(v).__name__

    elif type(obj).__name__ == "list":
        for k in obj:
            if type(k).__name__ == "dict":
                d.append({})
                analyze(d[-1], k)
            elif type(k).__name__ == "list":
                d.append([])
                analyze(d[-1], k)
            else:
                d.append(type(k).__name__)


if __name__ == "__main__":
    main(r"data\Parser20152017\Preprocessed files\2017\2017_01")
