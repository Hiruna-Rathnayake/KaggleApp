import sys
import json
import pickle
import pandas as pd
from scipy.sparse import hstack, csr_matrix
import os
import warnings
import numpy as np

warnings.filterwarnings("ignore")

# ------------------------------
# Paths and constants
# ------------------------------
scripts_folder = os.path.dirname(os.path.abspath(__file__))

NUM_FEATURES = [
    "body_length", "word_count", "has_url",
    "exclamations", "questions", "uppercase_ratio"
]

# ------------------------------
# Load TF-IDF vectorizers
# ------------------------------
with open(os.path.join(scripts_folder, "body_tfidf.pkl"), "rb") as f:
    body_tfidf = pickle.load(f)
with open(os.path.join(scripts_folder, "rule_tfidf.pkl"), "rb") as f:
    rule_tfidf = pickle.load(f)

# ------------------------------
# Load all per-rule models
# ------------------------------
models = {}
for file in os.listdir(scripts_folder):
    if file.startswith("logreg_model_rule_") and file.endswith(".pkl"):
        rule_name = file[len("logreg_model_rule_"):-4]
        with open(os.path.join(scripts_folder, file), "rb") as f:
            models[rule_name] = pickle.load(f)

# ------------------------------
# Utility functions
# ------------------------------
def clean_text(x):
    return str(x).lower() if x else ""

def add_numeric_features(df):
    df = df.copy()
    body = df["body"].fillna("")
    df["body_length"] = body.str.len()
    df["word_count"] = body.str.split().str.len()
    df["has_url"] = body.str.contains(r"http[s]?://", na=False).astype(int)
    df["exclamations"] = body.str.count("!")
    df["questions"] = body.str.count(r"\?")
    df["uppercase_ratio"] = body.apply(lambda x: sum(c.isupper() for c in x)/max(len(x),1))
    return df

# ------------------------------
# Persistent server loop
# ------------------------------
def main():
    print("Python server ready.", flush=True)

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue

        try:
            data = json.loads(line)
            comments = data.get("comments", [])
            if not comments:
                response = {"comments": [], "predictions": []}
                print(json.dumps(response), flush=True)
                continue

            df = pd.DataFrame({"body": comments})
            df = add_numeric_features(df)

            X_body = body_tfidf.transform(df["body"].apply(clean_text))
            X_num = csr_matrix(df[NUM_FEATURES].values, dtype="float32")

            # For each comment, track highest probability per rule
            results = []
            for i in range(len(df)):
                max_prob = 0
                violated_rule = None

                for rule_name, model in models.items():
                    # Use rule TF-IDF for this rule
                    X_rule = rule_tfidf.transform([rule_name]*1)
                    X_final = hstack([X_body[i], X_rule, X_num[i]]).tocsr()

                    prob = model.predict_proba(X_final)[:,1][0]
                    if prob > max_prob:
                        max_prob = prob
                        violated_rule = rule_name

                # Threshold for violation, e.g., 0.5
                status = "Safe"
                if max_prob >= 0.5:
                    status = "Violates rule"

                results.append({
                    "probability": float(max_prob),
                    "status": status,
                    "rule": violated_rule if status != "Safe" else ""
                })

            response = {
                "comments": comments,
                "results": results
            }
            print(json.dumps(response), flush=True)

        except Exception as e:
            err_resp = {"comments": comments if 'comments' in locals() else [], 
                        "results": [{"probability": 0.0, "status": "Error", "rule": "", "error": str(e)} for _ in comments]}
            print(json.dumps(err_resp), flush=True)

if __name__ == "__main__":
    main()
