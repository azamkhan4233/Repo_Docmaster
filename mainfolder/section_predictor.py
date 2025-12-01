# section_predictor.py
import os
import re
import pickle

# 🔴 Absolute path to your trained model folder (Option B)
# Change this ONLY if you move the NLP_Training folder somewhere else.
MODEL_DIR = r"C:\Users\MohdAzam\Desktop\NLP_Training\trained_tfidf_model"

if not os.path.isdir(MODEL_DIR):
    raise RuntimeError(f"MODEL_DIR does not exist: {MODEL_DIR}")

# Load vectorizer & classifier
with open(os.path.join(MODEL_DIR, "vectorizer.pkl"), "rb") as f:
    VECTORIZER = pickle.load(f)

with open(os.path.join(MODEL_DIR, "classifier.pkl"), "rb") as f:
    CLASSIFIER = pickle.load(f)


# --------------------------
# Rule-based helper
# --------------------------
def _rule_based_label(text: str):
    low = (text or "").lower().strip()

    # References: numeric bracket + DOI/URLs
    if re.match(r"^\[\d+\]", text.strip()):
        return "REFERENCES"
    if "doi.org" in low or "http://" in low or "https://" in low:
        return "REFERENCES"

    # Custom sections: acknowledgements, declarations, certificates, etc.
    if any(word in low for word in [
        "acknowledg", "acknowledgement", "acknowledgment",
        "declaration", "certificate",
        "plagiarism", "bona fide", "supervisor certificate"
    ]):
        return "CUSTOM"

    return None


# --------------------------
# Public API
# --------------------------
def predict_section(text: str) -> str:
    """
    Hybrid Section Classifier:
    - Uses rules for REFERENCES and CUSTOM-like content
    - Falls back to SVM classifier for others
    Returns label like 'ABSTRACT', 'INTRODUCTION', 'METHODOLOGY', etc.
    """
    text = (text or "").strip()
    if len(text) < 20:
        # Very short lines are usually noise / labels / misc
        return "CUSTOM"

    # 1) Rule-based first
    rule = _rule_based_label(text)
    if rule:
        return rule

    # 2) ML-based fallback
    vec = VECTORIZER.transform([text])
    label = CLASSIFIER.predict(vec)[0]
    return label
