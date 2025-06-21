import pandas as pd
import numpy as np
from lightgbm import LGBMClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, log_loss

# === Load Data ===
train_df = pd.read_excel("C:/Users/Mikołaj/Desktop/Enginer/TRAINPRESENTATION.xlsx", header=0)
test_df = pd.read_excel("C:/Users/Mikołaj/Desktop/Enginer/TESTPRESENTATION.xlsx", header=0)

# === Set Proper Column Names ===
column_names = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
    'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD',
    'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
    'AO', 'AP', 'AQ'
]
train_df.columns = column_names
test_df.columns = column_names

# === Set Categorical Columns ===
for col in ['A', 'B']:
    train_df[col] = train_df[col].astype('category')
    test_df[col] = test_df[col].astype('category')

# === Prepare Training Data ===
X = train_df.drop(columns=["D"])
y = train_df["D"].astype(int)

# === Train/Validation Split ===
X_train, X_val, y_train, y_val = train_test_split(X, y, test_size=0.2, random_state=42, stratify=y)

# === Train LightGBM Multiclass Classifier ===
model = LGBMClassifier(
    objective="multiclass",
    num_class=y.nunique(),
    n_estimators=300,
    class_weight='balanced',
    random_state=42
)
model.fit(X_train, y_train)

# === Evaluation ===
val_preds_proba = model.predict_proba(X_val)
val_preds = np.argmax(val_preds_proba, axis=1)

accuracy = accuracy_score(y_val, val_preds)
loss = log_loss(y_val, val_preds_proba)

print(f"Validation Accuracy: {accuracy:.4f}")
print(f"Validation Log Loss: {loss:.4f}")

# === Predict on Test Data ===
X_test = test_df.drop(columns=["D"], errors='ignore')
probas = model.predict_proba(X_test)
confidences = np.max(probas, axis=1)
predicted_classes = np.argmax(probas, axis=1)

# === Apply Confidence Threshold ===
threshold = 0.3
final_preds = [
    pred if conf >= threshold else 0
    for pred, conf in zip(predicted_classes, confidences)
]

test_df_with_predictions = test_df.copy()
test_df_with_predictions["D"] = final_preds

# === Save Output ===
test_df_with_predictions.to_csv(
    "C:/Users/Mikołaj/Desktop/Enginer/PIERD.csv",
    index=False
)
