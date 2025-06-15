import pandas as pd
import numpy as np
from lightgbm import LGBMClassifier
from sklearn.model_selection import train_test_split

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

train_df['A'] = train_df['A'].astype('category')
train_df['B'] = train_df['B'].astype('category')
test_df['A'] = test_df['A'].astype('category')
test_df['B'] = test_df['B'].astype('category')
# === Prepare Training Data ===
X_train = train_df.drop(columns=["D"])
y_train = train_df["D"].astype(int)
# === Train LightGBM Multiclass Classifier ===
model = LGBMClassifier(
    objective="multiclass",
    num_class=y_train.nunique(),
    n_estimators=300,
    class_weight='balanced',
    random_state=42
)
model.fit(X_train, y_train)

# === Prepare Test Data ===
X_test = test_df.drop(columns=["D"], errors='ignore')

# === Predict Probabilities ===
probas = model.predict_proba(X_test)
confidences = np.max(probas, axis=1)
predicted_classes = np.argmax(probas, axis=1)

# === Confidence Thresholding ===
threshold = 0.2  # You can experiment with 0.3–0.6
final_preds = [
    pred if conf >= threshold else 0
    for pred, conf in zip(predicted_classes, confidences)
]

# === Attach Predictions ===
test_df_with_predictions = test_df.copy()
test_df_with_predictions["D"] = final_preds

# === Save to CSV ===
test_df_with_predictions.to_csv(
    "C:/Users/Mikołaj/Desktop/Enginer/PIERD.csv",
    index=False
)