import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from sklearn.utils.class_weight import compute_class_weight
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense, Dropout
from tensorflow.keras.callbacks import EarlyStopping
from tensorflow.keras.utils import to_categorical
from tensorflow.keras.optimizers import Adam

# === Load Data ===
train_df = pd.read_excel("C:/Users/Mikołaj/Desktop/Enginer/TRAINPRESENTATION.xlsx", header=0)
test_df = pd.read_excel("C:/Users/Mikołaj/Desktop/Enginer/TESTPRESENTATION.xlsx", header=0)

# === Set Column Names ===
column_names = [
    'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
    'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
    'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD',
    'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN',
    'AO', 'AP', 'AQ'
]
train_df.columns = column_names
test_df.columns = column_names

# === Encode Categorical Columns ===
train_df['A'] = train_df['A'].astype('category').cat.codes
train_df['B'] = train_df['B'].astype('category').cat.codes
test_df['A'] = test_df['A'].astype('category').cat.codes
test_df['B'] = test_df['B'].astype('category').cat.codes

# === Define Time Series Columns ===
start_index = column_names.index("E")
end_index = column_names.index("AQ")
price_columns = column_names[start_index:end_index + 1]
num_days = len(price_columns)

# === Extract Time Series Data ===
X_prices = train_df[price_columns].values
X_test_prices = test_df[price_columns].values

# === One-Hot Encode Targets ===
y = train_df["D"].astype(int)
y_cat = to_categorical(y)
num_classes = y_cat.shape[1]

# === Normalize Each Time Series Individually ===
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X_prices)
X_test_scaled = scaler.transform(X_test_prices)
# === Reshape for LSTM: (samples, timesteps, features) ===
X_scaled = X_scaled.reshape((X_scaled.shape[0], num_days, 1))
X_test_reshaped = X_test_scaled.reshape((X_test_scaled.shape[0], num_days, 1))

# === Train/Validation Split ===
X_train, X_val, y_train, y_val = train_test_split(
    X_scaled, y_cat, test_size=0.2, random_state=42, stratify=y
)

# === Compute Class Weights ===
y_int = np.argmax(y_train, axis=1)
class_weights = compute_class_weight(
    class_weight='balanced',
    classes=np.unique(y_int),
    y=y_int
)
class_weights = dict(enumerate(class_weights))

# === Build LSTM Model ===
model = Sequential()
model.add(LSTM(128, return_sequences=True, input_shape=(num_days, 1)))
model.add(Dropout(0.4))
model.add(LSTM(64))
model.add(Dropout(0.3))
model.add(Dense(64, activation='relu'))
model.add(Dense(num_classes, activation='softmax'))

# === Compile Model ===
optimizer = Adam(learning_rate=1e-4)
model.compile(loss='categorical_crossentropy', optimizer=optimizer, metrics=['accuracy'])

# === Train Model ===
early_stop = EarlyStopping(monitor='val_loss', patience=10, restore_best_weights=True)
history = model.fit(
    X_train, y_train,
    epochs=200,
    batch_size=32,
    validation_data=(X_val, y_val),
    class_weight=class_weights,
    callbacks=[early_stop],
    verbose=1
)

# === Evaluate Model ===
val_loss, val_accuracy = model.evaluate(X_val, y_val, verbose=0)
print(f"Validation Accuracy: {val_accuracy:.4f}")
print(f"Validation Loss: {val_loss:.4f}")

# === Predict on Test Set ===
probas = model.predict(X_test_reshaped)
confidences = np.max(probas, axis=1)
predicted_classes = np.argmax(probas, axis=1)

# === Apply Confidence Threshold ===
threshold = 0.2
final_preds = [
    pred if conf >= threshold else 0
    for pred, conf in zip(predicted_classes, confidences)
]

# === Save Predictions ===
test_df_with_predictions = test_df.copy()
test_df_with_predictions["D"] = final_preds
test_df_with_predictions.to_csv(
    "C:/Users/Mikołaj/Desktop/Enginer/PIERD.csv",
    index=False
)
