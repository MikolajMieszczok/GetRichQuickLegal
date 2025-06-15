import pandas as pd
import numpy as np
from sklearn.preprocessing import StandardScaler
from sklearn.model_selection import train_test_split
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense, Dropout
from tensorflow.keras.utils import to_categorical

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

# === Separate Features and Target ===
X = train_df.drop(columns=["D"])
y = train_df["D"].astype(int)

# === One-hot encode the target ===
y_cat = to_categorical(y)
num_classes = y_cat.shape[1]

# === Scale the features ===
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X)
X_test = test_df.drop(columns=["D"], errors='ignore')
X_test_scaled = scaler.transform(X_test)

# === Reshape for LSTM [samples, timesteps, features] ===
X_reshaped = X_scaled.reshape((X_scaled.shape[0], 1, X_scaled.shape[1]))
X_test_reshaped = X_test_scaled.reshape((X_test_scaled.shape[0], 1, X_test_scaled.shape[1]))

# === Build LSTM Model ===
model = Sequential()
model.add(LSTM(64, input_shape=(1, X_reshaped.shape[2]), return_sequences=False))
model.add(Dropout(0.5))
model.add(Dense(64, activation='relu'))
model.add(Dense(num_classes, activation='softmax'))

model.compile(loss='categorical_crossentropy', optimizer='adam', metrics=['accuracy'])

# === Train Model ===
model.fit(X_reshaped, y_cat, epochs=200, batch_size=32, validation_split=0.2)

# === Predict Probabilities ===
probas = model.predict(X_test_reshaped)
confidences = np.max(probas, axis=1)
predicted_classes = np.argmax(probas, axis=1)

# === Confidence Thresholding ===
threshold = 0.2
final_preds = [
    pred if conf >= threshold else 0
    for pred, conf in zip(predicted_classes, confidences)
]

# === Attach Predictions and Save ===
test_df_with_predictions = test_df.copy()
test_df_with_predictions["D"] = final_preds

test_df_with_predictions.to_csv(
    "C:/Users/Mikołaj/Desktop/Enginer/PIERD.csv",
    index=False
)
