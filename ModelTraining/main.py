import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
import xgboost as xgb
import numpy as np
import openpyxl
df = pd.read_csv("TrainData4.csv", header=None)
df.columns = ['A', 'B', 'C', 'D'] + [f'F{i}' for i in range(df.shape[1] - 4)]
X = df.drop(columns=["D"])
y = df["D"]
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
rf = RandomForestRegressor(n_estimators=100, random_state=42)
rf.fit(X_train, y_train)
y_pred_rf = rf.predict(X_test)
xgbr = xgb.XGBRegressor(n_estimators=100, learning_rate=0.1)
xgbr.fit(X_train, y_train)
y_pred_xgb = xgbr.predict(X_test)
dict_people = {}
dict_stocks = {}
wb_stocks = openpyxl.load_workbook("C:/Users/Mikołaj/Desktop/Enginer/IdPeopleStocks.xlsx")
ws_stocks = wb_stocks.active
for row in range(1, 15230):
    dict_people[str(ws_stocks["A" + str(row)].value)] = ws_stocks["C" + str(row)].value
    dict_stocks[str(ws_stocks["B" + str(row)].value)] = ws_stocks["D" + str(row)].value

def evaluate(name, y_true, y_pred):
    print(f"\n{name} Results:")
    print("R² score:", r2_score(y_true, y_pred))
    print("RMSE:", np.sqrt(mean_squared_error(y_true, y_pred)))
evaluate("Random Forest", y_test, y_pred_rf)
evaluate("XGBoost", y_test, y_pred_xgb)
# Load future data from Excel
future_df = pd.read_excel("C:/Users/Mikołaj/Desktop/Enginer/THIRDEXCEL.xlsx", sheet_name='Sheet1', header=0)

# Reassign correct column names (same as training)
future_df.columns = ['A', 'B', 'C', 'D'] + [f'F{i}' for i in range(df.shape[1] - 4)]

# Keep original copy for output
output_df = future_df.copy()

# Predict using the trained model
X_future = future_df.drop(columns=["D"])
predictions = rf.predict(X_future)

# Round predictions for clarity (optional)
output_df["Predicted_D"] = np.round(predictions, 4)

# Save back to Excel (add a new column for prediction)
output_path = "C:/Users/Mikołaj/Desktop/Enginer/THIRDEXCEL_PREDICTED.xlsx"
output_df.to_excel(output_path, index=False)

print(f"\n✅ Predictions added to '{output_path}' successfully.")