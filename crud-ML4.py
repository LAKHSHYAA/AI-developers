from openpyxl import load_workbook
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_squared_error
from sklearn.preprocessing import OneHotEncoder
import matplotlib.pyplot as plt

# Load the Excel file into a DataFrame
wb = load_workbook("example.xlsx")
ws = wb.active
data = ws.values
columns = next(data)[0:]
df = pd.DataFrame(data, columns=columns)

# Prepare the data for training
X = df[['Age', 'Score']]
y = df['City']

# One-hot encode the categorical variable 'City'
encoder = OneHotEncoder()
y_encoded = encoder.fit_transform(y.values.reshape(-1, 1)).toarray()

# Split the data into training and test sets
X_train, X_test, y_train, y_test = train_test_split(X, y_encoded, test_size=0.2, random_state=42)

# Train the model
model = LinearRegression()
model.fit(X_train, y_train)

# Make predictions
y_pred = model.predict(X_test)

# Inverse transform the encoded predictions to get city names
y_pred_city = encoder.inverse_transform(y_pred)

# Evaluate the model
mse = mean_squared_error(y_test, y_pred)
print(f"Mean Squared Error: {mse}")

# Plot the predictions against the actual values
plt.scatter(y_test, y_pred)
plt.xlabel('Actual City')
plt.ylabel('Predicted City')
plt.title('Predicted vs Actual City')
plt.show()
