#It seems like the error is due to a misplaced explanation text within the code. Let's correct that and try again:


from openpyxl import load_workbook
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.neighbors import KNeighborsClassifier
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.preprocessing import OneHotEncoder
import matplotlib.pyplot as plt
import seaborn as sns

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

# Train the KNN model with 1 neighbor
model = KNeighborsClassifier(n_neighbors=1)
model.fit(X_train, y_train.argmax(axis=1))

# Make predictions
y_pred = model.predict(X_test)

# Evaluate the model
accuracy = accuracy_score(y_test.argmax(axis=1), y_pred)
print(f"Accuracy: {accuracy}")

# Create a confusion matrix
conf_matrix = confusion_matrix(y_test.argmax(axis=1), y_pred)
plt.figure(figsize=(8, 6))
sns.heatmap(conf_matrix, annot=True, cmap='Blues', xticklabels=encoder.categories_[0], yticklabels=encoder.categories_[0])
plt.xlabel('Predicted')
plt.ylabel('Actual')
plt.title('Confusion Matrix')
plt.show()


#This should resolve the syntax error, and the code should run without issues.
