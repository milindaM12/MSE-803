# from ucimlrepo import fetch_ucirepo 
  
# # fetch dataset 
# iris = fetch_ucirepo(id=53) 
  
# # data (as pandas dataframes) 
# X = iris.data.features 
# y = iris.data.targets 
  
# # metadata 
# print(iris.metadata) 
  
# # variable information 
# print(iris.variables) 

import pandas as pd
from ucimlrepo import fetch_ucirepo

# Fetch dataset
iris = fetch_ucirepo(id=53)

X = iris.data.features
y = iris.data.targets

# Number of features
print("Number of features:", X.shape[1])

# Feature names
print("Feature names:")
print(X.columns.tolist())

# Number of classes
print("Number of classes:", y.nunique().iloc[0])

# Class names
print("Classes:")
print(y.iloc[:, 0].unique())


# Combine features and target
iris_df = pd.concat([X, y], axis=1)

# Count duplicate rows
print("Number of duplicate rows:", iris_df.duplicated().sum())

# Display duplicate rows
print("\nDuplicate records:")
print(iris_df[iris_df.duplicated()])
