# DataMining

## Datasets

### 1. `Big Data Files - Data Analytics.xlsx`
* **Description**: The raw, original un-preprocessed car crash dataset.
* **Format**: Excel file `.xlsx`
* **Usage**: Raw analysis, EDA, and understanding the initial features before any drops or handling was performed. 

### 2. `Car_Crash_Preprocessed.xlsx`
* **Description**: Cleaned dataset. 
* **Transformations Applied**:
  * Dropped irrelevant or high-cardinality string identifiers (`ID`, `City`).
  * Handled missing and null values.
  * One-hot encoded nominal text string classes (`CrashType`, `ViolCat`).
* **Warning**: Contains ordinal numbers (e.g. `County` IDs ranging from 0-56) that distance/linear models might mistakenly assume have mathematical weight. Also contains boolean True/False variables.
* **Safe For**: Tree-based models (Decision Trees, Random Forest, XGBoost).

### 3. `Car_Crash_Fully_Preprocessed.csv`
* **Description**: Fully processed and standardized dataset for all ML algorithms. 
* **Transformations Applied (from step 2)**:
  * Categorical/cyclical integers (`County`, `Month`, `Weekday`) have been one-hot encoded using `pd.get_dummies()`. This expands the shape of the dataset significantly but prevents gradient/linear models from assuming ordinal relationships.
  * All boolean `True`/`False` variables have been explicitly cast to `1` or `0` integers.
* **Format**: Comma-Separated Values `.csv` for faster future data loading.
* **Safe For**: Distance-based models (K-Nearest Neighbors), Support Vector Machines (SVM), Neural Networks, and Linear/Logistic Regression. This can also be used for Tree-based models.