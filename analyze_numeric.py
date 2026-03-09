import pandas as pd
f2 = "Car_Crash_Preprocessed.xlsx"
df2 = pd.read_excel(f2, engine='openpyxl')
print("Unique values per column in preprocessed:\n")
for c in ['County', 'Weekday', 'Severity', 'ClearWeather', 'Month', 'Highway', 'Daylight']:
    print(c, ":", len(df2[c].unique()), "unique values, Min:", df2[c].min(), "Max:", df2[c].max())
