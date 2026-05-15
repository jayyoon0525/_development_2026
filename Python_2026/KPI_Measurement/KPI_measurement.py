import pandas as pd

data = [
    # E10 count, F10 score
    (11, 3),
    # G10 count, H10 score
    (21, 4),
    # I10 count, J10 score
    (23, 4.5),
    # K10 count, L10 score
    (16, 4.5)
]

df = pd.DataFrame(data, columns=["Count", "Score"])
df["Weighted"] = df["Count"] * df["Score"]
df["Percent"] = df["Weighted"] / df["Weighted"].sum()

print(df)