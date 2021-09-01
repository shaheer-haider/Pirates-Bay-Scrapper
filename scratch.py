import pandas as pd

moviesDF = pd.read_excel("scrapped.xlsx")

moviesDFNoDuplicate = moviesDF.drop_duplicates(subset="movieName")

print(len(moviesDF))
print(len(moviesDFNoDuplicate))

moviesDFNoDuplicate.to_excel("scrappingDone.xlsx",index=False)