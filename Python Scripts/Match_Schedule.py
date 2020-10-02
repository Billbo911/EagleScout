import pandas as pd
import numpy as np

df = pd.read_excel('/EagleScout/Python Scripts/Match_Schedule/Match_Schdule.xlsx', sheet_name = None)
df['Match Schedule'].to_csv('/EagleScout/Python Scripts/Match_Schedule/Match_Schedule.csv')
