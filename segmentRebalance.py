import pandas as pd
import numpy as np
import pulp
from pulp import PULP_CBC_CMD
import openpyxl
from pathlib import Path

# GLOBAL VARIABLES
# windows path to file
filepath = r"C:\Users\HB3245\NLG Capital Work\Segment Rebalance\159 Segment Rebalancing.xlsx"

# Portfolio data sheet name
dataSheetName = "nlgcap_holdings_plus"

# Funding target sheet name
fundingSheetName = "Funding Levels for Investments"

# Location of segments and "Net Liabilities or Desired Size" in funding level sheet:
rows = range(5, 27)
columns = ['A', 'F']

# Segments to balance
segments = [159, 165, 166, 155]

# Limit segment(s) to take assets out of:
# If left blank, program will default to totalMovesPercent as asset moves constraint,
# and assets can be moved between all segments
overFundSegments = [159]

# percentage of assets in portfolio, as a float, that can be moved.
totalMovesPercent = 5

# Run time limit in seconds. Sometimes the optimizer finds the optimum solution in seconds,
# other times it will run near infinitely. Early stopping often still gives a great solution
# that is a fraction of a percent off optimum, which I do not think is worth waiting more than
# 10 minutes for. If the optimizer stops early (final output in terminal: "Result - Stopped on time limit"),
# check the Objective value and Lower bound output values. Objective value is the optimum result found within time
# limits, Lower bound is the best possible result given infinite time. If they are very different, increasing the
# runtime could improve results.
runtime = 600


# Load assets as df
path_universal = Path(filepath)
df = pd.read_excel(path_universal, sheet_name=dataSheetName)
# Fill missing values with average of column. This method is not perfect and can be updated later
df['effective_duration'] = df['effective_duration'].fillna(value=df['effective_duration'].mean())

# Load target funding as funding
funding = pd.read_excel(path_universal, sheet_name=fundingSheetName)

# Helper function: convert excel column letter to int
excelColNum = lambda a: 0 if a.upper() == '' else ord(a[-1].upper()) - ord('A') + 26 * excelColNum(a[:-1].upper())
columns = [excelColNum(col) for col in columns]

funding = funding.iloc[rows, columns].rename(
    columns={funding.columns[columns[0]]: "segment", funding.columns[columns[1]]: "desiredlvl"})

# Filter portfolio for only segments selected for balancing
df = df[df["segment"].isin(segments)].reset_index()
df["uniqueid"] = df.index
funding = funding[funding["segment"].isin(segments)].reset_index()

#
assetClassList = list(df['mandate_level_2'].value_counts().index)