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


def loadData(cols:list):
    # Load assets as df
    path_universal = Path(filepath)
    df = pd.read_excel(path_universal, sheet_name=dataSheetName)
    # Fill missing values with average of column. This method is not perfect and can be updated later
    df['effective_duration'] = df['effective_duration'].fillna(value=df['effective_duration'].mean())

    # Load target funding as funding
    funding = pd.read_excel(path_universal, sheet_name=fundingSheetName)

    # Helper function: convert excel column letter to int
    excelColNum = lambda a: 0 if a.upper() == '' else ord(a[-1].upper()) - ord('A') + 26 * excelColNum(a[:-1].upper())
    columns = [excelColNum(col) for col in cols]

    funding = funding.iloc[rows, columns].rename(
        columns={funding.columns[columns[0]]: "segment", funding.columns[columns[1]]: "desiredlvl"})

    # Filter portfolio for only segments selected for balancing
    df = df[df["segment"].isin(segments)].reset_index()
    df["uniqueid"] = df.index
    funding = funding[funding["segment"].isin(segments)].reset_index()

    return df, funding


def runOptimizer(df, funding):
    # Get list of asset types from data
    assetClassList = list(df['mandate_level_2'].value_counts().index)

    # Create LpProblem object
    prob = pulp.LpProblem("Balance_Segments", pulp.LpMinimize)
    # Define decision variables
    asset_vars = pulp.LpVariable.dicts("asset",
                                   ((i, j) for i in df.index for j in funding['segment']),
                                   cat='Binary')
    # Variable to minimize. Equal to segment value difference from desired value
    segments_diff = pulp.LpVariable.dicts('segments_diff', segments, cat='Continuous')

    # Create dictionary with current asset allocation
    asset_set = {}
    for i in df.index:
        realalloc = df.loc[i,'segment']
        for j in funding['segment']:
            if j == realalloc:
                asset_set[(i,j)] = 1
            else:
                asset_set[(i,j)] = 0

    # CONSTRAINTS:

    # Only move assets out of specified segment(s). If overFundSegments is empty this will skip
    if overFundSegments:
        for i in df[~df['segment'].isin(overFundSegments)].index:
            for j in segments:
                prob += asset_vars[i, j] == asset_set[(i,j)]

    # Only move below totalMovesPercent of loan
    prob += pulp.lpSum([asset_vars[i, j] * asset_set[(i,j)]
                        for i in df.index
                        for j in funding['segment']]) >= int(len(df)*(1-(totalMovesPercent/100)))

    # Each loan can only be assigned to one segment
    for i in df.index:
        prob += pulp.lpSum(asset_vars[i, j] for j in funding['segment']) == 1

    for j in funding['segment']:
        # define absolute value of difference using
        # z[i] >= Ax[i] - By[i]
        # z[i] >= -(Ax[i] - By[i]) method
        prob += segments_diff[j] >= (
                    pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index]) - getDesiredFunding(
                j)) / getDesiredFunding(j)

        prob += segments_diff[j] >= (getDesiredFunding(j) - pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])) / getDesiredFunding(j)

        # Constraints: WA Yield stays nearly constant
        prob += pulp.lpSum([df.loc[i, 'by_gaap'] * df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                            for i in df.index]) >= getWAyield(j) * 0.995 * pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])
        prob += pulp.lpSum([df.loc[i, 'by_gaap'] * df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                            for i in df.index]) <= getWAyield(j) * 1.005 * pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])

        # Constraints: WA duration stays nearly constant
        prob += pulp.lpSum([df.loc[i, 'effective_duration'] * df.loc[i, 'mv'] * asset_vars[(i, j)]
                            for i in df.index]) >= getWAduration(j) * 0.995 * pulp.lpSum(
            [df.loc[i, 'mv'] * asset_vars[(i, j)] for i in df.index])
        prob += pulp.lpSum([df.loc[i, 'effective_duration'] * df.loc[i, 'mv'] * asset_vars[(i, j)]
                            for i in df.index]) <= getWAduration(j) * 1.005 * pulp.lpSum(
            [df.loc[i, 'mv'] * asset_vars[(i, j)] for i in df.index])

        # Constraints: asset class allocation stays nearly constant
        for assetclass in assetclasslist:
            prob += pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                                for i in df[df['mandate_level_2'] == assetclass].index]) >= getAllocation(j,
                                                                                                          assetclass) * pulp.lpSum(
                [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                 for i in df.index]) * 0.98

            prob += pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                                for i in df[df['mandate_level_2'] == assetclass].index]) <= getAllocation(j,
                                                                                                          assetclass) * pulp.lpSum(
                [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                 for i in df.index]) * 1.02

    # Objective: minimize sum of absolute value of percent difference from target funding
    prob += pulp.lpSum(segments_diff[j] for j in segments)

    prob.solve(PULP_CBC_CMD(timeLimit=runtime))

    # Output the results
    for v in prob.variables():
        print(v.name, "=", v.varValue)

    # Get the new asset distribution
    for i in df.index:
        for j in funding['segment']:
            if asset_vars[(i, j)].varValue == 1:
                df.loc[i, "newSegment"] = j
                if df.loc[i, "segment"] == j:
                    df.loc[i, "equal"] = 1
                else:
                    df.loc[i, "equal"] = 0