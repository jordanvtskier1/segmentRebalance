import pandas as pd
import numpy as np
import pulp
from pulp import PULP_CBC_CMD
from pathlib import Path
import sys

# GLOBAL VARIABLES
# windows path to file
filepath = r"C:\Users\HB3245\NLG Capital Work\Segment Rebalance\159 Segment Rebalancing.xlsx"

# Portfolio data sheet name
dataSheetName = "nlgcap_holdings_plus"

# Funding target sheet name
fundingSheetName = "Funding Levels for Investments"

# Outpath
outpath = r"C:\Users\HB3245\NLG Capital Work\Segment Rebalance\balancedOutput.xlsx"

# Location of segments and "Net Liabilities or Desired Size" in funding level sheet:
rows = range(5, 27)
excelcolumns = ['A', 'F']

# Segments to balance
segments = [159, 165, 166, 155]

# Limit segment(s) to take assets out of:
# If left blank, program will default to totalMovesPercent as asset moves constraint,
# and assets can be moved between all segments
overFundSegments = [#159
                    ]

# percentage of assets in portfolio, as a float, that can be moved.
totalMovesPercent = 5

# Run time limit in seconds. Sometimes the optimizer finds the optimum solution in seconds,
# other times it will run near infinitely. Early stopping often still gives a great solution
# that is a fraction of a percent off optimum.
# If the optimizer stops early (final output in terminal: "Result - Stopped on time limit"),
# check the Objective value and Lower bound output values. Objective value is the optimum result found within time
# limits, Lower bound is the best possible result given infinite time. If they are very different, increasing the
# runtime could improve results.
runtime = 600

# Percentage the yield is allowed to change from current/target, as a float
yieldConstraintPercent = 1

# Percentage the duration is allowed to change from current/target, as a float
durationConstraintPercent = 1

# Percentage the asset class allocation is allowed to change from current/target, as a float
allocationConstraintPercent = 5

# If desired yield target is different than current, enter values in segment: target format
# It is not required to enter a target for every segment, the default is the current value.
# The problem might not be feasible if ambitious targets are entered
yieldTargetDict = {#155: 4.5,
                   #165: 5
                   }


# Same as above for duration
durationTargetDict = {#155: 7.5,
                      #165: 7
                      }

# Same as above for asset classes, but with format (segment, assetclass): target
assetclassTargetDict = {#(155, "Corporate"): 61,
                        #(165, "Muni"): 61
                        }

# There shouldn't be a reason to want to change the desired funding target, but if there is this is here just in case
fundingTargetDict = {#155: 700000000,
                     #165: 5400000000
                     }

def main() -> None:
    # load dataframe
    df, funding = loadData(excelcolumns)

    df, status = runOptimizer(df, funding)

    match status:
        case 0:
            sys.exit("Problem not solved, try increasing runtime")
        case -1:
            sys.exit("Problem is infeasible, loosen constraints or try less aggressive targets")
        case 1:
            # Print results for funding
            output_template = "{}   Old funding: {:>13,}   New funding: {:>13,}   Target: {:>13,}   Diff: {:>12,}"
            for j in funding['segment']:
                old = df[df['segment'] == j]["bv_gaap"].sum()
                new = df[df['newSegment'] == j]["bv_gaap"].sum()
                target = funding[funding['segment'] == j]['desiredlvl'].values[0]
                diff = target - new
                print(output_template.format(round(j), round(old), round(new), round(target), round(diff)))

            print("\n\n")
            # Print results for book yield
            for j in funding['segment']:
                values = df[df['segment'] == j]['by_gaap']
                weights = df[df['segment'] == j]['bv_gaap']
                rate = np.average(values, weights=weights)

                values = df[df['newSegment'] == j]['by_gaap']
                weights = df[df['newSegment'] == j]['bv_gaap']
                newrate = np.average(values, weights=weights)

                print(round(j), " Old BY:", round(rate, 2), "  New BY:", round(newrate, 2))

            print("\n\n")
            # Print results for duration
            for j in funding['segment']:
                values = df[df['segment'] == j]['effective_duration']
                weights = df[df['segment'] == j]['mv']
                rate = np.average(values, weights=weights)

                values = df[df['newSegment'] == j]['effective_duration']
                weights = df[df['newSegment'] == j]['mv']
                newrate = np.average(values, weights=weights)

                print(round(j), " Old OAD:", round(rate, 2), "  New OAD:", round(newrate, 2))

            print("\n\n")
            # Print results for asset class distribution
            output_template = "{} {:>25}   %Old: {:>5}    %New: {:>5}    Diff: {:>5}"
            for segment in funding['segment']:
                for assetclass in list(df['mandate_level_2'].value_counts().index):
                    oldsumclass = df.loc[(df['segment'] == segment) & (df['mandate_level_2'] == assetclass)][
                        "bv_gaap"].sum()
                    oldsumtotal = df.loc[(df['segment'] == segment)]["bv_gaap"].sum()
                    old = round(100 * oldsumclass / oldsumtotal, 2)

                    newsumclass = df.loc[(df['newSegment'] == segment) & (df['mandate_level_2'] == assetclass)][
                        "bv_gaap"].sum()
                    newsumtotal = df.loc[(df['newSegment'] == segment)]["bv_gaap"].sum()
                    new = round(100 * newsumclass / newsumtotal, 2)

                    diff = round((old - new), 2)
                    print(output_template.format(round(segment), assetclass, old, new, diff))

            print(f"\nTotal Assets moved: {df['equal'].value_counts()[0]}")

            outpath_universal = Path(outpath)
            df.to_excel(outpath_universal)

            print(f"\nOutput returned to {outpath}")

        case -2:
            sys.exit("Problem is unbounded")
        case -3:
            sys.exit("Problem is undefined")
        case _:
            sys.exit(f"Unknown error code: {status}")


def loadData(cols: list) -> tuple[pd.DataFrame, pd.DataFrame]:
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
    funding = funding[funding["segment"].isin(segments)].reset_index()

    print("Finished loading data")

    return df, funding


def runOptimizer(df: pd.DataFrame, funding: pd.DataFrame) -> tuple[pd.DataFrame, int]:
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
        realalloc = df.loc[i, 'segment']
        for j in funding['segment']:
            if j == realalloc:
                asset_set[(i, j)] = 1
            else:
                asset_set[(i, j)] = 0

    # CONSTRAINTS:

    # Only move assets out of specified segment(s). If overFundSegments is empty this will skip
    if overFundSegments:
        for i in df[~df['segment'].isin(overFundSegments)].index:
            for j in segments:
                prob += asset_vars[i, j] == asset_set[(i, j)]

    # Only move below totalMovesPercent of loan
    prob += pulp.lpSum([asset_vars[i, j] * asset_set[(i, j)]
                        for i in df.index
                        for j in funding['segment']]) >= int(len(df) * (1 - (totalMovesPercent / 100)))

    # Each loan can only be assigned to one segment
    for i in df.index:
        prob += pulp.lpSum(asset_vars[i, j] for j in funding['segment']) == 1

    for j in funding['segment']:
        # define absolute value of difference using
        # z[i] >= Ax[i] - By[i]
        # z[i] >= -(Ax[i] - By[i]) method
        prob += segments_diff[j] >= (
                pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index]) - getDesiredFunding(
            j, funding)) / getDesiredFunding(j, funding)

        prob += segments_diff[j] >= (getDesiredFunding(j, funding) - pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])) / getDesiredFunding(j, funding)

        # Constraints: WA Yield stays nearly constant
        prob += pulp.lpSum([df.loc[i, 'by_gaap'] * df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                            for i in df.index]) >= getWAyield(j, df) * (1-yieldConstraintPercent/100) * pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])
        prob += pulp.lpSum([df.loc[i, 'by_gaap'] * df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                            for i in df.index]) <= getWAyield(j, df) * (1+yieldConstraintPercent/100) * pulp.lpSum(
            [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)] for i in df.index])

        # Constraints: WA duration stays nearly constant
        prob += pulp.lpSum([df.loc[i, 'effective_duration'] * df.loc[i, 'mv'] * asset_vars[(i, j)]
                            for i in df.index]) >= getWAduration(j, df) * (1-durationConstraintPercent/100) * pulp.lpSum(
            [df.loc[i, 'mv'] * asset_vars[(i, j)] for i in df.index])
        prob += pulp.lpSum([df.loc[i, 'effective_duration'] * df.loc[i, 'mv'] * asset_vars[(i, j)]
                            for i in df.index]) <= getWAduration(j, df) * (1+durationConstraintPercent/100) * pulp.lpSum(
            [df.loc[i, 'mv'] * asset_vars[(i, j)] for i in df.index])

        # Constraints: asset class allocation stays nearly constant
        for assetclass in assetClassList:
            prob += pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                                for i in df[df['mandate_level_2'] == assetclass].index]) >= getAllocation(j,
                                                                                                          assetclass,
                                                                                                          df) * pulp.lpSum(
                [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                 for i in df.index]) * (1-allocationConstraintPercent/100)

            prob += pulp.lpSum([df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                                for i in df[df['mandate_level_2'] == assetclass].index]) <= getAllocation(j,
                                                                                                          assetclass,
                                                                                                          df) * pulp.lpSum(
                [df.loc[i, 'bv_gaap'] * asset_vars[(i, j)]
                 for i in df.index]) * (1+allocationConstraintPercent/100)

    # Objective: minimize sum of absolute value of percent difference from target funding
    prob += pulp.lpSum(segments_diff[j] for j in segments)

    print("Finished building optimizer, starting run")

    prob.solve(PULP_CBC_CMD(msg=True, timeLimit=runtime))

    # Output the results
    #for v in prob.variables():
        #print(v.name, "=", v.varValue)

    # Get the new asset distribution
    for i in df.index:
        for j in funding['segment']:
            if asset_vars[(i, j)].varValue == 1:
                df.loc[i, "newSegment"] = j
                if df.loc[i, "segment"] == j:
                    df.loc[i, "equal"] = 1
                else:
                    df.loc[i, "equal"] = 0
    return df, prob.status


# HELPER FUNCTIONS:
# Returns desired funding for segment
def getDesiredFunding(segment: int, funding: pd.DataFrame) -> float:
    try:
        if fundingTargetDict[segment]:
            return fundingTargetDict[segment]
    except KeyError:
        return funding.loc[funding['segment'] == segment, 'desiredlvl'].values[0]


# Returns starting WA duration of segment
def getWAduration(segment: int, df: pd.DataFrame) -> float:
    try:
        if durationTargetDict[segment]:
            return durationTargetDict[segment]
    except KeyError:
        values = df[df['segment'] == segment]['effective_duration']
        weights = df[df['segment'] == segment]['mv']
        return np.average(values, weights=weights)


# Returns starting WA yield of segment
def getWAyield(segment: int, df: pd.DataFrame) -> float:
    try:
        if yieldTargetDict[segment]:
            return yieldTargetDict[segment]
    except KeyError:
        values = df[df['segment'] == segment]['by_gaap']
        weights = df[df['segment'] == segment]['bv_gaap']
        return np.average(values, weights=weights)


# Returns percentage allocation
def getAllocation(segment: int, assetclass: str, df: pd.DataFrame) -> float:
    try:
        if assetclassTargetDict[(segment, assetclass)]:
            return assetclassTargetDict[(segment, assetclass)] / 100
    except KeyError:
        sumclass = df.loc[(df['segment'] == segment) & (df['mandate_level_2'] == assetclass)]["bv_gaap"].sum()
        sumtotal = df.loc[(df['segment'] == segment)]["bv_gaap"].sum()
        return sumclass / sumtotal


main()