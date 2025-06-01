import pandas as pd
import yaml
from create_report import create_report, create_statistics
from natsort import natsorted
import sys

def load_config(confiig_path):
    with open(confiig_path, 'r') as file:
        return yaml.safe_load(file)

def calculate_statistics(merged_data, column_name, threshold):
    stats_gt = []
    stats_mod = []

    column_gt = merged_data.iloc[1:, 1]
    column_mod = merged_data.iloc[1:, 2]

    #print(column_gt)
    #print(column_mod)

    stats_gt.append({
        "Measurement": column_name+"(GT)",
        "Mean": column_gt.mean(),
        "Variance": column_gt.var(),
        "Max": column_gt.max(),
        "Min": column_gt.min(),
        "Median": column_gt.median(),
        "5%": column_gt.quantile(0.05),
        "95%": column_gt.quantile(0.95),
    })

    stats_mod.append({
        "Measurement": column_name+"(MOD)",
        "Mean": column_mod.mean(),
        "Variance": column_mod.var(),
        "Max": column_mod.max(),
        "Min": column_mod.min(),
        "Median": column_mod.median(),
        "5%": column_mod.quantile(0.05),
        "95%": column_mod.quantile(0.95),
    })

    return stats_gt, stats_mod

def compare_all_values(merged_data, threshold, df_gt_cases, df_mod_cases):
    results = []

    for i in range(merged_data.shape[0]):
        case_name = merged_data.iloc[i, 0]
        m1 = merged_data.iloc[i, 1]
        m2 = merged_data.iloc[i, 2]
        if case_name not in df_gt_cases and case_name in df_mod_cases:
            issue = "new sample in MOD"
            diff = None
        elif case_name not in df_mod_cases and case_name in df_gt_cases:
            issue = "missing sample in MOD"
            diff = None
        elif pd.isna(m1) and pd.isna(m2):
            issue = "both values are NaN"
            diff = None
        elif pd.isna(m1):
            issue = "GT is NaN"
            diff = None
        elif pd.isna(m2):
            issue = "MOD is NaN"
            diff = None
        else:
            if m1 == 0:
                diff = float('inf') if m2 != 0 else 0
            else:
                diff = abs(m1 - m2) / abs(m1) * 100

            if diff <= threshold:
                issue = "within threshold"
            else:
                issue = "above threshold"

        results.append({
            "CaseName": case_name,
            "GT Value": m1,
            "MOD Value": m2,
            "% Difference": None if m1 != m1 or m2 != m2 else round(diff, 2),
            "% Threshold": threshold,
            "Issue": issue
        })
    return results

def main():
    if len(sys.argv) != 2:
        sys.exit(1)

    config_path = sys.argv[1]
    config = load_config(config_path)

    csv_settings = config.get("csv_settings", {})
    delimiter = csv_settings.get("delimiter", ",")
    decimal = csv_settings.get("decimal", ".")
    na_values = csv_settings.get("missing_value", "NaN")

    gt_file = config["files"]["ground_truth"]
    mod_file = config["files"]["modified_data"]

    df1 = pd.read_csv(gt_file, delimiter=delimiter, decimal=decimal, na_values=na_values)
    df2 = pd.read_csv(mod_file, delimiter=delimiter, decimal=decimal, na_values=na_values)

    thresholds = config.get("thresholds", {})
    global_threshold = thresholds.get("global", 10.0)
    measurements_thresholds = thresholds.get("measurements", {})
    statistic_thresholds = thresholds.get("statistics", {})

    file_name = config["files"]["output_report"]

    offset = 0
    offset2 = 0

    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        for column_name in df1.columns:
            if column_name == "CaseName":
                continue


            merged = pd.merge(
                df1[['CaseName', column_name]],
                df2[['CaseName', column_name]],
                on='CaseName',
                how='outer',
                suffixes=('_gt', '_mod')
            )

            df1_cases = set(df1['CaseName'].dropna())
            df2_cases = set(df2['CaseName'].dropna())

            merged_sorted = merged.set_index('CaseName')
            merged_sorted = merged_sorted.loc[natsorted(merged_sorted.index)].reset_index()

            threshold = measurements_thresholds.get(column_name, global_threshold)
            compared_values = compare_all_values(merged_sorted, threshold, df1_cases, df2_cases)
            offset = create_report(compared_values, writer, offset, column_name, threshold)

            threshold = statistic_thresholds.get(column_name, global_threshold)
            stats_gt, stats_mod = calculate_statistics(merged_sorted, column_name, threshold)

            offset2 = create_statistics(writer, stats_gt, stats_mod, offset2, column_name, threshold)


if __name__ == "__main__":
    main()