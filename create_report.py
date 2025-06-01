from matplotlib import pyplot as plt
from io import BytesIO
from xlsxwriter.utility import xl_col_to_name
import pandas as pd



def create_chart_image(df_differences, measurement_name):
    fig, ax = plt.subplots(figsize=(7.67, 5))

    x = range(len(df_differences))

    ax.plot(x, df_differences["GT Value"], label="GT Value", marker='o')
    ax.plot(x, df_differences["MOD Value"], label="MOD Value", marker='x')

    ax.set_title(f"{measurement_name}(GT) vs {measurement_name}(MOD) Value Comparison")
    ax.set_xlabel("Sample Index")
    ax.set_ylabel("Value")
    ax.legend()

    ax.set_xticks(x[::max(1, len(x)//20)])
    ax.set_xticklabels([str(i+1) for i in x[::max(1, len(x)//20)]], rotation=0)

    ax.tick_params(axis='x', labelsize=8)

    plt.tight_layout()

    img_data = BytesIO()
    plt.savefig(img_data, format='png')
    plt.close()
    img_data.seek(0)

    return img_data

def create_statistics_chart_image(stats_gt, stats_mod, measurement_name):
    df = pd.DataFrame({
        "Metric": ["Mean", "Variance", "Max", "Min", "Median", "5%", "95%"],
        "GT": [stats_gt[0][key] for key in ["Mean", "Variance", "Max", "Min", "Median", "5%", "95%"]],
        "MOD": [stats_mod[0][key] for key in ["Mean", "Variance", "Max", "Min", "Median", "5%", "95%"]],
    })

    x = range(len(df["Metric"]))

    fig, ax = plt.subplots(figsize=(7, 5))
    bar_width = 0.35

    bars_gt = ax.bar([i - bar_width / 2 for i in x], df["GT"], width=bar_width, label="GT")
    bars_mod = ax.bar([i + bar_width / 2 for i in x], df["MOD"], width=bar_width, label="MOD")


    for bar in bars_gt:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, height, f'{height:.2f}',
                ha='center', va='bottom', fontsize=8)

    for bar in bars_mod:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, height, f'{height:.2f}',
                ha='center', va='bottom', fontsize=8)

    ax.set_xlabel("Metric")
    ax.set_ylabel("Value")
    ax.set_title(f"Statistic Comparison for {measurement_name}")
    ax.set_xticks(x)
    ax.set_xticklabels(df["Metric"], rotation=0, ha='center', fontsize=9)
    ax.legend()
    plt.tight_layout()

    img_data = BytesIO()
    plt.savefig(img_data, format='png')
    plt.close()
    img_data.seek(0)

    return img_data

def create_statistics(writer, stats_gt, stats_mod, start_col, measurement_name, threshold):
    worksheet = writer.sheets.get("Statistics")
    if not worksheet:
        worksheet = writer.book.add_worksheet("Statistics")
        writer.sheets["Statistics"] = worksheet

    header_text = f"Statistics for {measurement_name}"
    header_format = writer.book.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': 'white',
        'bg_color': '#4BACC6',
        'border': 1
    })

    center_format = writer.book.add_format({
        'align': 'center',
        'valign': 'vcenter'
    })
    left_format = writer.book.add_format({
        'align': 'left',
        'valign': 'vcenter'
    })

    metric_keys = ["Mean", "Variance", "Max", "Min", "Median", "5%", "95%"]
    table_end_col = start_col + 4

    worksheet.merge_range(0, start_col, 0, table_end_col, header_text, header_format)
    worksheet.set_row(0, 30)

    worksheet.write(1, start_col, "Metric", center_format)
    worksheet.write(1, start_col + 1, f"{measurement_name} (GT)", center_format)
    worksheet.write(1, start_col + 2, f"{measurement_name} (MOD)", center_format)
    worksheet.write(1, start_col + 3, "% Difference", center_format)
    worksheet.write(1, start_col + 4, "Threshold (%)", center_format)

    for row_idx, metric in enumerate(metric_keys, start=2):
        gt_value = stats_gt[0][metric]
        mod_value = stats_mod[0][metric]
        try:
            diff_percent = abs((mod_value - gt_value) / gt_value) * 100 if gt_value != 0 else 0
        except Exception:
            diff_percent = 0

        worksheet.write(row_idx, start_col, metric, left_format)
        worksheet.write(row_idx, start_col + 1, stats_gt[0][metric], center_format)
        worksheet.write(row_idx, start_col + 2, stats_mod[0][metric], center_format)
        worksheet.write(row_idx, start_col + 3, round(diff_percent, 2), center_format)
        worksheet.write(row_idx, start_col + 4, threshold, center_format)

    worksheet.add_table(1, start_col, row_idx, table_end_col, {
         'columns': [
             {'header': "Metric"},
             {'header': f"{measurement_name} (GT)"},
             {'header': f"{measurement_name} (MOD)"},
             {'header': "Difference (%)"},
             {'header': "Threshold (%)"},
         ],
         'name': f'Stats_{measurement_name.replace(" ", "_")}'
     })

    for i in range(start_col, table_end_col + 1):
        worksheet.set_column(i, i, 18, center_format)

    gt_letter = xl_col_to_name(start_col + 1)
    mod_letter = xl_col_to_name(start_col + 2)
    first_data_row = 3
    last_data_row = row_idx + 1

    red = writer.book.add_format({'bg_color': '#FFC7CE'})
    green = writer.book.add_format({'bg_color': '#C6EFCE'})

    worksheet.conditional_format(
        f'{mod_letter}{first_data_row}:{mod_letter}{last_data_row}',
        {
            'type': 'formula',
            'criteria': (
                f'=AND(ABS({mod_letter}{first_data_row}-{gt_letter}{first_data_row})/ABS({gt_letter}{first_data_row})*100 > {threshold}, '
                f'{mod_letter}{first_data_row} > {gt_letter}{first_data_row})'
            ),
            'format': red
        }
    )

    worksheet.conditional_format(
        f'{mod_letter}{first_data_row}:{mod_letter}{last_data_row}',
        {
            'type': 'formula',
            'criteria': (
                f'=AND(ABS({mod_letter}{first_data_row}-{gt_letter}{first_data_row})/ABS({gt_letter}{first_data_row})*100 > {threshold}, '
                f'{mod_letter}{first_data_row} < {gt_letter}{first_data_row})'
            ),
            'format': green
        }
    )
    img_data = create_statistics_chart_image(stats_gt, stats_mod, measurement_name)
    chart_row = row_idx + 4
    worksheet.insert_image(chart_row, start_col, 'stat_chart.png', {'image_data': img_data})

    return table_end_col + 2


def create_report(measurements_diff, writer, start_col, col_name, threshold):


    df = pd.DataFrame(measurements_diff)
    chart_image = create_chart_image(df, col_name)

    header_text = f"Comparison: {col_name} (GT) vs {col_name} (MOD)"
    header_format = writer.book.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_color': 'white',
        'bg_color': '#4F81BD',
        'border': 1
    })

    worksheet = writer.sheets.get("Comparison Results")
    if not worksheet:
        worksheet = writer.book.add_worksheet("Comparison Results")
        writer.sheets["Comparison Results"] = worksheet

    max_row, max_col = df.shape
    table_end_col = start_col + max_col - 1

    worksheet.merge_range(0, start_col, 0, table_end_col, header_text, header_format)
    worksheet.set_row(0, 30)

    for col_idx, col in enumerate(df.columns):
        worksheet.write(1, start_col + col_idx, col)
        for row_idx, val in enumerate(df[col]):
            if pd.isna(val) or val in [float('inf'), float('-inf')]:
                worksheet.write_blank(row_idx + 2, start_col + col_idx, None)
            else:
                worksheet.write(row_idx + 2, start_col + col_idx, val)

    worksheet.add_table(1, start_col, max_row + 1, table_end_col, {
        'columns': [{'header': col} for col in df.columns],
        'name': f'Table_{start_col}'
    })

    center_format = writer.book.add_format({'align': 'center'})
    for i in range(1, max_col):
        worksheet.set_column(start_col + i, start_col + i,  15, center_format)
    worksheet.set_column(start_col, start_col, 15)

    try:
        gt_col = df.columns.get_loc("GT Value")
        mod_col = df.columns.get_loc("MOD Value")
        diff_col = df.columns.get_loc("% Difference")
        issue_col = df.columns.get_loc("Issue")

        grey = writer.book.add_format({'bg_color': '#D9D9D9'})
        red = writer.book.add_format({'bg_color': '#FFC7CE'})
        green = writer.book.add_format({'bg_color': '#C6EFCE'})

        diff_letter = xl_col_to_name(start_col + diff_col)
        mod_letter = xl_col_to_name(start_col + mod_col)
        gt_letter = xl_col_to_name(start_col + gt_col)

        yellow = writer.book.add_format({'bg_color': '#FFEB9C'})
        issue_letter = xl_col_to_name(start_col + issue_col)
        worksheet.conditional_format(f'{mod_letter}3:{mod_letter}{max_row + 2}', {
            'type': 'formula',
            'criteria': (
                f'=OR(${issue_letter}3="new sample in MOD", '
                f'${issue_letter}3="missing sample in MOD")'
            ),
            'format': yellow
        })

        for col_offset in [gt_col, mod_col]:
            col_letter = xl_col_to_name(start_col + col_offset)
            worksheet.conditional_format(f'{col_letter}3:{col_letter}{max_row + 2}', {
                'type': 'blanks',
                'format': grey
            })

        worksheet.conditional_format(f'{mod_letter}3:{mod_letter}{max_row + 2}', {
            'type': 'formula',
            'criteria': f'${issue_letter}3="above threshold"',
            'format': red
        })

        worksheet.conditional_format(f'{mod_letter}3:{mod_letter}{max_row + 2}', {
            'type': 'formula',
            'criteria': f'=AND(${diff_letter}3 > {threshold}, ${mod_letter}3 > ${gt_letter}3)',
            'format': red
        })

        worksheet.conditional_format(f'{mod_letter}3:{mod_letter}{max_row + 2}', {
            'type': 'formula',
            'criteria': f'=AND(${diff_letter}3 > {threshold}, ${mod_letter}3 < ${gt_letter}3)',
            'format': green
        })



        worksheet.set_column(start_col + issue_col, start_col + issue_col, 25, center_format)

    except Exception as e:
        print(f"Conditional formatting error: {e}")

    chart_row = max_row + 5
    worksheet.insert_image(chart_row, start_col, 'chart.png', {'image_data': chart_image})

    return table_end_col + 2

