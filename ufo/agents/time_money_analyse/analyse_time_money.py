import json
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os




def load_and_merge_data(analysis_file, costs_file):
    """Loads and merges data from the analysis and cost files."""
    df_analysis = pd.read_json(analysis_file)

    costs_list = []
    with open(costs_file, 'r', encoding='utf-8') as f:
        for line in f:
            data = json.loads(line)
            costs_list.append({
                'file_name': data['file_name'],
                'total_money': data['money_cost']['total_money'],
                'total_time': data['time_cost']['total_time']
            })
    df_costs = pd.DataFrame(costs_list)

    df_merged = pd.merge(df_analysis, df_costs, left_on='subfolder_name', right_on='file_name', how='inner')
    return df_merged


def analyze_and_plot(df, item_col, category_dict, value_col, output_prefix, sub_bar_color, y_label, output_dir):
    """
    Performs data analysis and generates a nested bar chart.

    Args:
        df (pd.DataFrame): The merged dataframe.
        item_col (str): The column to analyze ('operation_object' or 'operation_type').
        category_dict (dict): Dictionary mapping categories to lists of items.
        value_col (str): The value column to average ('total_time' or 'total_money').
        output_prefix (str): Prefix for output file names.
        sub_bar_color (str): Matplotlib colormap name for sub-bars.
        y_label (str): Label for the Y-axis.
        output_dir (str): The directory to save output files.
    """
    item_to_cat_map = {item: cat for cat, items in category_dict.items() for item in items}
    df_exploded = df.explode(item_col)
    df_exploded['category'] = df_exploded[item_col].map(item_to_cat_map)
    df_exploded.dropna(subset=['category'], inplace=True)

    avg_item_values = df_exploded.groupby(item_col)[value_col].mean().reset_index()
    avg_item_values.rename(columns={value_col: f'avg_{value_col}', item_col: 'item'}, inplace=True)
    avg_item_values['category'] = avg_item_values['item'].map(item_to_cat_map)

    avg_cat_values = avg_item_values.groupby('category')[f'avg_{value_col}'].mean().reset_index()
    sorted_categories = avg_cat_values.sort_values(f'avg_{value_col}', ascending=False)

    final_data = []
    for _, cat_row in sorted_categories.iterrows():
        cat_name = cat_row['category']
        cat_avg = cat_row[f'avg_{value_col}']
        items_in_cat = avg_item_values[avg_item_values['category'] == cat_name]
        sorted_items = items_in_cat.sort_values(f'avg_{value_col}', ascending=False)
        final_data.append({
            'category': cat_name,
            f'avg_{value_col}': cat_avg,
            'items': sorted_items.to_dict('records')
        })

    # --- JSON Output ---
    json_output_path = os.path.join(output_dir, f'{output_prefix}_analysis.json')
    with open(json_output_path, 'w', encoding='utf-8') as f:
        json.dump(final_data, f, indent=4, ensure_ascii=False)
    print(f"Successfully created JSON output: {json_output_path}")

    # --- Plotting ---
    fig, ax = plt.subplots(figsize=(20, 10))
    plt.style.use('seaborn-v0_8-whitegrid')

    all_item_values = [item[f'avg_{value_col}'] for cat in final_data for item in cat['items']]
    if not all_item_values:
        print(f"No data to plot for {output_prefix}. Skipping plot generation.")
        plt.close(fig)
        return

    norm = plt.Normalize(vmin=min(all_item_values) * 0.9, vmax=max(all_item_values) * 1.1)
    cmap_sub = plt.get_cmap(sub_bar_color)

    all_cat_values = [cat[f'avg_{value_col}'] for cat in final_data]
    norm_main = plt.Normalize(vmin=min(all_cat_values) * 0.9, vmax=max(all_cat_values) * 1.1)
    cmap_main = plt.get_cmap('YlOrBr')

    x_ticks, x_tick_labels, current_pos = [], [], 0
    for cat_data in final_data:
        cat_name, cat_avg, items = cat_data['category'], cat_data[f'avg_{value_col}'], cat_data['items']
        num_items, start_pos = len(items), current_pos

        for item in items:
            item_name, item_avg = item['item'], item[f'avg_{value_col}']
            ax.bar(current_pos, item_avg, color=cmap_sub(norm(item_avg)), width=0.8, zorder=3)
            ax.text(current_pos, item_avg + (ax.get_ylim()[1] * 0.01), f'{item_avg:.2f}', ha='center', va='bottom',
                    fontsize=9, zorder=5)
            x_ticks.append(current_pos)
            x_tick_labels.append(item_name)
            current_pos += 1

        if num_items > 0:
            main_bar_center = start_pos + (num_items - 1) / 2.0
            main_bar_color = cmap_main(norm_main(cat_avg))
            ax.bar(main_bar_center, cat_avg, width=num_items, color=main_bar_color, alpha=0.5, zorder=1, align='center',
                   edgecolor='grey', linewidth=1)
            ax.text(main_bar_center, cat_avg + (ax.get_ylim()[1] * 0.01), f'{cat_name}\nAvg: {cat_avg:.2f}',
                    ha='center', va='bottom', fontsize=12, fontweight='bold', zorder=5, color='#333333')
        current_pos += 1

    ax.set_xticks(x_ticks)
    ax.set_xticklabels(x_tick_labels, rotation=45, ha='right', fontsize=10)
    ax.set_ylabel(y_label, fontsize=14, fontweight='bold')
    ax.set_title(f'Average {y_label} by {item_col.replace("_", " ").title()} and Category', fontsize=16,
                 fontweight='bold')
    ax.grid(axis='x')
    plt.tight_layout()

    plot_output_path = os.path.join(output_dir, f'{output_prefix}_chart.png')
    plt.savefig(plot_output_path, dpi=300)
    print(f"Successfully created chart: {plot_output_path}")
    plt.close(fig)


def main():
    """Main function to run the entire analysis pipeline."""
    # --- Configuration ---
    input_dir = './time_money_result'
    output_dir = './analysis_output'
    os.makedirs(output_dir, exist_ok=True)  # Create output directory if it doesn't exist



    # --- File Paths ---
    analysis_file = os.path.join(input_dir, 'folder_analysis_results.json')
    costs_file = os.path.join(input_dir, 'costs_and_times.jsonl')
    object_dict_file = os.path.join(input_dir, 'object_dict.json')
    type_dict_file = os.path.join(input_dir, 'type_dict.json')

    # --- Load Data ---
    print("Loading and merging data...")
    df_merged = load_and_merge_data(analysis_file, costs_file)
    with open(object_dict_file, 'r', encoding='utf-8') as f:
        object_dict = json.load(f)
    with open(type_dict_file, 'r', encoding='utf-8') as f:
        type_dict = json.load(f)
    print("Data loaded successfully.")
    print("-" * 20)

    # --- Run Analyses ---
    # 1. Object vs. Time
    print("Running Analysis 1: Object vs. Average Time")
    analyze_and_plot(df=df_merged, item_col='operation_object', category_dict=object_dict,
                     value_col='total_time', output_prefix='object_time', sub_bar_color='Greens',
                     y_label='Average Time (s)', output_dir=output_dir)
    print("-" * 20)

    # 2. Object vs. Money
    print("Running Analysis 2: Object vs. Average Money")
    analyze_and_plot(df=df_merged, item_col='operation_object', category_dict=object_dict,
                     value_col='total_money', output_prefix='object_money', sub_bar_color='Blues',
                     y_label='Average Cost ($)', output_dir=output_dir)
    print("-" * 20)

    # 3. Type vs. Time
    print("Running Analysis 3: Type vs. Average Time")
    analyze_and_plot(df=df_merged, item_col='operation_type', category_dict=type_dict,
                     value_col='total_time', output_prefix='type_time', sub_bar_color='Greens',
                     y_label='Average Time (s)', output_dir=output_dir)
    print("-" * 20)

    # 4. Type vs. Money
    print("Running Analysis 4: Type vs. Average Money")
    analyze_and_plot(df=df_merged, item_col='operation_type', category_dict=type_dict,
                     value_col='total_money', output_prefix='type_money', sub_bar_color='Blues',
                     y_label='Average Cost ($)', output_dir=output_dir)
    print("-" * 20)

    print(f"All analyses complete. Results are in the '{output_dir}' folder.")


if __name__ == '__main__':
    main()
