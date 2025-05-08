import pandas as pd
import itertools

input_file_path = "PythonSampleData0205250.xlsx"
output_file_path = "output.xlsx"

df = pd.read_excel(input_file_path, sheet_name="Input")
df.columns = df.columns.str.strip()  

best_groups = {}

for customer_id in df["Customer ID"].unique():
    customer_df = df[df["Customer ID"] == customer_id].reset_index(drop=True)
    n = len(customer_df)
    binary_combos = list(itertools.product([0, 1], repeat=n))
    binary_combos = [combo for combo in binary_combos if any(combo)]

    for combo in binary_combos:
        total = 0
        group_rows = []
        for i, use_amount in enumerate(combo):
            amount = customer_df.iloc[i]["Amount"] if use_amount else 0
            total += amount
            if use_amount:
                group_rows.append({
                    "Customer ID": customer_id,
                    "Invoice Number": customer_df.iloc[i]["Invoice Number"],
                    "Date": pd.to_datetime(customer_df.iloc[i]["Date"]).date(),
                    "Amount": customer_df.iloc[i]["Amount"],
                    "Sum": None
                })

        if group_rows:
            for r in group_rows:
                r["Sum"] = total

            group_key = (customer_id, total)
            earliest_date = min(r["Date"] for r in group_rows)

            if group_key not in best_groups:
                best_groups[group_key] = (earliest_date, group_rows)
            else:
                existing_date = best_groups[group_key][0]
                if earliest_date < existing_date:
                    best_groups[group_key] = (earliest_date, group_rows)

output2_rows = [r for _, group in best_groups.values() for r in group]
output2_df = pd.DataFrame(output2_rows)
output2_df = output2_df[["Customer ID", "Sum", "Invoice Number", "Date", "Amount"]]
output2_df.rename(columns={"Sum": "Key Amount"}, inplace=True)

with pd.ExcelWriter(output_file_path, engine="openpyxl") as writer:
    output2_df.to_excel(writer, index=False, sheet_name="Output2")

print(f"Filtered Output2 written to {output_file_path}")
