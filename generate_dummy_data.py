"""Generates the synthetic WD inventory dataset (seeded, reproducible)."""
import random
import pandas as pd

random.seed(42)

locations = ["Jamshedpur", "Ranchi", "Kolkata", "Patna", "Bhubaneswar", "Lucknow", "Guwahati", "Raipur"]
catalog = {
    "Beverages":      ["Mango Fruit Drink 200ml", "Packaged Water 1L", "Iced Tea 250ml", "Energy Drink 250ml"],
    "Packaged Foods": ["Instant Noodles 70g", "Basmati Rice 5kg", "Refined Oil 1L", "Atta 10kg"],
    "Personal Care":  ["Shampoo 340ml", "Toothpaste 150g", "Body Wash 250ml", "Face Wash 100g"],
    "Home Care":      ["Detergent Powder 1kg", "Dish Wash Liquid 500ml", "Floor Cleaner 1L", "Toilet Cleaner 500ml"],
    "Snacks":         ["Potato Chips 52g", "Cream Biscuits 120g", "Namkeen Mix 200g", "Cookies 150g"],
    "Dairy & Frozen": ["UHT Milk 1L", "Paneer 200g", "Frozen Peas 500g", "Ice Cream Tub 500ml"],
}

# Target mix: ~14 Critical, ~12 Warning, ~24 Safe (verified after generation)
plan = ["C"] * 14 + ["W"] * 12 + ["S"] * 24
random.shuffle(plan)

rows = []
used = set()
for i, tag in enumerate(plan):
    cat = random.choice(list(catalog))
    loc = random.choice(locations)
    sku = f"SKU-{1001 + i * 7}"
    lt = random.randint(2, 9)
    demand = random.choice([18, 25, 32, 40, 55, 70, 90, 120, 150, 200, 260])
    if tag == "C":
        cover = random.uniform(0.4, lt - 0.3)
    elif tag == "W":
        cover = random.uniform(lt + 0.3, lt + 1.8)
    else:
        cover = random.uniform(lt + 3.0, lt + 14.0)
    stock = max(1, round(demand * cover))
    p_delay = 0.65 if tag == "C" else 0.35 if tag == "W" else 0.12
    otif = "Delayed" if random.random() < p_delay else "On-Time"
    rows.append([loc, sku, cat, stock, demand, lt, otif])

df = pd.DataFrame(rows, columns=["WD_Location", "SKU_ID", "Product_Category", "Current_Stock_Qty",
                                 "Avg_Daily_Demand", "Transit_Lead_Time_Days", "Last_Delivery_OTIF_Status"])
df.to_csv("wd_inventory_data.csv", index=False)

# sanity check of resulting mix
doc = df.Current_Stock_Qty / df.Avg_Daily_Demand
lt = df.Transit_Lead_Time_Days
risk = ["Critical Risk" if d <= l else "Warning" if d <= l + 2 else "Safe" for d, l in zip(doc, lt)]
print(pd.Series(risk).value_counts())
print(df.shape)
