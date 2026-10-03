"""Create a deliberately messy sales file (no dataset was provided with the task)."""
import numpy as np, pandas as pd
rng = np.random.default_rng(11)
N = 1500
P = {"Laptop":("Electronics",900),"Phone":("Electronics",650),"Headphones":("Electronics",110),
     "Desk":("Furniture",320),"Chair":("Furniture",170),"Notebook":("Stationery",8),
     "Pen Pack":("Stationery",15),"Sneakers":("Apparel",80),"Jacket":("Apparel",120)}
names = ["Aarav Mehta","Priya Nair","John Smith","Sara Khan","Li Wei","Maria Garcia","Omar Ali","Emma Brown","Rahul Verma","Nina Patel"]
dates = pd.to_datetime("2025-01-01") + pd.to_timedelta(rng.integers(0, 640, N), unit="D")
prod = rng.choice(list(P), N)
df = pd.DataFrame({"Order_ID": [f"ORD{i:05d}" for i in range(1, N+1)], "Order_Date": dates,
    "Customer_Name": rng.choice(names, N), "Product": prod,
    "Category": [P[p][0] for p in prod], "Quantity": rng.integers(1, 6, N).astype(float),
    "Unit_Price": [round(P[p][1]*rng.uniform(.9,1.1),2) for p in prod],
    "Region": rng.choice(["North","South","East","West"], N),
    "Payment_Method": rng.choice(["Credit Card","Debit Card","UPI","Cash"], N),
    "Status": rng.choice(["Completed","Returned","Cancelled"], N, p=[.85,.08,.07])})
df["Email"] = df.Customer_Name.str.lower().str.replace(" ", ".") + "@example.com"
# ---- make it messy ----
fmt = rng.choice(["%Y-%m-%d","%d-%b-%Y","%m/%d/%Y"], N)
df["Order_Date"] = [d.strftime(f) for d, f in zip(df.Order_Date, fmt)]                 # mixed date formats
df["Customer_Name"] = [rng.choice([n, n.upper(), n.lower(), "  "+n+" "]) if rng.random()<.35 else n for n in df.Customer_Name]
df["Region"] = [rng.choice([r.lower(), r.upper(), r+" ", r[0]+"."]) if rng.random()<.3 else r for r in df.Region]
df["Payment_Method"] = [rng.choice(["credit card","CC","Credit-Card","UPI ","upi","cash","CASH"]) if rng.random()<.2 and p!="Debit Card" else p for p in df.Payment_Method]
df["Status"] = [rng.choice(["completed","COMPLETED","Done","returned","cancelled"]) if rng.random()<.15 and s=="Completed" else s for s in df.Status]
df["Unit_Price"] = df.Unit_Price.astype(object)
df["Unit_Price"] = [f"${v:,.2f}" if rng.random()<.4 else v for v in df.Unit_Price]  # currency text
for c, frac in [("Quantity",.04),("Unit_Price",.04),("Customer_Name",.03),("Region",.03),("Email",.05)]:
    df.loc[rng.choice(N, int(N*frac), replace=False), c] = np.nan                      # missing values
df.loc[rng.choice(N, 12, replace=False), "Quantity"] = -1                               # invalid values
df.loc[rng.choice(N, 6, replace=False), "Email"] = "not-an-email"
df.loc[rng.choice(N, 3, replace=False), "Order_Date"] = np.nan
df = pd.concat([df, df.sample(40, random_state=2)])                                      # exact duplicates
upd = df.sample(15, random_state=5).copy(); upd["Status"] = "Returned"; df = pd.concat([df, upd])  # same ID, updated
df = df.sample(frac=1, random_state=3).reset_index(drop=True)
df.to_csv("data/raw_sales.csv", index=False); print(df.shape)
