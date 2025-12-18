# pip install matplotlib

import matplotlib.pyplot as plt

# ---- Data (tCO2e) ----
scope1 = 5.34
scope2 = 148.11   # Electricity
scope3 = 0.00

labels = ["Scope 1", "Scope 2 (Electricity)", "Scope 3"]
sizes = [scope1, scope2, scope3]
colors = ["#6BA292", "#007A3D", "#C1C1C1"]  # 深綠主色系，灰色代表未揭露

# ---- Create pie chart ----
fig = plt.figure(figsize=(7, 5), dpi=150)
wedges, texts, autotexts = plt.pie(
    sizes,
    labels=labels,
    autopct=lambda p: f"{p:.1f}%\n({p/100*sum(sizes):.2f} t)",
    startangle=90,
    colors=colors,
    textprops={"fontsize": 10, "color": "black"}
)

# ---- Title ----
plt.title("Annual GHG Emissions Share (tCO₂e)", fontsize=13, fontweight='bold')
plt.axis('equal')  # 保持圓形比例

# ---- Save ----
plt.savefig("ghg_pie_scope1_2_3.png", bbox_inches="tight", dpi=300)
plt.show()
