import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import statsmodels.api as sm
from statsmodels.compat import lzip
import statsmodels.formula.api as smf

print("\nSTATS MODELS | OLS Linear Regression")

# ----------------------------- DATASET
df = sm.datasets.get_rdataset("Guerry", "HistData").data
print(f'{df.columns}\n')

# ----------------------------- COLUMNS
dep = 'Suicides' # dep. variable
col1 = 'Wealth'
col2 = 'Prostitutes'
col3 = 'Distance'
col4 = 'Literacy'

# ----------------------------- MODEL
df2 = df[[dep, col1, col2, col3, col4]].dropna()
y = df2[[dep]]
x = df2[[col1, col2, col3, col4]]
print(f'{round(df2.describe())}\n')

# ----------------------------- FIT
mod = sm.OLS(y, x).fit()
print(mod.summary())

# ----------------------------- GRAPHS | PLOTS
exog_idx = col1
plt.rc("figure", figsize=(12, 8))
plt.rc("font", size=10)

# ----------------------------- Single Variable Regression Diagnostics
fig = sm.graphics.plot_regress_exog(mod, exog_idx)
fig.tight_layout(pad=1.0)
plt.show()

# ----------------------------- Fit Plot
fig = sm.graphics.plot_fit(mod, exog_idx)
fig.tight_layout(pad=1.0)
plt.show()

# ----------------------------- Influence Plot
fig = sm.graphics.influence_plot(mod, exog_idx)
fig.tight_layout(pad=1.0)
# plt.show()
