import pandas as pd

# Data for the test import file
data = {
    'mpn': ['NEW-MPN-1', 'NEW-MPN-2', 'CONFLICT-MPN-1'],
    'cpn': ['NEW-CPN-A', 'NEW-CPN-B', 'SHOULD-BE-SKIPPED']
}

df = pd.DataFrame(data)
df.to_excel('test_import.xlsx', index=False)

print("Test import file 'test_import.xlsx' created successfully.")
