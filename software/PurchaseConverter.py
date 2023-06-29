import pandas as pd


def purchaseConverter():
    # In[22]:

    old_purchase_path = input("Enter *OLD PURCHASE* File Path: ")

    # Loading old purchase file
    old_purchase = pd.read_excel(old_purchase_path)

    # Getting party name
    party_name = old_purchase.columns.tolist()[0]

    # dropping first row for column names
    old_purchase.columns = old_purchase.iloc[0].values
    old_purchase = old_purchase[1:].reset_index(drop=True)

    # In[29]:

    # Renaming column names matching new purchase file
    old_purchase.rename(
        columns={
            "Date": "Date",
            "Party Name": "Particulars",
            "VCode": "Voucher Type",
            "Bill No.": "Supplier Invoice No.",
            "Bill Date": "Supplier Invoice Date",
            "Gst No.": "GSTIN/UIN",
            "Bill Amt": "Gross Total",
            "Basic": "BASIC",
            "CGST Amt": "CGST",
            "SGST Amt": "SGST",
            "IGST Amt": "IGST",
        },
        inplace=True,
    )

    # In[30]:

    # Exporting new purchase

    file_name = f"{party_name} New Purchase ITC Reco..xlsx"

    with pd.ExcelWriter(file_name, mode="w") as writer:
        old_purchase.to_excel(
            writer, sheet_name="Purchase Register", index=False, startrow=3
        )
