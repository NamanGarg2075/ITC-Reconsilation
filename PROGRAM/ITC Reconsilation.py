import pandas as pd
import numpy as np
import time


def exepy():
    #!/usr/bin/env python
    # coding: utf-8

    # In[2]:

    pur_path = input("Enter *Purchase* File Path: ")
    b2b_path = input("Enter *B2B* File Path: ")

    last_month_input = input("Do you want to add *LAST MONTH* file (Y/N): ")

    if last_month_input.lower() == "y":
        last_month_path = input("Enter *LAST MONTH* File Name: ")

    # starting time
    start_time = time.time()

    # loading purchase books file
    purchase = pd.read_excel(pur_path, skiprows=3)

    # loading B2B file
    b2b = pd.read_excel(b2b_path, sheet_name="B2B", skiprows=4)
    b2b.drop(0, inplace=True)

    # Loading NOTE sheet
    b2b_cdnr = pd.read_excel(b2b_path, sheet_name="B2B-CDNR", skiprows=4)
    b2b_cdnr.drop(0, inplace=True)

    if last_month_input.lower() == "y":
        # Loading LAST MONTH PENDING and EXTRA sheet
        pending_last_month = pd.read_excel(last_month_path, sheet_name="PENDING")
        extra_last_month = pd.read_excel(last_month_path, sheet_name="EXTRA")

    # getting current month of file
    current_monn_file = pd.read_excel(b2b_path, sheet_name="Read me", skiprows=3).iloc[
        :, 2
    ][0]
    current_month_name = current_monn_file

    # getting file gst number
    file_gst_no = pd.read_excel(b2b_path, sheet_name="Read me", skiprows=3).iloc[:, 2][
        1
    ]

    # In[3]:

    # Changing LAST MONTH RECO file Voucher Date to Date
    if last_month_input.lower() == "y":
        pending_last_month.rename(columns={"Voucher Date": "Date"}, inplace=True)
        extra_last_month.rename(columns={"Voucher Date": "Date"}, inplace=True)

    # In[4]:

    # Renaming columns of B2B and B2B CDNR
    b2b.rename(
        columns={
            "Invoice Details": "Invoice number",
            "Unnamed: 3": "Invoice type",
            "Unnamed: 4": "Invoice Date",
            "Unnamed: 5": "Invoice Value(₹)",
            "Tax Amount": "Integrated Tax(₹)",
            "Unnamed: 11": "Central Tax(₹)",
            "Unnamed: 12": "State/UT Tax(₹)",
            "Unnamed: 13": "Cess(₹)",
        },
        inplace=True,
    )
    b2b.reset_index(drop=True, inplace=True)

    b2b_cdnr.rename(
        columns={
            "Credit note/Debit note details": "Invoice number",
            "Unnamed: 3": "Invoice type",
            "Unnamed: 4": "Note Supply type",
            "Unnamed: 5": "Invoice Date",
            "Unnamed: 6": "Invoice Value(₹)",
            "Tax Amount": "Integrated Tax(₹)",
            "Unnamed: 12": "Central Tax(₹)",
            "Unnamed: 13": "State/UT Tax(₹)",
            "Unnamed: 14": "Cess(₹)",
        },
        inplace=True,
    )
    b2b_cdnr.reset_index(drop=True, inplace=True)

    if last_month_input.lower() == "y":
        # Renaming LAST MONTH DATA SUPPLIER NAME
        pending_last_month.rename(columns={"SUPPLIER NAME": "Trade Name"}, inplace=True)
        extra_last_month.rename(columns={"SUPPLIER NAME": "Trade Name"}, inplace=True)

    # In[5]:

    # Dropping 'Regular' column from b2b_cdnr
    b2b_cdnr.drop(columns={"Note Supply type"}, inplace=True)

    # In[6]:

    # converting to negative values of 'Credit Note'
    for i in range(b2b_cdnr.shape[0]):
        if b2b_cdnr.loc[i, "Invoice type"] == "Credit Note":
            b2b_cdnr.loc[i, "Invoice Value(₹)"] = 0 - b2b_cdnr["Invoice Value(₹)"][i]
            b2b_cdnr.loc[i, "Taxable Value (₹)"] = 0 - b2b_cdnr["Taxable Value (₹)"][i]
            b2b_cdnr.loc[i, "Integrated Tax(₹)"] = 0 - b2b_cdnr["Integrated Tax(₹)"][i]
            b2b_cdnr.loc[i, "Central Tax(₹)"] = 0 - b2b_cdnr["Central Tax(₹)"][i]
            b2b_cdnr.loc[i, "State/UT Tax(₹)"] = 0 - b2b_cdnr["State/UT Tax(₹)"][i]

    # In[7]:

    # Concatinating b2b_cdnr data and B2B to B2B
    b2b = pd.concat([b2b, b2b_cdnr], ignore_index=True)

    # Taking only those rows with 'No' values
    b2b = b2b[b2b["Supply Attract Reverse Charge"] == "No"]

    # In[8]:

    # Getting Required Data
    purchase = purchase[
        [
            "Date",
            "Particulars",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "GSTIN/UIN",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
        ]
    ]
    b2b = b2b.iloc[:, :13]

    # In[9]:

    # Changing PURCHASE columns data type
    purchase[["Gross Total", "BASIC", "CGST", "SGST", "IGST"]] = purchase[
        ["Gross Total", "BASIC", "CGST", "SGST", "IGST"]
    ].astype("float64")
    purchase["Date"] = pd.to_datetime(purchase["Date"])
    purchase["Supplier Invoice Date"] = pd.to_datetime(
        purchase["Supplier Invoice Date"]
    )

    # Changing B2B columns data type
    b2b[
        [
            "Invoice Value(₹)",
            "Taxable Value (₹)",
            "Integrated Tax(₹)",
            "Central Tax(₹)",
            "State/UT Tax(₹)",
        ]
    ] = b2b[
        [
            "Invoice Value(₹)",
            "Taxable Value (₹)",
            "Integrated Tax(₹)",
            "Central Tax(₹)",
            "State/UT Tax(₹)",
        ]
    ].astype(
        "float64"
    )
    b2b["Invoice Date"] = pd.to_datetime(b2b["Invoice Date"], dayfirst=True)

    # In[10]:

    # Renaming columns
    purchase.rename(columns={"Particulars": "Trade Name"}, inplace=True)

    # Adding Columns
    REMARKS = ""
    STATUS = ""
    purchase.insert(2, "REMARKS", REMARKS)
    purchase.insert(12, "STATUS", STATUS)

    # In[11]:

    # Filling REMARKS column
    purchase["REMARKS"] = "AS PER BOOKS"

    # In[12]:

    # Creating Dummy Dataframe for concatinating
    b2b_df = pd.DataFrame()

    b2b_df["Date"] = b2b["Invoice Date"]
    b2b_df["Trade Name"] = b2b["Trade/Legal name"]
    b2b_df["REMARKS"] = "AS PER 2B"
    b2b_df["Voucher Type"] = b2b["Invoice type"]
    b2b_df["Supplier Invoice No."] = b2b["Invoice number"]
    b2b_df["Supplier Invoice Date"] = b2b["Invoice Date"]
    b2b_df["GSTIN/UIN"] = b2b["GSTIN of supplier"]
    b2b_df["Gross Total"] = ""
    b2b_df["BASIC"] = b2b["Taxable Value (₹)"]
    b2b_df["CGST"] = b2b["Central Tax(₹)"]
    b2b_df["SGST"] = b2b["State/UT Tax(₹)"]
    b2b_df["IGST"] = b2b["Integrated Tax(₹)"]
    b2b_df["STATUS"] = ""

    # In[13]:

    if last_month_input.lower() == "y":
        # Adding B2B data below PURCHASE
        main_data = pd.concat(
            [purchase, b2b_df, pending_last_month, extra_last_month], ignore_index=True
        )
    else:
        main_data = pd.concat([purchase, b2b_df], ignore_index=True)

    # In[14]:

    # Setting GST NO column
    main_data["GSTIN/UIN"] = main_data["GSTIN/UIN"].str.replace("_x000D_\n", "")
    main_data["GSTIN/UIN"] = main_data["GSTIN/UIN"].str.replace("\n", "")

    # In[15]:

    # Calculating GROSS column
    main_data["Gross Total"] = (
        main_data["BASIC"] + main_data["CGST"] + main_data["SGST"] + main_data["IGST"]
    )

    # In[16]:

    # Creating SUPPLIER NAME column
    supplier_name = ""
    main_data.insert(2, "SUPPLIER NAME", supplier_name)

    # In[17]:

    # Separating 'AS PER BOOKS' and 'AS PER 2B' from main_data
    main_asp_books = main_data[main_data["REMARKS"] == "AS PER BOOKS"]
    main_asp_2b = main_data[main_data["REMARKS"] == "AS PER 2B"]

    # In[18]:

    # Grouping on the basis of GSTIN
    books_gstin_join = (
        main_asp_books.groupby("GSTIN/UIN", dropna=False)
        .agg({"Trade Name": "first"})
        .reset_index()
    )

    # Performing join on both above data
    main_asp_2b = pd.merge(main_asp_2b, books_gstin_join, how="left", on="GSTIN/UIN")
    main_asp_2b["Trade Name_y"].fillna(main_asp_2b["Trade Name_x"], inplace=True)

    # In[19]:

    # Transfering column [Supplier Name]
    main_asp_2b.loc[:, "SUPPLIER NAME"] = main_asp_2b["Trade Name_y"]
    main_asp_books.loc[:, "SUPPLIER NAME"] = main_asp_books["Trade Name"]

    main_data = pd.concat([main_asp_books, main_asp_2b], ignore_index=True)
    main_data.drop(columns=["Trade Name", "Trade Name_x", "Trade Name_y"], inplace=True)

    # In[ ]:

    # In[ ]:

    # In[20]:

    # Filling NAN GSTIN

    gstin_join_2b = (
        main_asp_2b.groupby(["SUPPLIER NAME", "GSTIN/UIN"])["BASIC"].sum().reset_index()
    )

    gstin_join_2b = gstin_join_2b.drop(columns=["BASIC"])

    gstin_join_books = (
        main_asp_books.groupby(["SUPPLIER NAME", "GSTIN/UIN"])["BASIC"]
        .sum()
        .reset_index()
    )

    gstin_join_books = gstin_join_books.drop(columns=["BASIC"])

    main_data = pd.merge(main_data, gstin_join_2b, how="outer", on="SUPPLIER NAME")

    main_data["GSTIN/UIN_y"].fillna(main_data["GSTIN/UIN_x"], inplace=True)

    main_data["GSTIN/UIN_x"] = main_data["GSTIN/UIN_y"]

    main_data.drop(columns=["GSTIN/UIN_y"], inplace=True)

    main_data.rename(columns={"GSTIN/UIN_x": "GSTIN/UIN"}, inplace=True)

    # In[21]:

    # Extracting last 2 character of bill Number
    main_data["Bill No"] = main_data["Supplier Invoice No."].apply(
        lambda x: str(x)[-3:]
    )

    # Getting Month and Year of each row
    main_data["Month"] = main_data["Supplier Invoice Date"].dt.month
    main_data["Year"] = main_data["Supplier Invoice Date"].dt.year

    # In[22]:

    # Creating Pivot Table

    # Separating 'AS PER BOOKS' and 'AS PER 2B'
    aspbooks = main_data[main_data["REMARKS"] == "AS PER BOOKS"].reset_index(drop=True)
    asp2b = main_data[main_data["REMARKS"] == "AS PER 2B"].reset_index(drop=True)

    # Creating Pivot for both
    aspbooks_pivot = (
        aspbooks.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    asp2b_pivot = (
        asp2b.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )

    # In[23]:

    # Merging Both Pivot Table
    merged_df = pd.merge(
        asp2b_pivot,
        aspbooks_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Renaming Columns Name
    merged_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[24]:

    # Filling 'NaN' values with '0'
    merged_df["2B BASIC"] = merged_df["2B BASIC"].fillna(0)
    merged_df["2B CGST"] = merged_df["2B CGST"].fillna(0)
    merged_df["2B SGST"] = merged_df["2B SGST"].fillna(0)
    merged_df["2B IGST"] = merged_df["2B IGST"].fillna(0)
    merged_df["BOOKS BASIC"] = merged_df["BOOKS BASIC"].fillna(0)
    merged_df["BOOKS CGST"] = merged_df["BOOKS CGST"].fillna(0)
    merged_df["BOOKS SGST"] = merged_df["BOOKS SGST"].fillna(0)
    merged_df["BOOKS IGST"] = merged_df["BOOKS IGST"].fillna(0)

    # In[25]:

    # Calculating Difference
    merged_df["BASIC_diff"] = merged_df["2B BASIC"] - merged_df["BOOKS BASIC"]
    merged_df["CGST_diff"] = merged_df["2B CGST"] - merged_df["BOOKS CGST"]
    merged_df["SGST_diff"] = merged_df["2B SGST"] - merged_df["BOOKS SGST"]
    merged_df["IGST_diff"] = merged_df["2B IGST"] - merged_df["BOOKS IGST"]

    # Addition of All GST's Diff
    merged_df["GST_diff_total"] = (
        merged_df["CGST_diff"] + merged_df["SGST_diff"] + merged_df["IGST_diff"]
    )

    # Rearrancing Columns
    merged_df = merged_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Bill No",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[26]:

    # Creating STATUS column
    merged_df.loc[:, "STATUS"] = ""

    # Filling STATUS column with 'OK', 'PENDING' and 'EXTRA'
    for i in range(merged_df.shape[0]):
        if merged_df.loc[i, "GST_diff_total"] > 9:
            merged_df.loc[i, "STATUS"] = "Extra Claimed"
        elif merged_df.loc[i, "GST_diff_total"] < -9:
            merged_df.loc[i, "STATUS"] = "Pending Claim"
        else:
            merged_df.loc[i, "STATUS"] = "OK"

    # In[27]:

    # Joining Data for STATUS column
    main_data_df = pd.merge(
        main_data,
        merged_df[["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year", "STATUS"]],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Dropping STATUS_x column
    main_data_df.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    main_data_df.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # In[28]:

    # /******************************  LAST MONTH ********************************/
    current_month = pd.Timestamp.now().month
    current_yr = pd.Timestamp.now().year

    # creating LAST CLAIM MONTH column
    main_data_df["LAST CLAIM MONTH"] = ""

    month = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.month
    year = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.year

    data_month = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.month_name()
    data_year = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.year

    # converting current file month to number
    current_monn_file = pd.to_datetime(current_monn_file, format="%B").month

    for i in range(main_data_df.shape[0]):
        if (month[i] < current_monn_file) or (year[i] < current_yr):
            main_data_df.loc[
                i, "LAST CLAIM MONTH"
            ] = f"{data_month[i]} {data_year[i]} CLAIM"

    # In[29]:

    # Conterting datetime to date only for export
    main_data_df["Date"] = pd.to_datetime(main_data_df["Date"]).dt.date
    main_data_df["Supplier Invoice Date"] = pd.to_datetime(
        main_data_df["Supplier Invoice Date"]
    ).dt.date

    # Renaming Date Column
    main_data_df.rename(columns={"Date": "Voucher Date"}, inplace=True)

    # In[30]:

    # Separating pending and extra rows
    pending_data = main_data_df[main_data_df["STATUS"] == "Pending Claim"].reset_index(
        drop=True
    )
    extra_data = main_data_df[main_data_df["STATUS"] == "Extra Claimed"].reset_index(
        drop=True
    )

    # In[31]:

    # Creating Separate Data for both 'Pending' and 'Extra' rows
    pending_asp_book = pending_data[pending_data["REMARKS"] == "AS PER BOOKS"]
    pending_asp_2b = pending_data[pending_data["REMARKS"] == "AS PER 2B"]

    extra_asp_book = extra_data[extra_data["REMARKS"] == "AS PER BOOKS"]
    extra_asp_2b = extra_data[extra_data["REMARKS"] == "AS PER 2B"]

    # In[32]:

    # Creating Pivot for pending both
    pending_asp_books_pivot = (
        pending_asp_book.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    pending_asp_2b_pivot = (
        pending_asp_2b.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    # Creating Pivot for pending both
    extra_asp_books_pivot = (
        extra_asp_book.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    extra_asp_2b_pivot = (
        extra_asp_2b.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )

    # In[33]:

    # Merging Both Pivot Table 'Pending'
    pending_df = pd.merge(
        pending_asp_2b_pivot,
        pending_asp_books_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Renaming Columns Name of 'Pending_df'
    pending_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # Merging Both Pivot Table 'Extra'
    extra_df = pd.merge(
        extra_asp_2b_pivot,
        extra_asp_books_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Renaming Columns Name of 'extra_df'
    extra_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[34]:

    # Calculating Difference
    pending_df["BASIC_diff"] = pending_df["2B BASIC"] - pending_df["BOOKS BASIC"]
    pending_df["CGST_diff"] = pending_df["2B CGST"] - pending_df["BOOKS CGST"]
    pending_df["SGST_diff"] = pending_df["2B SGST"] - pending_df["BOOKS SGST"]
    pending_df["IGST_diff"] = pending_df["2B IGST"] - pending_df["BOOKS IGST"]

    # Addition of All GST's Diff
    pending_df["GST_diff_total"] = (
        pending_df["CGST_diff"] + pending_df["SGST_diff"] + pending_df["IGST_diff"]
    )

    # Rearrancing Columns
    pending_df = pending_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Bill No",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[35]:

    # Calculating Difference
    extra_df["BASIC_diff"] = extra_df["2B BASIC"] - extra_df["BOOKS BASIC"]
    extra_df["CGST_diff"] = extra_df["2B CGST"] - extra_df["BOOKS CGST"]
    extra_df["SGST_diff"] = extra_df["2B SGST"] - extra_df["BOOKS SGST"]
    extra_df["IGST_diff"] = extra_df["2B IGST"] - extra_df["BOOKS IGST"]

    # Addition of All GST's Diff
    extra_df["GST_diff_total"] = (
        extra_df["CGST_diff"] + extra_df["SGST_diff"] + extra_df["IGST_diff"]
    )

    # Rearrancing Columns
    extra_df = extra_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Bill No",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[36]:

    # Pending
    pending_df["STATUS"] = ""

    for i in range(pending_df.shape[0]):
        if (pending_df["GST_diff_total"][i] < 0) or (
            pending_df["GST_diff_total"][i] > 0
        ):
            pending_df.loc[i, "STATUS"] = "Review Required"
        else:
            pending_df.loc[i, "STATUS"] = "Pending Claim"

    # In[37]:

    # Extra
    extra_df["STATUS"] = ""

    for i in range(extra_df.shape[0]):
        if (extra_df["GST_diff_total"][i] < 0) or (extra_df["GST_diff_total"][i] > 0):
            extra_df.loc[i, "STATUS"] = "Review Required"
        else:
            extra_df.loc[i, "STATUS"] = "Extra Claimed"

    # In[38]:

    # Joining Data for STATUS column
    pending_data = pd.merge(
        pending_data,
        pending_df[
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year", "STATUS"]
        ],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Dropping STATUS_x column
    pending_data.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    pending_data.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # Separating Review/Pending rows
    pending_review = pending_data[pending_data["STATUS"] == "Review Required"]
    pending_data = pending_data[pending_data["STATUS"] == "Pending Claim"]

    # In[39]:

    # Joining Data for STATUS column
    extra_data = pd.merge(
        extra_data,
        extra_df[["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year", "STATUS"]],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Dropping STATUS_x column
    extra_data.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    extra_data.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # Separating Review/Extra rows
    extra_review = extra_data[extra_data["STATUS"] == "Review Required"]
    extra_data = extra_data[extra_data["STATUS"] == "Extra Claimed"]

    # In[40]:

    # All Review rows
    review_data = pd.concat([pending_review, extra_review], ignore_index=True)

    # OK status rows
    ok_data = main_data_df[main_data_df["STATUS"] == "OK"]

    # In[41]:

    # Creating main_data_df
    main_data_df = pd.concat(
        [pending_data, extra_data, ok_data, review_data], ignore_index=True
    )

    # In[42]:

    # Creating SUMMARY AGAIN on NAME and GSTN

    # Separating 'AS PER BOOKS' and 'AS PER 2B'
    aspbooks = main_data_df[main_data_df["REMARKS"] == "AS PER BOOKS"].reset_index(
        drop=True
    )
    asp2b = main_data_df[main_data_df["REMARKS"] == "AS PER 2B"].reset_index(drop=True)

    # Creating Pivot for both
    aspbooks_pivot = (
        aspbooks.groupby(["SUPPLIER NAME", "GSTIN/UIN", "Month", "Year"], dropna=False)
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    asp2b_pivot = (
        asp2b.groupby(["SUPPLIER NAME", "GSTIN/UIN", "Month", "Year"], dropna=False)
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )

    # Merging Both Pivot Table
    summary_df = pd.merge(
        asp2b_pivot,
        aspbooks_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Month", "Year"],
    )

    # Renaming Columns Name
    summary_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[43]:

    # Filling 'NaN' values with '0'
    summary_df["2B BASIC"] = summary_df["2B BASIC"].fillna(0)
    summary_df["2B CGST"] = summary_df["2B CGST"].fillna(0)
    summary_df["2B SGST"] = summary_df["2B SGST"].fillna(0)
    summary_df["2B IGST"] = summary_df["2B IGST"].fillna(0)
    summary_df["BOOKS BASIC"] = summary_df["BOOKS BASIC"].fillna(0)
    summary_df["BOOKS CGST"] = summary_df["BOOKS CGST"].fillna(0)
    summary_df["BOOKS SGST"] = summary_df["BOOKS SGST"].fillna(0)
    summary_df["BOOKS IGST"] = summary_df["BOOKS IGST"].fillna(0)

    # Calculating Difference
    summary_df["BASIC_diff"] = summary_df["2B BASIC"] - summary_df["BOOKS BASIC"]
    summary_df["CGST_diff"] = summary_df["2B CGST"] - summary_df["BOOKS CGST"]
    summary_df["SGST_diff"] = summary_df["2B SGST"] - summary_df["BOOKS SGST"]
    summary_df["IGST_diff"] = summary_df["2B IGST"] - summary_df["BOOKS IGST"]

    # Addition of All GST's Diff
    summary_df["GST_diff_total"] = (
        summary_df["CGST_diff"] + summary_df["SGST_diff"] + summary_df["IGST_diff"]
    )

    # Rearrancing Columns
    summary_df = summary_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[44]:

    # Creating STATUS column
    summary_df.loc[:, "STATUS"] = ""

    # Filling STATUS column with 'OK', 'PENDING' and 'EXTRA'
    for i in range(summary_df.shape[0]):
        if summary_df.loc[i, "GST_diff_total"] > 9:
            summary_df.loc[i, "STATUS"] = "Extra Claimed"
        elif summary_df.loc[i, "GST_diff_total"] < -9:
            summary_df.loc[i, "STATUS"] = "Pending Claim"
        else:
            summary_df.loc[i, "STATUS"] = "OK"

    # In[45]:

    # Joining Data for STATUS column
    main_data_df = pd.merge(
        main_data_df,
        summary_df[["SUPPLIER NAME", "GSTIN/UIN", "Month", "Year", "STATUS"]],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Month", "Year"],
    )

    # Dropping STATUS_x column
    main_data_df.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    main_data_df.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # In[46]:

    # /******************************  LAST MONTH ********************************/
    current_month = pd.Timestamp.now().month
    current_yr = pd.Timestamp.now().year

    # creating LAST CLAIM MONTH column
    main_data_df["LAST CLAIM MONTH"] = ""

    month = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.month
    year = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.year

    data_month = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.month_name()
    data_year = pd.to_datetime(main_data_df["Supplier Invoice Date"]).dt.year

    # converting current file month to number
    current_monn_file = pd.to_datetime(current_month_name, format="%B").month

    for i in range(main_data_df.shape[0]):
        if (month[i] < current_monn_file) or (year[i] < current_yr):
            main_data_df.loc[
                i, "LAST CLAIM MONTH"
            ] = f"{data_month[i]} {data_year[i]} CLAIM"

    # In[47]:

    pending_data = main_data_df[main_data_df["STATUS"] == "Pending Claim"]
    extra_data = main_data_df[main_data_df["STATUS"] == "Extra Claimed"]

    # In[ ]:

    # In[ ]:

    # In[48]:

    pending_asp_book = pending_data[pending_data["REMARKS"] == "AS PER BOOKS"]
    pending_asp_2b = pending_data[pending_data["REMARKS"] == "AS PER 2B"]

    extra_asp_book = extra_data[extra_data["REMARKS"] == "AS PER BOOKS"]
    extra_asp_2b = extra_data[extra_data["REMARKS"] == "AS PER 2B"]

    # In[49]:

    # Creating Pivot for pending both
    pending_asp_books_pivot = (
        pending_asp_book.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    pending_asp_2b_pivot = (
        pending_asp_2b.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )

    # Creating Pivot for extra both
    extra_asp_books_pivot = (
        extra_asp_book.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )
    extra_asp_2b_pivot = (
        extra_asp_2b.groupby(
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"], dropna=False
        )
        .agg({"BASIC": "sum", "CGST": "sum", "SGST": "sum", "IGST": "sum"})
        .reset_index()
    )

    # In[50]:

    # Merging Both Pivot Table 'Pending'
    pending_df = pd.merge(
        pending_asp_2b_pivot,
        pending_asp_books_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Renaming Columns Name of 'pending'
    pending_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # Merging Both Pivot Table 'Extra'
    extra_df = pd.merge(
        extra_asp_2b_pivot,
        extra_asp_books_pivot,
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Renaming Columns Name of 'extra_df'
    extra_df.rename(
        columns={
            "Gross Total_x": "2B Gross Total",
            "BASIC_x": "2B BASIC",
            "CGST_x": "2B CGST",
            "SGST_x": "2B SGST",
            "IGST_x": "2B IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[51]:

    # Calculating Difference
    pending_df.loc[:, "BASIC_diff"] = (
        pending_df.loc[:, "2B BASIC"] - pending_df.loc[:, "BOOKS BASIC"]
    )
    pending_df.loc[:, "CGST_diff"] = (
        pending_df.loc[:, "2B CGST"] - pending_df.loc[:, "BOOKS CGST"]
    )
    pending_df.loc[:, "SGST_diff"] = (
        pending_df.loc[:, "2B SGST"] - pending_df.loc[:, "BOOKS SGST"]
    )
    pending_df.loc[:, "IGST_diff"] = (
        pending_df.loc[:, "2B IGST"] - pending_df.loc[:, "BOOKS IGST"]
    )

    # Addition of All GST's Diff
    pending_df.loc[:, "GST_diff_total"] = (
        pending_df.loc[:, "CGST_diff"]
        + pending_df.loc[:, "SGST_diff"]
        + pending_df.loc[:, "IGST_diff"]
    )

    # Rearrancing Columns
    pending_df = pending_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Bill No",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[52]:

    # Calculating Difference
    extra_df.loc[:, "BASIC_diff"] = (
        extra_df.loc[:, "2B BASIC"] - extra_df.loc[:, "BOOKS BASIC"]
    )
    extra_df.loc[:, "CGST_diff"] = (
        extra_df.loc[:, "2B CGST"] - extra_df.loc[:, "BOOKS CGST"]
    )
    extra_df.loc[:, "SGST_diff"] = (
        extra_df.loc[:, "2B SGST"] - extra_df.loc[:, "BOOKS SGST"]
    )
    extra_df.loc[:, "IGST_diff"] = (
        extra_df.loc[:, "2B IGST"] - extra_df.loc[:, "BOOKS IGST"]
    )

    # Addition of All GST's Diff
    extra_df.loc[:, "GST_diff_total"] = (
        extra_df.loc[:, "CGST_diff"]
        + extra_df.loc[:, "SGST_diff"]
        + extra_df.loc[:, "IGST_diff"]
    )

    # Rearrancing Columns
    extra_df = extra_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Bill No",
            "Month",
            "Year",
            "2B BASIC",
            "BOOKS BASIC",
            "2B CGST",
            "BOOKS CGST",
            "2B SGST",
            "BOOKS SGST",
            "2B IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[53]:

    # Extra
    pending_df["STATUS"] = ""

    for i in range(pending_df.shape[0]):
        if (pending_df["GST_diff_total"][i] < 0) or (
            pending_df["GST_diff_total"][i] > 0
        ):
            pending_df.loc[i, "STATUS"] = "Review Required"
        elif pending_df["GST_diff_total"][i] == 0:
            pending_df.loc[i, "STATUS"] = "OK"
        else:
            pending_df.loc[i, "STATUS"] = "Pending Claim"

    # In[54]:

    # Extra
    extra_df["STATUS"] = ""

    for i in range(extra_df.shape[0]):
        if (extra_df["GST_diff_total"][i] < 0) or (extra_df["GST_diff_total"][i] > 0):
            extra_df.loc[i, "STATUS"] = "Review Required"
        elif extra_df["GST_diff_total"][i] == 0:
            extra_df.loc[i, "STATUS"] = "OK"
        else:
            extra_df.loc[i, "STATUS"] = "Extra Claimed"

    # In[ ]:

    # In[55]:

    # Joining Data for STATUS column
    pending_data = pd.merge(
        pending_data,
        pending_df[
            ["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year", "STATUS"]
        ],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Dropping STATUS_x column
    pending_data.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    pending_data.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # Separating Review/Pending rows
    pending_review = pending_data[pending_data["STATUS"] == "Review Required"]
    pending_ok = pending_data[pending_data["STATUS"] == "OK"]
    pending_data = pending_data[pending_data["STATUS"] == "Pending Claim"]

    # In[56]:

    # Joining Data for STATUS column
    extra_data = pd.merge(
        extra_data,
        extra_df[["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year", "STATUS"]],
        how="outer",
        on=["SUPPLIER NAME", "GSTIN/UIN", "Bill No", "Month", "Year"],
    )

    # Dropping STATUS_x column
    extra_data.drop(columns={"STATUS_x"}, inplace=True)

    # Renaming 'STATUS_y' column to 'STATUS'
    extra_data.rename(columns={"STATUS_y": "STATUS"}, inplace=True)

    # Separating Review/Extra rows
    extra_review = extra_data[extra_data["STATUS"] == "Review Required"]
    extra_ok = extra_data[extra_data["STATUS"] == "OK"]
    extra_data = extra_data[extra_data["STATUS"] == "Extra Claimed"]

    # In[57]:

    # All Review rows
    review_data = pd.concat([pending_review, extra_review], ignore_index=True)

    # OK status rows
    ok_data = main_data_df[main_data_df["STATUS"] == "OK"]

    # In[58]:

    # Creating main_data_df
    main_data_df = pd.concat(
        [pending_data, extra_data, ok_data, review_data, extra_ok, pending_ok],
        ignore_index=True,
    )

    # In[ ]:

    # In[ ]:

    # In[ ]:

    # In[ ]:

    # In[59]:

    # Dropping useless columns for export
    # merged_df.drop(columns=['Bill No','Month','Year'],inplace=True)
    main_data_df.drop(columns=["Bill No", "Month", "Year"], inplace=True)

    pending_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)
    extra_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)
    # ok_data.drop(columns=['Bill No','Month','Year'],inplace=True)
    review_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)

    # In[60]:

    # Empty GST Number
    empty_gdt_data = main_data_df[main_data_df["GSTIN/UIN"].isnull()].reset_index(
        drop=True
    )

    # In[61]:

    # Getting state name using GST number
    states = {
        "03": "Punjab",
        "01": "Jammu",
        "17": "Meghalaya",
        "27": "Pune",
        "04": "Chandigarh",
        "02": "Himachal Pardesh",
        "09": "Uttar Pardesh",
        "23": "Madhya Pardesh",
        "08": "Rajasthan",
        "06": "Haryana",
        "21": "Odisha",
    }

    state_name = states[file_gst_no[:2]]

    # In[62]:

    # Exporting Data To Another sheet

    file_name = f"ITC Reco. 2B VS {file_gst_no} {state_name} {current_month_name}.xlsx"

    with pd.ExcelWriter(file_name, mode="w") as writer:
        summary_df.to_excel(writer, sheet_name="SUMMARY", index=False)
        # merged_df.to_excel(writer, sheet_name='SUMMARY', index=False)

    with pd.ExcelWriter(file_name, mode="a", if_sheet_exists="replace") as writer:
        main_data_df.to_excel(writer, sheet_name="ITC 2B VS BOOKS", index=False)
        pending_data.to_excel(writer, sheet_name="PENDING", index=False)
        extra_data.to_excel(writer, sheet_name="EXTRA", index=False)
        review_data.to_excel(writer, sheet_name="REVIEW", index=False)
        empty_gdt_data.to_excel(writer, sheet_name="MISSING GSTIN", index=False)

    # ending time
    end_time = time.time()

    total_time = end_time - start_time

    print()
    print("TASK SUCCESSFULLY DONE IN:", total_time, "Seconds")
    print()

    end_or_run = input("Want to RUN this program again?: (Y/N): ")

    if end_or_run.lower() == "y":
        exepy()
    else:
        quit()


exepy()
