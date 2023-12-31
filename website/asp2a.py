from flask import Flask, render_template, request, send_file, make_response
import pandas as pd
import os
import xlsxwriter


def asp2a(app):
    pur_path = request.files["purchase"]
    b2b_path = request.files["b2b"]
    lastMonth = request.files.get("lastMonthFile")

    # loading purchase books file
    purchase = pd.read_excel(pur_path, skiprows=3, engine="openpyxl")

    # loading B2B file
    b2b = pd.read_excel(b2b_path, sheet_name="B2B", skiprows=4, engine="openpyxl")
    b2b.drop(0, inplace=True)

    # Loading NOTE sheet
    b2b_cdnr = pd.read_excel(b2b_path, sheet_name="CDNR", skiprows=4, engine="openpyxl")
    b2b_cdnr.drop(0, inplace=True)

    # getting current month of file
    current_monn_file = pd.read_excel(b2b_path, sheet_name="Read me")["Unnamed: 4"][0][
        0:2
    ]
    current_monn_file = pd.to_datetime(current_monn_file, format="%m").month_name()
    if current_monn_file not in ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November','December']:
        current_monn_file = 'March'
    current_month_name = current_monn_file

    # getting file gst number
    file_gst_no = pd.read_excel(b2b_path, sheet_name="Read me")["Unnamed: 2"][0]

    if request.form.get("lastMonth") and lastMonth:
        # Loading LAST MONTH PENDING and EXTRA sheet
        pending_last_month = pd.read_excel(
            lastMonth, sheet_name="PENDING", engine="openpyxl"
        )
        extra_last_month = pd.read_excel(
            lastMonth, sheet_name="EXTRA", engine="openpyxl"
        )
        pending_last_month["Data From"] = "Last Month Pending Claim"
        extra_last_month["Data From"] = "Last Month Extra Claim"
        pending_last_month.rename(columns={"Voucher Date": "Date"}, inplace=True)
        extra_last_month.rename(columns={"Voucher Date": "Date"}, inplace=True)
        # Renaming LAST MONTH DATA SUPPLIER NAME
        pending_last_month.rename(columns={"SUPPLIER NAME": "Trade Name"}, inplace=True)
        extra_last_month.rename(columns={"SUPPLIER NAME": "Trade Name"}, inplace=True)

    b2b = b2b[~b2b["Invoice details"].astype(str).str.contains("Total")]
    b2b = b2b[~b2b["GSTIN of supplier"].isna()].reset_index(drop=True)

    b2b_cdnr = b2b_cdnr[~b2b_cdnr["Unnamed: 3"].astype(str).str.contains("Total")]
    b2b_cdnr = b2b_cdnr[~b2b_cdnr["GSTIN of Supplier"].isna()].reset_index(drop=True)

    # Renaming columns of B2B and B2B CDNR
    b2b.rename(
        columns={
            "Invoice details": "Invoice number",
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
            "Trade/Legal name of the supplier": "Trade/Legal name of the Supplier",
            "GSTIN of Supplier": "GSTIN of supplier",
            "Credit note/Debit note details": "Invoice number",
            "Credit note/Debit note details   \n": "Invoice type",
            "Unnamed: 3": "Invoice number",
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

    # Dropping 'Regular' column from b2b_cdnr
    b2b_cdnr.drop(columns={"Note Supply type"}, inplace=True)

    # converting to negative values of 'Credit Note'
    for i in range(b2b_cdnr.shape[0]):
        if b2b_cdnr.loc[i, "Invoice type"].lower() == "credit note":
            b2b_cdnr.loc[i, "Invoice Value(₹)"] = 0 - b2b_cdnr["Invoice Value(₹)"][i]
            b2b_cdnr.loc[i, "Taxable Value (₹)"] = 0 - b2b_cdnr["Taxable Value (₹)"][i]
            b2b_cdnr.loc[i, "Integrated Tax(₹)"] = 0 - b2b_cdnr["Integrated Tax(₹)"][i]
            b2b_cdnr.loc[i, "Central Tax(₹)"] = 0 - b2b_cdnr["Central Tax(₹)"][i]
            b2b_cdnr.loc[i, "State/UT Tax(₹)"] = 0 - b2b_cdnr["State/UT Tax(₹)"][i]
            b2b_cdnr.loc[i, "Cess(₹)"] = 0 - b2b_cdnr["Cess(₹)"][i]

    # Concatinating b2b_cdnr data and B2B to B2B
    b2b = pd.concat([b2b, b2b_cdnr], ignore_index=True)

    # Taking only those rows with 'No' values
    b2b = b2b[(b2b['Supply Attract Reverse Charge'].str.lower()=='n') | (b2b['Supply Attract Reverse Charge'].str.lower()=='no')]

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
    purchase["Cess(₹)"] = 0
    b2b = b2b.iloc[:, :14]

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

    # Renaming columns
    purchase.rename(columns={"Particulars": "Trade Name"}, inplace=True)

    # Adding Columns
    REMARKS = ""
    STATUS = ""
    data_from = ""
    purchase.insert(2, "REMARKS", REMARKS)
    purchase.insert(12, "STATUS", STATUS)
    purchase.insert(13, "Data From", data_from)

    # Filling REMARKS column
    purchase["REMARKS"] = "AS PER BOOKS"

    # Creating Dummy Dataframe for concatinating
    b2b_df = pd.DataFrame()

    b2b_df["Date"] = b2b["Invoice Date"]
    b2b_df["Trade Name"] = b2b["Trade/Legal name of the Supplier"]
    b2b_df["REMARKS"] = "AS PER 2A"
    b2b_df["Voucher Type"] = b2b["Invoice type"]
    b2b_df["Supplier Invoice No."] = b2b["Invoice number"]
    b2b_df["Supplier Invoice Date"] = b2b["Invoice Date"]
    b2b_df["GSTIN/UIN"] = b2b["GSTIN of supplier"]
    b2b_df["Gross Total"] = ""
    b2b_df["BASIC"] = b2b["Taxable Value (₹)"]
    b2b_df["CGST"] = b2b["Central Tax(₹)"]
    b2b_df["SGST"] = b2b["State/UT Tax(₹)"]
    b2b_df["IGST"] = b2b["Integrated Tax(₹)"]
    b2b_df["Cess(₹)"] = b2b["Cess(₹)"]
    b2b_df["STATUS"] = ""

    if request.form.get("lastMonth") and lastMonth:
        main_data = pd.concat(
            [purchase, b2b_df, pending_last_month, extra_last_month], ignore_index=True
        )
    else:
        main_data = pd.concat([purchase, b2b_df], ignore_index=True)

    # Setting GST NO column
    main_data["GSTIN/UIN"] = main_data["GSTIN/UIN"].str.replace("_x000D_\n", "")
    main_data["GSTIN/UIN"] = main_data["GSTIN/UIN"].str.replace("\n", "")

    # Calculating GROSS column
    main_data["Gross Total"] = (
        main_data["BASIC"] + main_data["CGST"] + main_data["SGST"] + main_data["IGST"]
    )

    # Creating SUPPLIER NAME column
    supplier_name = ""
    main_data.insert(2, "SUPPLIER NAME", supplier_name)

    # Separating 'AS PER BOOKS' and 'AS PER 2B' from main_data
    main_asp_books = main_data[main_data["REMARKS"] == "AS PER BOOKS"]
    main_asp_2b = main_data[main_data["REMARKS"] == "AS PER 2B"]

    # Grouping on the basis of GSTIN
    books_gstin_join = (
        main_asp_books.groupby("GSTIN/UIN", dropna=False)
        .agg({"Trade Name": "first"})
        .reset_index()
    )

    # Performing join on both above data
    main_asp_2b = pd.merge(main_asp_2b, books_gstin_join, how="left", on="GSTIN/UIN")
    main_asp_2b["Trade Name_y"].fillna(main_asp_2b["Trade Name_x"], inplace=True)

    # Transfering column [Supplier Name]
    main_asp_2b.loc[:, "SUPPLIER NAME"] = main_asp_2b["Trade Name_y"]
    main_asp_books.loc[:, "SUPPLIER NAME"] = main_asp_books["Trade Name"]

    main_data = pd.concat([main_asp_books, main_asp_2b], ignore_index=True)
    main_data.drop(columns=["Trade Name", "Trade Name_x", "Trade Name_y"], inplace=True)

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

    # Extracting last 2 character of bill Number
    main_data["Bill No"] = main_data["Supplier Invoice No."].apply(
        lambda x: str(x)[-3:]
    )

    # Getting Month and Year of each row
    main_data["Month"] = main_data["Supplier Invoice Date"].dt.month
    main_data["Year"] = main_data["Supplier Invoice Date"].dt.year

    # Creating Pivot Table

    # Separating 'AS PER BOOKS' and 'AS PER 2A'
    aspbooks = main_data[main_data["REMARKS"] == "AS PER BOOKS"].reset_index(drop=True)
    asp2b = main_data[main_data["REMARKS"] == "AS PER 2A"].reset_index(drop=True)

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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )
    # In[26]:

    # Filling 'NaN' values with '0'
    merged_df["2A BASIC"] = merged_df["2A BASIC"].fillna(0)
    merged_df["2A CGST"] = merged_df["2A CGST"].fillna(0)
    merged_df["2A SGST"] = merged_df["2A SGST"].fillna(0)
    merged_df["2A IGST"] = merged_df["2A IGST"].fillna(0)
    merged_df["BOOKS BASIC"] = merged_df["BOOKS BASIC"].fillna(0)
    merged_df["BOOKS CGST"] = merged_df["BOOKS CGST"].fillna(0)
    merged_df["BOOKS SGST"] = merged_df["BOOKS SGST"].fillna(0)
    merged_df["BOOKS IGST"] = merged_df["BOOKS IGST"].fillna(0)

    # In[27]:

    # Calculating Difference
    merged_df["BASIC_diff"] = merged_df["2A BASIC"] - merged_df["BOOKS BASIC"]
    merged_df["CGST_diff"] = merged_df["2A CGST"] - merged_df["BOOKS CGST"]
    merged_df["SGST_diff"] = merged_df["2A SGST"] - merged_df["BOOKS SGST"]
    merged_df["IGST_diff"] = merged_df["2A IGST"] - merged_df["BOOKS IGST"]

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
            "2A BASIC",
            "BOOKS BASIC",
            "2A CGST",
            "BOOKS CGST",
            "2A SGST",
            "BOOKS SGST",
            "2A IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[28]:

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

    # In[29]:

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

    # In[30]:

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

    # In[31]:

    # Conterting datetime to date only for export
    main_data_df["Date"] = pd.to_datetime(main_data_df["Date"])
    main_data_df["Supplier Invoice Date"] = pd.to_datetime(
        main_data_df["Supplier Invoice Date"]
    )

    # Renaming Date Column
    main_data_df.rename(columns={"Date": "Voucher Date"}, inplace=True)

    # In[32]:

    # Separating pending and extra rows
    pending_data = main_data_df[main_data_df["STATUS"] == "Pending Claim"].reset_index(
        drop=True
    )
    extra_data = main_data_df[main_data_df["STATUS"] == "Extra Claimed"].reset_index(
        drop=True
    )

    # In[33]:

    # Creating Separate Data for both 'Pending' and 'Extra' rows
    pending_asp_book = pending_data[pending_data["REMARKS"] == "AS PER BOOKS"]
    pending_asp_2b = pending_data[pending_data["REMARKS"] == "AS PER 2A"]

    extra_asp_book = extra_data[extra_data["REMARKS"] == "AS PER BOOKS"]
    extra_asp_2b = extra_data[extra_data["REMARKS"] == "AS PER 2A"]

    # In[34]:

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

    # In[35]:

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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[36]:

    # Calculating Difference
    pending_df["BASIC_diff"] = pending_df["2A BASIC"] - pending_df["BOOKS BASIC"]
    pending_df["CGST_diff"] = pending_df["2A CGST"] - pending_df["BOOKS CGST"]
    pending_df["SGST_diff"] = pending_df["2A SGST"] - pending_df["BOOKS SGST"]
    pending_df["IGST_diff"] = pending_df["2A IGST"] - pending_df["BOOKS IGST"]

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
            "2A BASIC",
            "BOOKS BASIC",
            "2A CGST",
            "BOOKS CGST",
            "2A SGST",
            "BOOKS SGST",
            "2A IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[37]:

    # Calculating Difference
    extra_df["BASIC_diff"] = extra_df["2A BASIC"] - extra_df["BOOKS BASIC"]
    extra_df["CGST_diff"] = extra_df["2A CGST"] - extra_df["BOOKS CGST"]
    extra_df["SGST_diff"] = extra_df["2A SGST"] - extra_df["BOOKS SGST"]
    extra_df["IGST_diff"] = extra_df["2A IGST"] - extra_df["BOOKS IGST"]

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
            "2A BASIC",
            "BOOKS BASIC",
            "2A CGST",
            "BOOKS CGST",
            "2A SGST",
            "BOOKS SGST",
            "2A IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[38]:

    # Pending
    pending_df["STATUS"] = ""

    for i in range(pending_df.shape[0]):
        if (pending_df["GST_diff_total"][i] < 0) or (
            pending_df["GST_diff_total"][i] > 0
        ):
            pending_df.loc[i, "STATUS"] = "Review Required"
        else:
            pending_df.loc[i, "STATUS"] = "Pending Claim"

    # In[39]:

    # Extra
    extra_df["STATUS"] = ""

    for i in range(extra_df.shape[0]):
        if (extra_df["GST_diff_total"][i] < 0) or (extra_df["GST_diff_total"][i] > 0):
            extra_df.loc[i, "STATUS"] = "Review Required"
        else:
            extra_df.loc[i, "STATUS"] = "Extra Claimed"

    # In[40]:

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

    # In[41]:

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

    # In[42]:

    # All Review rows
    review_data = pd.concat([pending_review, extra_review], ignore_index=True)

    # OK status rows
    ok_data = main_data_df[main_data_df["STATUS"] == "OK"]

    # In[43]:

    # Creating main_data_df
    main_data_df = pd.concat(
        [pending_data, extra_data, ok_data, review_data], ignore_index=True
    )

    # In[44]:

    # Creating SUMMARY AGAIN on NAME and GSTN

    # Separating 'AS PER BOOKS' and 'AS PER 2A'
    aspbooks = main_data_df[main_data_df["REMARKS"] == "AS PER BOOKS"].reset_index(
        drop=True
    )
    asp2b = main_data_df[main_data_df["REMARKS"] == "AS PER 2A"].reset_index(drop=True)

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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[45]:

    # Filling 'NaN' values with '0'
    summary_df["2A BASIC"] = summary_df["2A BASIC"].fillna(0)
    summary_df["2A CGST"] = summary_df["2A CGST"].fillna(0)
    summary_df["2A SGST"] = summary_df["2A SGST"].fillna(0)
    summary_df["2A IGST"] = summary_df["2A IGST"].fillna(0)
    summary_df["BOOKS BASIC"] = summary_df["BOOKS BASIC"].fillna(0)
    summary_df["BOOKS CGST"] = summary_df["BOOKS CGST"].fillna(0)
    summary_df["BOOKS SGST"] = summary_df["BOOKS SGST"].fillna(0)
    summary_df["BOOKS IGST"] = summary_df["BOOKS IGST"].fillna(0)

    # Calculating Difference
    summary_df["2A_VS_Books_BASIC_diff"] = (
        summary_df["2A BASIC"] - summary_df["BOOKS BASIC"]
    )
    summary_df["2A_VS_Books_CGST_diff"] = (
        summary_df["2A CGST"] - summary_df["BOOKS CGST"]
    )
    summary_df["2A_VS_Books_SGST_diff"] = (
        summary_df["2A SGST"] - summary_df["BOOKS SGST"]
    )
    summary_df["2A_VS_Books_IGST_diff"] = (
        summary_df["2A IGST"] - summary_df["BOOKS IGST"]
    )

    # Addition of All GST's Diff
    summary_df["Total_GST_diff"] = (
        summary_df["2A_VS_Books_CGST_diff"]
        + summary_df["2A_VS_Books_SGST_diff"]
        + summary_df["2A_VS_Books_IGST_diff"]
    )

    # Rearrancing Columns
    summary_df = summary_df[
        [
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "Month",
            "Year",
            "2A BASIC",
            "2A CGST",
            "2A SGST",
            "2A IGST",
            "BOOKS BASIC",
            "BOOKS CGST",
            "BOOKS SGST",
            "BOOKS IGST",
            "2A_VS_Books_BASIC_diff",
            "2A_VS_Books_CGST_diff",
            "2A_VS_Books_SGST_diff",
            "2A_VS_Books_IGST_diff",
            "Total_GST_diff",
        ]
    ]

    # In[46]:

    # Creating STATUS column
    summary_df.loc[:, "STATUS"] = ""

    # Filling STATUS column with 'OK', 'PENDING' and 'EXTRA'
    for i in range(summary_df.shape[0]):
        if summary_df.loc[i, "Total_GST_diff"] > 9:
            summary_df.loc[i, "STATUS"] = "Extra Claimed"
        elif summary_df.loc[i, "Total_GST_diff"] < -9:
            summary_df.loc[i, "STATUS"] = "Pending Claim"
        else:
            summary_df.loc[i, "STATUS"] = "OK"

    for i in range(summary_df.shape[0]):
        if (
            (
                (summary_df.loc[i, "2A_VS_Books_CGST_diff"] < -1)
                or (summary_df.loc[i, "2A_VS_Books_CGST_diff"] > 1)
            )
            and (
                (summary_df.loc[i, "2A_VS_Books_SGST_diff"] < -1)
                or (summary_df.loc[i, "2A_VS_Books_SGST_diff"] > 1)
            )
            and (
                (summary_df.loc[i, "2A_VS_Books_IGST_diff"] < -1)
                or (summary_df.loc[i, "2A_VS_Books_IGST_diff"] > 1)
            )
        ):
            summary_df.loc[i, "STATUS"] = "Mismatch GST"

    # In[47]:

    # Adding total row in summary_df
    total_sum_summary = pd.DataFrame(summary_df.select_dtypes(include="number").sum()).T
    summary_df = pd.concat([summary_df, total_sum_summary])

    # In[48]:

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

    # In[49]:

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

    # In[50]:

    pending_data = main_data_df[main_data_df["STATUS"] == "Pending Claim"]
    extra_data = main_data_df[main_data_df["STATUS"] == "Extra Claimed"]

    # In[51]:

    pending_asp_book = pending_data[pending_data["REMARKS"] == "AS PER BOOKS"]
    pending_asp_2b = pending_data[pending_data["REMARKS"] == "AS PER 2A"]

    extra_asp_book = extra_data[extra_data["REMARKS"] == "AS PER BOOKS"]
    extra_asp_2b = extra_data[extra_data["REMARKS"] == "AS PER 2A"]

    # In[52]:

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

    # In[53]:

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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
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
            "Gross Total_x": "2A Gross Total",
            "BASIC_x": "2A BASIC",
            "CGST_x": "2A CGST",
            "SGST_x": "2A SGST",
            "IGST_x": "2A IGST",
            "Gross Total_y": "BOOKS Gross Total",
            "BASIC_y": "BOOKS BASIC",
            "CGST_y": "BOOKS CGST",
            "SGST_y": "BOOKS SGST",
            "IGST_y": "BOOKS IGST",
        },
        inplace=True,
    )

    # In[54]:

    # Calculating Difference
    pending_df.loc[:, "BASIC_diff"] = (
        pending_df.loc[:, "2A BASIC"] - pending_df.loc[:, "BOOKS BASIC"]
    )
    pending_df.loc[:, "CGST_diff"] = (
        pending_df.loc[:, "2A CGST"] - pending_df.loc[:, "BOOKS CGST"]
    )
    pending_df.loc[:, "SGST_diff"] = (
        pending_df.loc[:, "2A SGST"] - pending_df.loc[:, "BOOKS SGST"]
    )
    pending_df.loc[:, "IGST_diff"] = (
        pending_df.loc[:, "2A IGST"] - pending_df.loc[:, "BOOKS IGST"]
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
            "2A BASIC",
            "BOOKS BASIC",
            "2A CGST",
            "BOOKS CGST",
            "2A SGST",
            "BOOKS SGST",
            "2A IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[55]:

    # Calculating Difference
    extra_df.loc[:, "BASIC_diff"] = (
        extra_df.loc[:, "2A BASIC"] - extra_df.loc[:, "BOOKS BASIC"]
    )
    extra_df.loc[:, "CGST_diff"] = (
        extra_df.loc[:, "2A CGST"] - extra_df.loc[:, "BOOKS CGST"]
    )
    extra_df.loc[:, "SGST_diff"] = (
        extra_df.loc[:, "2A SGST"] - extra_df.loc[:, "BOOKS SGST"]
    )
    extra_df.loc[:, "IGST_diff"] = (
        extra_df.loc[:, "2A IGST"] - extra_df.loc[:, "BOOKS IGST"]
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
            "2A BASIC",
            "BOOKS BASIC",
            "2A CGST",
            "BOOKS CGST",
            "2A SGST",
            "BOOKS SGST",
            "2A IGST",
            "BOOKS IGST",
            "BASIC_diff",
            "CGST_diff",
            "SGST_diff",
            "IGST_diff",
            "GST_diff_total",
        ]
    ]

    # In[56]:

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

    # In[57]:

    # Extra
    extra_df["STATUS"] = ""

    for i in range(extra_df.shape[0]):
        if (extra_df["GST_diff_total"][i] < 0) or (extra_df["GST_diff_total"][i] > 0):
            extra_df.loc[i, "STATUS"] = "Review Required"
        elif extra_df["GST_diff_total"][i] == 0:
            extra_df.loc[i, "STATUS"] = "OK"
        else:
            extra_df.loc[i, "STATUS"] = "Extra Claimed"

    # In[58]:

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

    # In[59]:

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

    # In[60]:

    # All Review rows
    review_data = pd.concat([pending_review, extra_review], ignore_index=True)

    # OK status rows
    ok_data = main_data_df[main_data_df["STATUS"] == "OK"]

    # Mismatch GST status rows
    mismatch_gst_data = main_data_df[main_data_df["STATUS"] == "Mismatch GST"]

    # In[61]:

    # Creating main_data_df
    main_data_df = pd.concat(
        [
            pending_data,
            extra_data,
            ok_data,
            review_data,
            extra_ok,
            pending_ok,
            mismatch_gst_data,
        ],
        ignore_index=True,
    )

    # In[62]:

    # Dropping useless columns for export
    # merged_df.drop(columns=['Bill No','Month','Year'],inplace=True)
    main_data_df.drop(columns=["Bill No", "Month", "Year"], inplace=True)

    pending_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)
    extra_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)
    # ok_data.drop(columns=['Bill No','Month','Year'],inplace=True)
    review_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)
    mismatch_gst_data.drop(columns=["Bill No", "Month", "Year"], inplace=True)

    # In[63]:

    # Empty GST Number
    empty_gdt_data = main_data_df[main_data_df["GSTIN/UIN"].isnull()].reset_index(
        drop=True
    )

    # In[67]:

    # Adding 'DATA FROM' column to main data
    main_data_df["DATA FROM"] = main_data_df["Data From"]
    main_data_df.drop(columns=["Data From"], inplace=True)

    # Adding 'DATA FROM' column to pending data
    pending_data["DATA FROM"] = pending_data["Data From"]
    pending_data.drop(columns=["Data From"], inplace=True)

    # Adding 'DATA FROM' column to extra data
    extra_data["DATA FROM"] = extra_data["Data From"]
    extra_data.drop(columns=["Data From"], inplace=True)

    # Adding 'DATA FROM' column to review data
    review_data["DATA FROM"] = review_data["Data From"]
    review_data.drop(columns=["Data From"], inplace=True)

    # Adding 'DATA FROM' column to empty_gdt_data data
    empty_gdt_data["DATA FROM"] = empty_gdt_data["Data From"]
    empty_gdt_data.drop(columns=["Data From"], inplace=True)

    # Adding 'DATA FROM' column to mismatch_gst_data data
    mismatch_gst_data["DATA FROM"] = mismatch_gst_data["Data From"]
    mismatch_gst_data.drop(columns=["Data From"], inplace=True)

    # In[69]:

    # Rearraging all data
    main_data_df = main_data_df[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]
    pending_data = pending_data[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]
    extra_data = extra_data[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]
    review_data = review_data[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]
    mismatch_gst_data = mismatch_gst_data[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]
    empty_gdt_data = empty_gdt_data[
        [
            "Voucher Date",
            "SUPPLIER NAME",
            "GSTIN/UIN",
            "REMARKS",
            "Voucher Type",
            "Supplier Invoice No.",
            "Supplier Invoice Date",
            "Gross Total",
            "BASIC",
            "CGST",
            "SGST",
            "IGST",
            "Cess(₹)",
            "LAST CLAIM MONTH",
            "STATUS",
            "DATA FROM",
        ]
    ]

    # Applying colors
    def highlight_rows(row):
        if (row["CGST"] < 0) or (row["SGST"] < 0) or (row["IGST"] < 0):
            return ["color: red"] * len(row)
        else:
            return [""] * len(row)  # empty style for other rows

    main_data_df = main_data_df.style.apply(highlight_rows, axis=1)
    pending_data = pending_data.style.apply(highlight_rows, axis=1)
    extra_data = extra_data.style.apply(highlight_rows, axis=1)
    review_data = review_data.style.apply(highlight_rows, axis=1)
    mismatch_gst_data = mismatch_gst_data.style.apply(highlight_rows, axis=1)
    empty_gdt_data = empty_gdt_data.style.apply(highlight_rows, axis=1)

    # In[70]:

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

    # In[71]:

    # Exporting Data To Another sheet

    file_name = f"ITC Reco. 2A VS {file_gst_no} {state_name} {current_month_name}.xlsx"

    # Create the output Excel file
    output_dir = os.path.join(app.root_path, "output")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, file_name)
    with pd.ExcelWriter(
        output_path, mode="w", engine="xlsxwriter", datetime_format="dd-mm-yyyy"
    ) as writer:
        summary_df.to_excel(writer, sheet_name="SUMMARY", index=False)
        # merged_df.to_excel(writer, sheet_name='SUMMARY', index=False)
        main_data_df.to_excel(writer, sheet_name="ITC 2A VS BOOKS", index=False)
        pending_data.to_excel(writer, sheet_name="PENDING", index=False)
        extra_data.to_excel(writer, sheet_name="EXTRA", index=False)
        review_data.to_excel(writer, sheet_name="REVIEW", index=False)
        mismatch_gst_data.to_excel(writer, sheet_name="MISMATCH GST", index=False)
        empty_gdt_data.to_excel(writer, sheet_name="MISSING GSTIN", index=False)

    # Set the response headers
    response = make_response(
        send_file(
            output_path,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    )

    # Set the filename for the downloaded file
    response.headers[
        "Content-Disposition"
    ] = "attachment; filename=" + os.path.basename(output_path)

    return response
