import pandas as pd
import numpy as np
import time
import sys
import ctypes
from asp2b import asp2b
from asp2a import asp2a
from PurchaseConverter import purchaseConverter

# /********************************* MAIN *********************************/

print("****************************************")
print("Developed by: - Naman Garg")
print("Contact: - namangarg2075@gmail.com")
print("****************************************")


def app():
    print()
    convert = input("Want to create new *BOOK PURCHASE* File? (Y/N): ")

    if convert.lower() == "y":
        purchaseConverter()

    asp_2b_or_2a = input("ITC Reconsiliation With 2A or 2B (2A/2B): ")

    if asp_2b_or_2a.lower() == "2a":
        asp2a()
    elif asp_2b_or_2a.lower() == "2b":
        asp2b()

    end_or_run = input("Want to RUN this program again?: (Y/N): ")

    if end_or_run.lower() == "y":
        app()
    else:
        print("THANKS FOR USING ME :)")
        sys.exit(0)


app()
