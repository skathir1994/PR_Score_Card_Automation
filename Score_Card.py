# Importing the necessary packages
import openpyxl as open1
import pandas as pd

# Reading Rejection Soure file
source = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Rejection source List.xlsx"
df_source = pd.read_excel(source)
# Reading Rejection market place file
name = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Market place name.xlsx"
df_name = pd.read_excel(name)
# Reading merchants id file
merchants = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Overall merchants.xlsx"
df_merchants = pd.read_excel(merchants)
# Reading Clasification update file
Classification = (
    r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Clasification_updated.xlsx"
)
df_Classification = pd.read_excel(Classification)
# Reading Controllable file
Controllable = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Control.xlsx"
df_Controllable = pd.read_excel(Controllable)
# Reading Include/Exclude file
Exclude = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Include_Exclude.xlsx"
df_Exclude = pd.read_excel(Exclude)
# Reading Reason_For_Exclusion file
Reason = r"C:\Users\skathir\Desktop\PR File\Score Card Handy\Reason_For_Exclusion.xlsx"
df_Reason = pd.read_excel(Reason)

# Reading input base file
base = r"C:\Users\skathir\Desktop\PR File\2024\Input_Scorecard\Score Card Base_Week_22_IN.xlsx"
df_base1 = pd.read_excel(base)
#
# Splitting date from Date time
#
df_base1["Audit date"] = pd.to_datetime(df_base1["Audit date"]).dt.date
df_base1["Week"] = pd.to_datetime(df_base1["Audit date"]).dt.isocalendar().week
df_base1["Month"] = pd.to_datetime(df_base1["Audit date"]).dt.month_name()
df_base1["Quarter"] = pd.to_datetime(df_base1["Audit date"]).dt.to_period("Q")
#
# Creating the mkpl_id column using the merge function
#
df_base = df_base1.merge(df_name, on="marketplaceId")
#
# Creating new DF & Adding new Rejection_Source column using the merge function
#
df_Rejection_Source = pd.DataFrame(df_base)
df_Rejection_Source = pd.merge(df_Rejection_Source, df_source, on="rejectedBy")
df_base = df_base1.merge(df_name, on="marketplaceId")
df_base = pd.merge(df_base1, df_name, on="marketplaceId")
df_rejection1 = df_base1.merge(df_source, on="rejectedBy")
df_Rejection = pd.merge(df_base1, df_source, on="rejectedBy")
#
# Creating new DF & Adding new classification column using the merge function
#
df_classification = pd.DataFrame(df_Rejection_Source)
df_classification = pd.merge(
    df_Rejection_Source, df_Classification, on="Re-Classification Post RCA"
)
#
# Creating new DF & Adding new controllable column using the merge function
#
df_controllable = pd.DataFrame(df_classification)
df_controllable = pd.merge(
    df_classification, df_Controllable, on="Broad Classification"
)
#
# Creating new DF & Adding new Conn column using the concatenate method
#
df_Conc = pd.DataFrame(df_controllable)
df_Conc["Conc"] = (
    df_controllable["asin"]
    + df_controllable["competitorName"]
    + df_controllable["rejectReason"]
    + df_controllable["recommendedPrice"].astype(str)
    + df_controllable["URL"]
)
#
# Finding Unique/Duplicate values using Boolean method
#
df_Unique = pd.DataFrame(df_Conc)
df_Unique["Unique/Duplicate"] = df_Conc.duplicated("Conc").replace(
    {False: "Unique", True: "Duplicate"}
)
#
# Splitting the rejected date from the rejected Datetime.
#
df_Unique["rejectDate"] = pd.to_datetime(df_Unique["rejectTime"]).dt.date
#
# Finding the DPMU inscope merchants.
#
df_merchants1 = pd.DataFrame(df_Unique)
df_merchants1 = pd.merge(df_Unique, df_merchants, on="merchantId")
df_Conc1 = pd.DataFrame(df_merchants1)
df_Conc1["IN/EX"] = (
    df_merchants1["Unique/Duplicate"] + df_merchants1["Re-Classification Post RCA"]
)
df_InEx = pd.DataFrame(df_Conc1)
df_InEx = pd.merge(df_Conc1, df_Exclude, on="IN/EX")
#
df_Conc2 = pd.DataFrame(df_InEx)
df_Conc2["Reason"] = (
    df_InEx["Unique/Duplicate"]
    + df_InEx["Re-Classification Post RCA"]
    + df_InEx["DPMU Included/Excluded"]
)
df_Reason_for_ex = pd.DataFrame(df_Conc2)
#
df_Reason_for_ex = pd.merge(df_Conc2, df_Reason, on="Reason")
#
# Creating  final bussiness format.
#
df = pd.DataFrame(
    data=df_Reason_for_ex,
    columns=[
        "Week",
        "Month",
        "Quarter",
        "Audit date",
        "Marketplace Name",
        "marketplaceId",
        "merchantId",
        "asin",
        "competitorName",
        "rejectReason",
        "recommendedPrice",
        "rejectTime",
        "rejectedBy",
        "glName",
        "categoryCode",
        "URL",
        "competitorPrice",
        "AUDIT_STATUS ",
        "ERROR_TYPE",
        "ISSUE_GROUP",
        "ISSUE_DESCRIPTION",
        "mapper id",
        "mapped date",
        "Comments",
        "Auditor ",
        "Rejection Source (PA/DSM/PDM/Andon)",
        "Re-Classification Post RCA",
        "Broad Classification",
        "Controllable/Uncontrollable",
        "DPMU Included/Excluded",
        "Reason For Exclusion",
        "Conc",
        "Unique/Duplicate",
        "DPMU metric inclusive merchants",
        "rejectDate",
    ],
)
#
# Export to Excel format in the desired location.
#
df.to_excel(
    "C:\\Users\\skathir\\Desktop\\PR File\\2024\\Input_Scorecard\\Output\Score Card Base_Week_22.xlsx",
    index=False,
)
