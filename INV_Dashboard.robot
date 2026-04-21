*** Settings ***
Library          INV_Dashboard.py
Suite Setup      Initialize Dashboard    ${DASHBOARD_FILE}
Suite Teardown   Finalize Dashboard

*** Variables ***
${DASHBOARD_FILE}    C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/Inventory Stock vs Utilization (Dashboard) V5.xlsb
${ISSUANCE_FILE}     C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/test data/CMPak_INV_Issuances_Report_160326.xls
${AGING_FILE}        C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/test data/Aging_Report_Inventory__CMPak_160326.xls
${RNR_FILE}          C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/test data/CMPAK_RNR_SLA_Report_200126.xls
${RECEIPT_FILE}      C:/Users/DELL XPS/OneDrive/Desktop/Inventory Mangement project/dashbaord & test data/dashboards/test data/CMPAK_Inventory_Receiving_Repo_200126.xls

*** Tasks ***
Update Issuance Sheet
    Update Issuance    ${ISSUANCE_FILE}

Update Inventory Sheet
    Update Inventory    ${AGING_FILE}

Update RnR Sheet
    Update Rnr    ${RNR_FILE}

Update Receipt Sheet
    Update Receipt    ${RECEIPT_FILE}
