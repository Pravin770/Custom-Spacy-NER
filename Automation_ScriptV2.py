import pyodbc
import schedule
import time
import subprocess
import datetime

#************************Check for Vessel**************************************************************************************************/
# Define a function to check for new records in the table
def check_for_new_records():
    # Connect to the database
    conn_str = pyodbc.connect(
        r'Driver=SQL Server;'
        r'Server=xxxxx;'
        r'Database=xxxxx;'
        r'Trusted_Connection=yes;'
        )
    cursor = conn_str.cursor()

    # Execute a query to check for new records
    vess_query = "SELECT COUNT(*) FROM [xxxx].[dbo].[tblFeedbackQuoteVessel] WHERE fldCreatedOn > DATEADD(minute, -10, GETDATE())"
    
    cursor.execute(vess_query)
    # Get the count of new records
    vess_count = cursor.fetchone()[0]

    # Execute a query to check for new records
    cust_query = "SELECT COUNT(*) FROM [xxxx].[dbo].[tblFeedbackQuoteCustomer] WHERE fldCreatedOn > DATEADD(minute, -10, GETDATE())"
    
    cursor.execute(cust_query)
    # Get the count of new records
    cust_count = cursor.fetchone()[0]

    # Close the connection
    conn_str.close()

    # If there are new records, run the script
    if vess_count > 0:
        run_script3()
    
    if cust_count > 0:
        run_script4()


def run_script1():
    subprocess.run(["python", r"C:\Users\pravin.subramanian\Downloads\Python_Scripts\email_enquiries_one.py"])   #Automate_Notebook email_enquiries_one
    print('Script1 time: ', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
def run_script2():
    subprocess.run(["python", r"C:\Users\pravin.subramanian\Downloads\Python_Scripts\AUK_Email_Enquiries_test.py"])   #Automate_Notebook email_enquiries_one
    print('Script2 time: ', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
def run_script3():
    print('run_script3')
    # subprocess.run(["python", r"C:\Users\pravin.subramanian\Downloads\Python_Scripts\AUK_Email_Enquiries_test.py"])   #Automate_Notebook email_enquiries_one
    print('Script3 time: ', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
def run_script4():
    print('run_script4')
    # subprocess.run(["python", r"C:\Users\pravin.subramanian\Downloads\Python_Scripts\AUK_Email_Enquiries_test.py"])   #Automate_Notebook email_enquiries_one
    print('Script4 time: ', datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

# schedule.every(10).minutes.do(run_script1)
# schedule.every(10).minutes.do(run_script2)
schedule.every(1).minutes.do(check_for_new_records)

while True:
    schedule.run_pending()
    time.sleep(1)
    
