import cx_Oracle
import datetime as dt
import pandas as pd
connection='c##ankur/sql2324@localhost:1521/orcl'
def get_user_input():
    st_name = input("Enter student name: ")
    dob_str = input("Enter date of birth (YYYY-MM-DD): ")
    dob = dt.datetime.strptime(dob_str, '%Y-%m-%d')
    studentid = int(input("Enter student ID: "))
    return st_name, dob, studentid

conn = None
try:
    conn = cx_Oracle.connect(connection)
    cursor = conn.cursor()
    """dataInsertionTuples = [
        ('pen', dt.datetime(2009, 9, 12), 8415),
        ('span', dt.datetime(2012, 10, 17), 6548)
    ]"""
    dataInsertionTuples = []
    while True:
        user_input = input("Do you want to enter student data? (yes/no): ").lower()
        if user_input == 'no':
            break
        elif user_input == 'yes':
            dataInsertionTuples.append(get_user_input())
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")

    sqlTxt = 'INSERT INTO students\
                (st_name, dob, studentid)\
                VALUES (:1, :2, :3)'
    #cursor.executemany(sqlTxt, [x for x in dataInsertionTuples])
    cursor.executemany(sqlTxt, dataInsertionTuples)

    rowCount = cursor.rowcount
   

    cursor.execute('SELECT * FROM students')
    result = cursor.fetchall()

    columns = ['id','st_name', 'dob', 'studentid']
    df = pd.DataFrame(result, columns=columns)
    excel_filename = 'student_data.xlsx'


    with pd.ExcelWriter(excel_filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
        worksheet.set_column('C:C', None, date_format)
    
    print("number of inserted rows =", rowCount)
    print(f"Data saved to {excel_filename}")

    conn.commit()

except Exception as err:
    print('Error while connecting to the db')
    print(err)

finally:
    if(conn):
        cursor.close()
        conn.close()
print("execution complete!")


