"""
This program wil obtain retrieve and archive volume in MB and GB for each ASP customer.
"""
import pyodbc
import datetime
import smtplib
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.styles import PatternFill
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders


def archive_retrieve_report():

    """
    This function will do 4 tasks:
    1.  Connect to SQL Server.
    2.  Run query and retrieve data from SQL server Data Base.
    3.  Write the data to Excel worksheet.
    4.  Email the worksheet to end user.
    """

    # File name for Excel Worksheet
    excel_filename = ""

    # Get all Virtual Archive names from SQL Server
    virtual_archives = get_archives()

    # Set up Excel Worksheet.
    work_book = Workbook()
    work_sheet = work_book.active
    work_sheet.title = "EA Monthly Report"
    work_sheet.append(['Archive Name', 'Average Retrieve Volume'])

    # Assign font and background color properties for Column Title cells
    f = Font(name="Arial", size=14, bold=True, color="FF000000")
    fill = PatternFill(fill_type="solid", start_color="00FFFF00")

    work_sheet["A1"].fill = fill
    work_sheet["B1"].fill = fill

    work_sheet["A1"].font = f
    work_sheet["B1"].font = f

    # Set Column width
    work_sheet.column_dimensions["A"].width = 25.0
    work_sheet.column_dimensions["B"].width = 45.0

    # Obtain Exam Volume for all virtual archives and write the data to excel sheet.
    for archive in virtual_archives:
        # Obtains archive volume from SQL server.
        rows = retrieve_volume(archive)

        total_retrieves = 0
        avg_retrieves = 0
        for row in rows:
            total_retrieves += row[2]

        # Calculate average retrieves
        avg_retrieves = total_retrieves / 30

        # Adds archive name and exam volume to Worksheet.
        work_sheet.append([archive,  # Archive Name
                           round(avg_retrieves, 2),  # Average Retrieves
                           ])

        print("Added {} Average Retrieve Volume to Workbook!".format(archive))

    # Saves Excel worksheet.
    excel_filename = "ASP_Average_Retrieves_{}.xlsx".format(datetime.datetime.now().strftime("%Y-%m-%d"))
    work_book.save(excel_filename)

    # # Send email with attachment.
    send_email(excel_filename)


# Obtain Archive Names from SQL Server.
def get_archives():
    """
    This function will return list of Virtual Archive Names from SQL Server.
    """
    # Create list to store archive names.
    archive_list = []

    # Define data base connection parameters.
    sqlserver = 'SQL1'
    database = 'RSAdmin'
    username = 'admin'
    password = 'Bosnia66s'

    # Establish DB connections.
    conn = pyodbc.connect(
        r'DRIVER={SQL Server};'
        r'SERVER=' + sqlserver + ';'
        r'DATABASE=' + database + ';'
        r'UID=' + username + ';'
        r'PWD=' + password + ''
    )
    cur = conn.cursor()
    # Execute query on Data Base.
    cur.execute("""
                SELECT DBName from tblArchive
                ORDER BY DBName
                """)
    rows = cur.fetchall()

    # Add Archive names to archive list.
    for row in rows:
        archive_list.append(row[0])
    # Close SQL Connection.
    cur.close()
    conn.close()

    return archive_list


# Obtain archive volume.
def retrieve_volume(db_name):
    """
    This function will obtain archive volume form SQL server.
    """

    # Define data base connection parameters.
    sqlserver = 'SQL1'
    database = db_name
    username = 'admin'
    password = 'Bosnia66s'

    # Establish DB connections.
    conn = pyodbc.connect(
        r'DRIVER={SQL Server};'
        r'SERVER='+sqlserver+';'
        r'DATABASE='+database+';'
        r'UID='+username+';'
        r'PWD='+password+''
        )
    cur = conn.cursor()
    # Execute query on Data Base.
    cur.execute("""
                select datepart(day,datestart) as day, case command
                when 16385 then 'Retrieve'
                end as "command", COUNT(distinct studyuid) as StudyCount from tblAuditTrailDICOM
                where DateStart>dateadd(day,-30,getdate())
                and CompletionCode=0 and Command in (16385)
                group by datepart(day,datestart), command
                """)
    rows = cur.fetchall()
    # Close SQL Server Connections.
    cur.close()
    conn.close()
    return rows


# Send email with Report
def send_email(file_attachment):
    """This function will send email with the attachment.
    It takes attachment file name as argument.
    """

    # Define email body
    body = "This is Average Retrieve volume by ASP Customer for last 30 days."
    content = MIMEText(body, 'plain')

    # Open file attachment
    filename = file_attachment
    infile = open(filename, "rb")

    # Set up attachment to be send in email
    part = MIMEBase("application", "octet-stream")
    part.set_payload(infile.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=filename)

    msg = MIMEMultipart("alternative")

    # Define email recipients
    to_email = [
        "nerminkekic@ge.com", "aspmonitoring@ge.com"
        ]
    # Define From email
    from_email = "aspmonitoring@ge.com"

    # Create email content
    msg["Subject"] = "ASP Monthly Report {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    msg["From"] = from_email
    msg["To"] = ",".format(to_email)
    msg.attach(part)
    msg.attach(content)
    # Send email to SMTP server
    s = smtplib.SMTP("10.4.1.1", 25)
    s.sendmail(from_email, to_email, msg.as_string())
    s.close()

# Run script
archive_retrieve_report()
