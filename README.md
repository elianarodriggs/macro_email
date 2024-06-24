# Email Automation for Return Process

This Python script automates the process of sending emails related to return authorizations. It reads data from a CSV file, composes the emails, and sends them using the Gmail SMTP server.

# Prerequisites

Before running this script, ensure you have the following installed:

  • Python 3.x
  • Pandas
  • smtplib (included in Python's standard library)
  • email (included in Python's standard library)
  • google.colab (if running in Google Colab) (can use in another IDE too)

# Libraries Used

• pandas: For reading and manipulating CSV files.
• time, datetime: For tracking and printing timestamps.
• email: For constructing email messages and handling attachments.
• smtplib: For sending emails via SMTP.
• google.colab: For mounting Google Drive in Google Colab.

# How to Use

# Step 1: Import Libraries
The script starts by importing the necessary libraries.

```

import pandas as pd
import time
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.colab import drive

```
# Step 2: Mount Google Drive
If you are running the script in Google Colab, mount your Google Drive to access files.

```
drive.mount('/content/gdrive')

```

# Step 3: Read Transporters and Emails List
The script reads the list of transporters and emails from a CSV file.

```
transportadoras = pd.read_csv('/content/gdrive/Shareddrives/.../to_send.csv', sep=';')

```
# Step 4: Configure Email Server
Enter your Gmail credentials to log in to the SMTP server.

```
your_email = input('Enter your gmail: ')
your_password = input('Enter your password gmail: ')

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(your_email, your_password)

```

# Step 5: Read and Process Email List
Read the email list and extract relevant information.

```
email_list = pd.read_csv('/content/gdrive/Shareddrives/.../to_send.csv', sep=';')
email_list['email'] = email_list['Lista Emails'].astype(str)
email_list['cc'] = email_list['cc'].astype('string')

```

# Step 6: Send Emails
The script loops through the email list and sends emails with attachments.

```
for i in range(len(emails)):
    msg = MIMEMultipart()
    ...
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(your_email, your_password)
    s.sendmail(your_email, email.split(",") + cc.split(";"), msg.as_string())
    s.quit()
```

# Step 7: Close SMTP Server
Finally, close the SMTP server.

```
server.close()

```

# Email Body
The body of the email includes instructions and a link to a shared sheet for further actions.

```
filename1 = f'/content/gdrive/Shareddrives/.../Perdas_{assunto}_{today}.xlsx'
filename2 = f'/content/gdrive/Shareddrives/.../Perdas_{assunto}_{today}.pdf'
...
msg.attach(p1)
msg.attach(p2)
```
# Attachments
The script attaches two files (.xlsx and .pdf) to each email before sending.

```
filename1 = f'/content/gdrive/Shareddrives/.../Perdas_{assunto}_{today}.xlsx'
filename2 = f'/content/gdrive/Shareddrives/.../Perdas_{assunto}_{today}.pdf'
...
msg.attach(p1)
msg.attach(p2)
```

# Notes
Ensure that your Gmail account has "Less secure app access" enabled to use SMTP.
Update the paths in the script to match your directory structure in Google Drive.
The script handles errors such as missing files and invalid email addresses.

#Conclusion
This script automates the tedious process of sending multiple emails with attachments, saving time and reducing manual effort. Ensure you have the correct permissions and paths set up before running the script. If you encounter any issues or have questions, feel free to reach out for support.
