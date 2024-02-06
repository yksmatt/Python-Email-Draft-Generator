# Import necessary packages
import csv
# import win32com for calling outlook application
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
# import os for the function of adding attachment
import os

# Change "YOUR COMPANY NAME" to suit your case
emailSignature = 'Regards,<br>YOUR COMPANY NAME'
counter = 0


# Change "YOUR PATH" to suit your case
with open(r'C:\YOUR PATH\email_list.csv') as file_obj:
      
    # Skips the heading
    # Using next() method
    heading = next(file_obj)
      
    # Create reader object by passing the file 
    # object to reader method
    reader_obj = csv.reader(file_obj)
      
    # Iterate over each row in the csv file 
    # using reader object      
    for row in reader_obj:
        if row[1].lower() == 'yes':
            mail = outlook.CreateItem(0)
            mail.Subject = row[3]
            mail.To = row[2]
            bodyString = row[4] + ',<br><br>' + row[5] + '<br><br>' + emailSignature
            mail.HTMLBody = bodyString
            mail.Save()
            counter = counter + 1
    print('Drafts created: ' + str(counter))
    
