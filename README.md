# Google Appscripts

### Approval and Reject with Email Notification — approval-reject-sheets.md

This google script is an Approval and Reject with email notifications. Special thanks [Maksim Rogov](http://www.nullriver.com) cleaning up my code.

**Note**: You need to change the column numbers for Email Message, Email of the user, and Email notification "Sent" post after email is sent.
	
- Create a **Sheet** with column **2** enter **Yes** or **No** in the last row. Move row down to bottom to approve. 
- **Yes** moves the row to **APPROVED** sheet
- **No** moves the row to **REJECTED**
- Then it will email notification to user (you must collect user email to do this)
- Once the email sent, "Sent" will be posted in the colummn 23

### Sorting by date by SheetName — sorting-by-date-bysheetname.md

Sort rows by column assign - Make sure you set Timer on your screen onOpen...

- Date should be column 1 (You can change to which column you want to sort)
- tableRange rows content
- That is it! simple...


