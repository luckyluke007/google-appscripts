# Google Appscripts

### approval-reject-sheets.md
**Note**: You need to change the column numbers for Email Message, Email of the user, and Email notification "Sent" post after email is sent.

This google script is an approving process with email notification. 
	- Sheet 1 - Column 2 enter **Yes** or **No** in the last row. Move row down to bottom to approve. 
	- **Yes** moves the row to **APPROVED** sheet
	- **No** moves the row to **REJECTED**
	- Then it will email notification to user (you must collect user email to do this)
	- Once the email sent, Sent will be posted in the colummn 23

- Approval and Reject with email notifications.
   - Thanks [Maksim Rogov](http://www.nullriver.com) cleaning up my code.

### Sorting by date - by SheetName

Sort rows by column assign - Make sure you set Timer on your screen onOpen...

- Date should be column 1 (You can change to which column you want to sort)
- tableRange rows content
- That is it! simple...


