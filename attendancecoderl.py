import xlrd, smtplib, sys

workbook = xlrd.open_workbook('/Users/muhammadjamal/Documents/attendance.xlsx')

sheet = workbook.sheet_by_name('Sheet1')

late_students = {}

for student in range(1,7):
	count = 0
	
	for day in range(2,7):
		
		attendance = sheet.cell(student,day).value
		
		if attendance == 'L':
			
			count += 1
			
			if count >= 3:
				name = sheet.cell(student, 0).value
				mail = sheet.cell(student, 1).value
				late_students[name] = mail

print (late_students)

email_obj = smtplib.SMTP('smtp-mail.outlook.com', 587)
email_obj.ehlo()
email_obj.starttls()
email_obj.login('email', 'password')

for kid in late_students:
	message = 'Subject: {}\n\n{}'.format('URGENT REMINDER FOR GUARDIAN OF %s' % (kid), 'Dear guardian, I would like to inform you that %s has been late for 3 or more out of the 5 classes he had this week' % (kid))
	#body = 'Subject: URGENT REMINDER FOR GUARDIAN OF %s.\nDear guardian, \n I would like to inform you that %s has been late for 3 or more out of the 5 classes he had this week' %(kid, kid) 
	print ('sending email to %s' % late_students[kid])
	mailstatus = email_obj.sendmail('email', late_students[kid], body)
	if mailstatus != {}:
		print ('there was a problem sending the mail to %s' % late_students[kid])
	else:
		print ('you have succesfull sent the mails to the appropriate Guardian.')
email_obj.quit()


