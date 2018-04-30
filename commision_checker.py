import xlrd
import smtplib
import sys

workbook = xlrd.open_workbook('/Users/User/Desktop/file.xlsx')

sheet = workbook.sheet_by_name('sheetname')

declining_employees = {}
slacking_employees = {}


for employee in range(1,6):

	neg_counter = 0
	
	for month in range (2,13):

		monthly_dif = (sheet.cell(employee, month + 1).value) - (sheet.cell(employee,month).value)
		if monthly_dif < 0:
			neg_counter += 1
		if neg_counter >= 3:
			name = sheet.cell(employee, 0).value
			mail = sheet.cell(employee,1).value
			declining_employees[name] = mail

for employee in range(1,6):
	
	for month in range(2,14):

		revenue = sheet.cell(employee, month).value
		if revenue < 100:
			slack_name = sheet.cell(employee,0).value
			slack_mail = sheet.cell(employee,1).value
			slacking_employees[slack_name] = slack_mail

print (declining_employees)
print (slacking_employees)

email_obj = smtplib.SMTP('smtp-mail.outlook.com', 587)
email_obj.ehlo()
email_obj.starttls()
email_obj.login('email', 'password')

for employee in slacking_employees:

	message1 = 'Subject: {}\n\n{}'.format('URGENT NOTICE FOR %s' % (employee), 'Dear %s, \n I would like to inform you that your revenue generated does not uphold to our companies standards, for this reason it is with deep regret that I must place you on a temporary leave'  % (employee))
	
	#body_1 = 'Subject: URGENT NOTICE FOR %s.\nDear %s, \n I would like to inform you that your revenue generated does not uphold to our companies standards, for this reason it is with deep regret that I must place you on a temporary leave' %(employee, employee) 
	
	print ('sending email to %s' % slacking_employees[employee])
	
	mailstatus = email_obj.sendmail('m.m.adil@hotmail.com', slacking_employees[employee], message1)
	if mailstatus != {}:
		print ('there was a problem sending the mail to %s' % late_students[kid])
	else:
		print ('you have succesfull sent the mails to the appropriate individual.')
#email_obj.quit()


for worker in declining_employees:
	message2 = 'Subject: {}\n\n{}'.format('Reminder for %s.' % (worker), 'Dear %s I would like to inform you that this year, 3 times, your sales were lower than the previous month, please try to go at a steady pace and be more consistent, thanks'  % (worker))

	#body_2 = 'Subject: Reminder for %s.\nDear %s I would like to inform you that this year, 3 times, your sales were lower than the previous month, please try to go at a steady pace and be more consistent, thank you' % (worker, worker) 
	
	print ('sending email to %s' % declining_employees[worker])
	mailstatus = email_obj.sendmail('email',declining_employees[worker], message2)
	if mailstatus != {}:
		print ('there was a problem sending the mail to %s' % late_students[kid])
	else:
		print ('you have succesfull sent the mails to the appropriate individual.')
#email_obj.quit()









