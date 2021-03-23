import pandas as pd
import smtplib 
import string
import uuid
from email.mime.multipart import MIMEMultipart 
from email.mime.text import MIMEText 
from email.mime.base import MIMEBase 
from email import encoders 

from PIL import Image, ImageDraw, ImageFont
from pandas import ExcelWriter
from pandas import ExcelFile

password = "vit1234$" #Your Password
fromaddr = "anjali.deore19@vit.edu" #Your Email
df = pd.read_excel('sample-excel-sheet.xlsx')   # Your excel sheet with Participants Name and Email
print(df.index)

for i in df.index:
	image = Image.open('certificates/certificate.jpg')           #Your certficate template
	draw = ImageDraw.Draw(image)
	font = ImageFont.truetype('fonts/GoogleSans-Regular.ttf', size=200)
	font1 = ImageFont.truetype('fonts/GoogleSans-Regular.ttf', size=80)
	color = 'rgb(255,255,255)'
	color_name='#323232'
	color_no='rgb(0,153,255)'
	name = df['Name'][i]
	print( df['Name'][i])

	print(type(name))
	print(str(name))
	name.upper()
	print(i+1,name)
	#lenght1=len(str(i+1))
	length=len(str(name))
	z=11728/2-(length*50)/2

	y=12880/2-(length*50)/2
	x=7000/2-(length*200)/2   #Here I have calculated the coordinates of the name to be placed in the certificate assuming each alphabet taking 200 pixel and the certificate width is 5000 pixel
	print(x)

	draw.text((x, 2500), str(name), fill=color_name, font=font)
	draw.text((y, 4320), str(uuid.uuid1()), fill=color_no, font=font1)	
	draw.text((z, 4320), " Certificate  No:", fill=color, font=font1)	
	imageName = "certificates/"+str(name)+".pdf"
	image.save(imageName)
	
	toaddr = df['Email'][i]
	msg = MIMEMultipart() 
	msg['From'] = "anjali.deore19@vit.edu"   #Your Email Id 
	msg['Subject'] = "Certificate of Appreciation for contributing as trainers for MLfest'21."
	body = '''Congratulations! Please find attached your certificate for contributing as trainers in MLfest'21."
			  
	Regards
	DSC '''

	msg.attach(MIMEText(body, 'plain')) 

	filename = name +".pdf"
	attachment = open(imageName, "rb")
	p = MIMEBase('application', 'octet-stream') 

	p.set_payload((attachment).read()) 
	encoders.encode_base64(p) 
	p.add_header('Content-Disposition', "attachment; filename= %s" % filename) 
	msg.attach(p) 

	s = smtplib.SMTP('smtp.gmail.com', 587) 
	s.starttls() 
	s.login(fromaddr, password) 
	text = msg.as_string() 
	s.sendmail(fromaddr, toaddr, text) 

s.quit() 
