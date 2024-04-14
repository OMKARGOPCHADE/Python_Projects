#import cgi
import qrcode

# Create instance of FieldStorage
#form = cgi.FieldStorage()

# Get data from form
#data = form.getvalue('python')

def qrCode(filename):
        data = "https://docs.google.com/spreadsheets/d/1M84OMIr51Tok72jqioo7g6FoWWElGrgSE0UwDo5qXjs/edit?usp=sharing"
        # Generate QR code
        img = qrcode.make(data)
        
        # Save the QR code image as a file
        img.save(filename)
filename = "Student_IT.png"
qrCode(filename)

