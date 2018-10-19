# InterimPeerEvaluation
A peer evaluation process that automatically emails students.

This is a semi-automated process for compiling and emailing peer evaluations. 

1. Setup a Google Form [similar to this one](https://docs.google.com/forms/d/e/1FAIpQLSdxT9KISgskdF4X_iPEPRpamMhmsCsJe0531dnsOhD7rOdELA/viewform?usp=sf_link).  If you modify the order or types of questions asked, the associated script below will need to be modified.

2. Post a copy of the form to your courses, modifying to include the correct student names, which you can copy and paste from the course roster.

3. After the deadline, download a *.csv version of the responses.  You'll need to save the file as an *.xlsx and then add a tab called "Contacts" with 3 columns - First Name, Last Name, and ASURITE for each student.  Here's an example.Preview the document

4. Install [Python 3.x](https://www.python.org/) and [Openpyxl](https://openpyxl.readthedocs.io/en/stable/). Download the Peer Eval Script. 

5.  Make sure the script and excel file are in the same directory.  Edit the script line 
'''
reader = Eval_Reader("FSE100A_Sp18.xlsx") 
'''
with the correct filename.  If you want to do a test, comment out the "reader.send_emails()" line before you execute.  A preview will be written to the test_out.txt file.  

6. When you are ready to send emails, make sure your outlook is open and logged in, then run the script with the 'reader.send_emails()' line included.  
