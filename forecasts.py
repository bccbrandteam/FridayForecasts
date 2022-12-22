import os
import logging
import datetime
import smtplib, ssl
import sys

sys.path.append( '/home/pi/Desktop/CodeFiles/FridayForecasts' )
#import jobs

cwd = os.path.abspath(os.getcwd())
exceptions = os.path.join(cwd, 'Exceptions')

# now = datetime.datetime.now()
# stamp = now.strftime('%Y%m%d_%H%M%S')
# log_name = '{}.log'.format(stamp)
#mpath_to_log = os.path.join(exceptions, log_name)
# logging.basicConfig(level=logging.ERROR, filename=path_to_log)

# Email notification function for exceptions
def emailNotification(pythonFile):
    # SMPT/SSL Constants
    port = 465
    account = 'bcc.notification.noreply@gmail.com'
    password = 'pkkqrzkfwtvlsgqg'
    recipients = {
        'Andrea Gettys' : 'msbee_andrea@byu.edu'
    }
    # Create a Secure SSL context
    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(account, password)
        for i in recipients:

            # Edit this message as needed
            message = 'Dear {i}, \n1% \n1%You have received this message because a file located at \n1%{}\n1%has attempted to run 3 times and has logged an exception within the {} py file.'.format(i, cwd, pythonFile) 
            #message = f
             #       """Dear {i},
              #      
               #     You have received this message because a file 
                #    located at 
                 #   {cwd} 
                  #  has attempted to run 3 times and has logged an exception within
                   # the {pythonFile} py file
                #"""
            # Sending the email (sender, receiver, message)
            server.sendmail(account,recipients[i],message)

try: 
    attempts = 0
    success_record = [False,False,False]
    success_count = 0

    while attempts < 5:
        if success_record[0] == False:
            try:
                print('Running marketing...')
                import marketing
                success_record[0] = True
                print('Successfully finished running marketing!', '\n')
            except Exception:
                print("jobs.py failed")
                if attempts == 4:
                    emailNotification('marketing')
                    logging.exception('MARKETING.PY FAILED TO EXECUTE')

        for success in success_record:
            if success == True:
                success_count += 1
        if success_count == 1:
            break
        else:
            success_count = 0

        attempts += 1
    
    print("\nSuccessfully ran FridayForecasts!\n")
    logging.shutdown()

except:
    print("\nfridayforecasts failed\n")
