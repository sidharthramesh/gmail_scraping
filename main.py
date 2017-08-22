import imaplib, email
import re
import pandas as pd
import sys

def get_entries():
    m = imaplib.IMAP4_SSL('imap.gmail.com')
    m.login('igclareg@gmail.com','igclareg17')
    m.select('INBOX')

    igcla_mails = m.search(None,'SUBJECT IGCLA')
    igcla_mails = igcla_mails[1][0].decode().split(' ')
    igcla_mails_len = len(igcla_mails)
    print("Found {} emails with subject IGCLA".format(igcla_mails_len))
    pattern = re.compile(r'<p>([\w\s\d\,]+)\s*:\s*([\w\s\d@.]+)')

    all_data = []
    for n,i in enumerate(igcla_mails):
        typ, data = m.fetch(i, '(RFC822)')
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_string(response_part[1].decode('utf-8'))
                payload = msg.get_payload()
                #print(payload)
                a = payload.replace('=0D=0A','').replace('=','').replace('\r\n','')
                d = dict(pattern.findall(a))
                progress = int((n/igcla_mails_len) * 100)
                sys.stdout.write("Progress: {}% - Completed {} emails\r".format(progress, n + 1))
                
                sys.stdout.flush()
                #print(d)
                all_data.append(d)
    return all_data
data = get_entries()
print("Writing to file igcla_registration.csv")
df = pd.DataFrame(data)
df.to_csv('igcla_registration.csv',index=False)
print("Done")




