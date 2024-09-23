import paramiko
import pandas as pd
#connection
uname = ''
pword = ''
host = ""

session = paramiko.SSHClient()
session.set_missing_host_key_policy(paramiko.AutoAddPolicy)

session.connect(hostname=host, 
                username=uname,
                password=pword,
                allow_agent=False,
                look_for_keys=False)

sftp = session.open_sftp()


def getFile(carrier_list):
    df = pd.read_csv(sftp.open(carrier_list))

    return df


def uploadFile(new_file_path,df):

    with sftp.open(new_file_path, 'w') as f:
        df.to_csv(f, index=False)

