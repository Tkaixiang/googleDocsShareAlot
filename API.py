from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from csv import reader 

def oauth():
    gauth = GoogleAuth()

    try:
        gauth.LocalWebserverAuth() # client_secrets.json need to be in the same directory as the script
    except Exception as e:
        print(e)
        return False
    else:
        gauth.SaveCredentialsFile('credentials/credentials.json')
        return gauth 

def loadauth():
    gauth = GoogleAuth()

    try:
        gauth.LoadCredentialsFile('credentials/credentials.json')
    except:
        return False
    else:
        return gauth

def processCSV(file_path): ## param file_path from pyQT file dialog  
    with open(file_path) as csvfile: 
        readcsv = reader(csvfile, delimiter=',')
        enames = {row[0]:row[1] for row in readcsv}

    return enames 

# def driveConnect(gauth):
#     drive = GoogleDrive(gauth) 
#     return drive 

def createFolder(gauth, fname): 
    # Creates a folder in drive to store documents 
    drive = GoogleDrive(gauth) 
    folder = drive.CreateFile({'title' : fname, 'mimeType' : 'application/vnd.google-apps.folder'})
    folder.Upload()

    fid = folder['id']
    return fid 

def upload(gauth, student, fid, path):  ## where student=[name, email]
    drive = GoogleDrive(gauth) 
    filename = student[0] 
    file = drive.CreateFile({'title': filename, 
                             'mimeType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', "parents": [{"kind": "drive#fileLink",
                     "id": fid}]}) ## setting id creates the doc within the specified _temp_folder 

    file.SetContentFile(path)
    file.Upload(param={'convert': True}) #param True converts the word document into Google Document

    # Insert the permission
    permission = file.InsertPermission({
                            'type': 'user',
                            'value': student[1],
                            'role': 'writer',
                            'WithLink':False})
