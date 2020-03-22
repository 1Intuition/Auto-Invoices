import os, sys, requests, hashlib

DOWNLOAD_LINK = "https://raw.githubusercontent.com/1Intuition/Auto-Invoices/master/auto-invoices.py"
LOCAL_FILE_PATH = "auto-invoices.py"

def sha256sum(filename):
    h  = hashlib.sha256()
    mv = memoryview(bytearray(128*1024))
    with open(filename, 'rb', buffering=0) as f:
        for n in iter(lambda : f.readinto(mv), 0):
            h.update(mv[:n])
    return h.hexdigest()

def update(req):
    with open(LOCAL_FILE_PATH,'wb') as f:
        f.write(req.content)
    print("Update complete!")

def checkForUpdates():
    try:
        req = requests.get(DOWNLOAD_LINK)
    except Exception:
        print("No internet connexion. Cannot install...")
        return
    if os.access(LOCAL_FILE_PATH, os.F_OK) is not True:
        print("Installing program...")
        update(req)
        return
    if "\n" not in req.text:
        print("Corrupt online file!")
        return
    fl_online = req.text[0:req.text.index("\n")]
    with open(LOCAL_FILE_PATH) as f:
        fl_local = f.readline().strip()
    if fl_local == fl_online:
        if hashlib.sha256(req.content).hexdigest() == sha256sum(LOCAL_FILE_PATH):
            print("Latest file! No need to update.")
        else:
            print("Corrupt file! Repairing file...")
            update(req)
    else:
        print("You are not running the latest version. Need to update!")
        update(req)

if __name__ == '__main__':
    checkForUpdates()
    exec(open(LOCAL_FILE_PATH).read())
